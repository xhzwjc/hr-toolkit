from __future__ import annotations

import hashlib
import os
import re
import shutil
import threading
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..common.excel_compat import ensure_xlsx_workbook, is_supported_excel_file
from ..common.filenames import safe_filename

# OCR 引擎全局单例与线程安全锁
_OCR_ENGINE = None
_OCR_ATTEMPTED = False
_OCR_LOCK = threading.Lock()


def _get_ocr_engine():
    global _OCR_ENGINE, _OCR_ATTEMPTED
    with _OCR_LOCK:
        if not _OCR_ATTEMPTED:
            _OCR_ATTEMPTED = True
            try:
                from rapidocr_onnxruntime import RapidOCR
                _OCR_ENGINE = RapidOCR()
            except Exception:
                _OCR_ENGINE = None
        return _OCR_ENGINE


TOOL_NAME = "需求9-员工资料自动打包与信息提取"

MODE_BY_EMPLOYEE = "by_employee"
MODE_BY_MATERIAL = "by_material"
MODE_FLAT = "flat"
MODES = {MODE_BY_EMPLOYEE, MODE_BY_MATERIAL, MODE_FLAT}

MODE_LABELS = {
    "按员工归类（每人一个文件夹）": MODE_BY_EMPLOYEE,
    "按材料归类（每类材料一个文件夹）": MODE_BY_MATERIAL,
    "平铺输出（所有文件在同一文件夹）": MODE_FLAT,
}
MODE_LABELS_REVERSE = {v: k for k, v in MODE_LABELS.items()}

# 预置常见材料类型及其多维度别名同义词库（互斥排他分类，绝不串门）
MATERIAL_SYNONYMS: dict[str, list[str]] = {
    "身份证": [
        "身份证", "身分证", "sfz", "idcard", "id_card", "id", "identity",
        "正面", "反面", "人像面", "国徽面", "人像", "国徽", "A面", "B面", "正反面", "正反",
        "zhengmian", "fanmian", "zm", "fm",
        "身份证正面", "身份证反面", "身份证正反面", "身份证复印件", "身份证照片", "身份证件", "证件",
    ],
    "护照": [
        "护照", "passport", "通行证", "港澳通行证", "台湾通行证", "国际旅行证件",
    ],
    "劳动合同": [
        "劳动合同", "劳动协议", "劳务合同", "劳务协议", "用工合同", "用工协议", "聘用合同", "聘用协议",
        "续签合同", "续签协议", "合同", "协议", "contract", "hetong", "ht", "劳动关系", "劳动手册", "协议书", "聘书", "用工",
    ],
    "学历证明": [
        "学历证", "毕业证", "学位证", "学历证书", "毕业证书", "学位证书", "学信网", "备案表", "教育部",
        "学历证明", "学历认证", "学历", "毕业", "学位", "文凭", "xueli", "biye", "xuexin",
        "学籍", "大专", "本科", "硕士", "博士", "中专", "高中",
    ],
    "资格证书": [
        "资格证", "职业资格证", "职称证", "技能证", "操作证", "登高证", "电工证", "焊工证", "安全员",
        "特种作业", "驾驶证", "驾照", "上岗证", "从业资格", "资格", "职称", "技能", "证书", "zige", "jineng",
        "certificate", "license",
    ],
    "体检报告": [
        "体检", "体检报告", "体检表", "入职体检", "健康体检", "健康证", "健康检查", "健康证明",
        "报告单", "tijian", "jiankang", "医院", "诊断", "化验单",
    ],
    "证件照片": [
        "一寸照", "二寸照", "证件照", "寸照", "登记照", "蓝底", "白底", "红底", "个人照片", "照片",
        "相片", "头像", "个人照", "photo", "pic", "avatar", "head",
    ],
    "离职证明": [
        "离职证明", "解除劳动合同证明", "解除劳动关系", "解除协议", "离职申请", "退工证明", "解除证明",
        "离职表", "退工单", "离职", "退工", "lizhi", "tuigong",
    ],
    "个人简历": [
        "简历", "个人简历", "履历表", "员工登记表", "入职登记表", "应聘登记表", "花名册",
        "履历", "resume", "cv", "jianli", "基本信息",
    ],
    "银行卡": [
        "银行卡", "工资卡", "卡号", "开户行", "存折", "银行账号", "bank", "card", "yinhang",
    ],
}

_NOISE_WORDS: set[str] = {
    "序号", "姓名", "员工", "人员", "名字", "部门", "项目", "项目部", "所属部门", "归属部门",
    "身份证", "身份证号码", "身份证号", "证件号码", "工号", "员工编号", "职务", "岗位", "状态",
    "备注", "合计", "总计", "花名册", "统计表", "名单", "汇总", "公司", "集团", "制表",
    "日期", "未命名", "需要材料", "材料", "需求", "所需材料", "资料类型", "全部", "正面", "反面",
}

_ALL_MATERIAL_KEYWORDS: set[str] = set()
for _syns in MATERIAL_SYNONYMS.values():
    for _s in _syns:
        _ALL_MATERIAL_KEYWORDS.add(_s)
for _k in MATERIAL_SYNONYMS:
    _ALL_MATERIAL_KEYWORDS.add(_k)

SUPPORTED_FILE_EXTENSIONS = {
    ".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp",
    ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt",
}

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}
DOC_EXTENSIONS = {".doc", ".docx", ".pdf", ".txt"}

_IGNORED_FILENAMES = {
    ".ds_store", "thumbs.db", "desktop.ini", ".localized", "ehthumbs.db",
}

_PHONE_RE = re.compile(r"^1[3-9]\d{9}$")
_ID_CARD_RE = re.compile(r"^\d{17}[\dXx]$|^\d{15}$")


def _is_junk_or_temp_file(path: Path | str) -> bool:
    """检查是否为系统垃圾文件或 Office 临时锁文件。"""
    name = Path(path).name.lower()
    if name.startswith(".") or name.startswith("~$") or name in _IGNORED_FILENAMES:
        return True
    return False


def _is_path_nested(child: Path, parent: Path) -> bool:
    """检查 child 路径是否处于 parent 路径内部或为同一路径。"""
    try:
        child_res = child.resolve()
        parent_res = parent.resolve()
        if child_res == parent_res:
            return True
        child_res.relative_to(parent_res)
        return True
    except (ValueError, RuntimeError):
        return False


def _get_file_signature(path: Path) -> tuple[int, str]:
    """计算文件大小和前 64KB 哈希，作为同员工去重特征。"""
    try:
        size = path.stat().st_size
        with open(path, "rb") as f:
            chunk = f.read(65536)
        h = hashlib.md5(chunk).hexdigest()
        return size, h
    except Exception:
        return -1, str(path)


@dataclass(frozen=True)
class TargetEmployee:
    name: str
    id_card: str = ""
    employee_no: str = ""
    department: str = ""
    phone: str = ""
    per_person_materials: tuple[str, ...] = ()
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def identity_key(self) -> str:
        """员工唯一复合主键，防止同名员工覆盖。"""
        return f"{self.name}_{self.id_card}_{self.phone}"

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "id_card": self.id_card,
            "employee_no": self.employee_no,
            "department": self.department,
            "phone": self.phone,
            "per_person_materials": list(self.per_person_materials),
        }


@dataclass(frozen=True)
class MaterialFileMatch:
    employee_name: str
    material_type: str
    source_path: Path
    relative_source_path: str
    matched_by: str  # "filename", "ocr", "doc_content", "id_card", "phone", "read_failed"
    target_filename: str = ""
    target_path: Path | None = None
    extracted_person_name: str = ""
    extracted_id_card: str = ""
    mismatch_warning: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "employee_name": self.employee_name,
            "material_type": self.material_type,
            "source_path": str(self.source_path),
            "relative_source_path": self.relative_source_path,
            "matched_by": self.matched_by,
            "target_filename": self.target_filename,
            "target_path": str(self.target_path) if self.target_path else "",
            "extracted_person_name": self.extracted_person_name,
            "extracted_id_card": self.extracted_id_card,
            "mismatch_warning": self.mismatch_warning,
        }


@dataclass
class MaterialCollectResult:
    library_dir: Path
    output_dir: Path
    zip_path: Path | None = None
    report_path: Path | None = None
    mode: str = MODE_BY_EMPLOYEE
    target_employees: list[TargetEmployee] = field(default_factory=list)
    requested_materials: list[str] = field(default_factory=list)
    matches: list[MaterialFileMatch] = field(default_factory=list)
    missing_records: dict[str, list[str]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    folder_match_counts: dict[str, int] = field(default_factory=dict)

    @property
    def total_employees(self) -> int:
        return len(self.target_employees)

    @property
    def matched_file_count(self) -> int:
        return len(self.matches)

    @property
    def complete_employee_count(self) -> int:
        return sum(1 for emp in self.target_employees if not self.missing_records.get(emp.name))

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": TOOL_NAME,
            "library_dir": str(self.library_dir),
            "output_dir": str(self.output_dir),
            "zip_path": str(self.zip_path) if self.zip_path else None,
            "report_path": str(self.report_path) if self.report_path else None,
            "mode": self.mode,
            "total_employees": self.total_employees,
            "matched_file_count": self.matched_file_count,
            "complete_employee_count": self.complete_employee_count,
            "requested_materials": self.requested_materials,
            "matches": [m.to_dict() for m in self.matches],
            "missing_records": self.missing_records,
            "warnings": self.warnings,
        }


# ---------------------------------------------------------------------------
# 材料需求识别：将自由文本映射到标准材料类型
# ---------------------------------------------------------------------------

def _resolve_material_text(text: str) -> list[str]:
    """将自由文本（如"身份证，合同"）解析为标准材料类型列表。"""
    if not text:
        return []
    result: list[str] = []
    parts = re.split(r"[,，、;；\s]+", text.strip())
    for part in parts:
        part = part.strip()
        if not part:
            continue
        matched = False
        for mat_type, synonyms in MATERIAL_SYNONYMS.items():
            if part == mat_type or any(syn == part or syn in part for syn in synonyms):
                if mat_type not in result:
                    result.append(mat_type)
                matched = True
                break
        if not matched and part not in result:
            result.append(part)
    return result


def _is_valid_person_name(name: str) -> bool:
    """校验是否为一个合法的员工姓名（过滤数字、表头、噪音词）。"""
    name = str(name or "").strip()
    if not name:
        return False
    if name in _NOISE_WORDS:
        return False
    if re.match(r"^\d+$", name):
        return False
    if any(noise in name for noise in ("花名册", "统计表", "汇总表", "总人数", "部门：", "公司：", "制表人")):
        return False
    if len(name) > 30:
        return False
    return True


# ---------------------------------------------------------------------------
# 目标员工名单解析（支持直接输入单人/多人文本、或 Excel 文件）
# ---------------------------------------------------------------------------

def _parse_single_text_item(item_str: str) -> TargetEmployee | None:
    """解析单个字符串条目（如 '张三', '张三 440111199001011234', '440111199001011234', '张三 身份证'）。"""
    item_str = item_str.strip()
    if not item_str or item_str.startswith("#"):
        return None

    parts = [p for p in re.split(r"[\s\t]+", item_str) if p]
    if not parts:
        return None

    if _ID_CARD_RE.match(parts[0]):
        id_card = parts[0]
        name = parts[1] if len(parts) > 1 and _is_valid_person_name(parts[1]) else id_card
        mat_text = " ".join(parts[2:]) if len(parts) > 2 else ""
        per_mats = tuple(_resolve_material_text(mat_text))
        return TargetEmployee(name=name, id_card=id_card, per_person_materials=per_mats)

    if _PHONE_RE.match(parts[0]):
        phone = parts[0]
        name = parts[1] if len(parts) > 1 and _is_valid_person_name(parts[1]) else phone
        return TargetEmployee(name=name, phone=phone)

    name = parts[0]
    if not _is_valid_person_name(name):
        return None

    id_card = ""
    phone = ""
    emp_no = ""
    materials_parts: list[str] = []

    for p in parts[1:]:
        if _ID_CARD_RE.match(p) and not id_card:
            id_card = p
        elif _PHONE_RE.match(p) and not phone:
            phone = p
        elif any(kw in p for kw in _ALL_MATERIAL_KEYWORDS):
            materials_parts.append(p)
        elif re.match(r"^[A-Za-z0-9_-]+$", p) and len(p) <= 10 and not emp_no:
            emp_no = p
        else:
            materials_parts.append(p)

    per_mats = tuple(_resolve_material_text(" ".join(materials_parts)))
    return TargetEmployee(
        name=name,
        id_card=id_card,
        phone=phone,
        employee_no=emp_no,
        per_person_materials=per_mats,
    )


def parse_employee_roster(
    source: str | Path | list[dict[str, Any]] | list[str],
) -> list[TargetEmployee]:
    """Parse employee list from an Excel workbook, text content, or structured list."""
    employees: list[TargetEmployee] = []
    seen_keys: set[str] = set()

    def _add_emp(emp: TargetEmployee | None) -> None:
        if not emp or not emp.name:
            return
        key = emp.identity_key
        if key not in seen_keys:
            seen_keys.add(key)
            employees.append(emp)

    if isinstance(source, list):
        for item in source:
            if isinstance(item, dict):
                name = str(item.get("name") or item.get("姓名") or "").strip()
                id_card = str(item.get("id_card") or item.get("身份证号码") or item.get("身份证") or "").strip()
                emp_no = str(item.get("employee_no") or item.get("工号") or "").strip()
                dept = str(item.get("department") or item.get("部门") or item.get("项目") or "").strip()
                phone = str(item.get("phone") or item.get("手机号") or item.get("电话") or "").strip()
                mat_text = str(item.get("materials") or item.get("材料") or item.get("需要材料") or "").strip()
                per_mats = tuple(_resolve_material_text(mat_text))
                if _is_valid_person_name(name) or _ID_CARD_RE.match(id_card):
                    _add_emp(TargetEmployee(
                        name=name or id_card, id_card=id_card, employee_no=emp_no,
                        department=dept, phone=phone, per_person_materials=per_mats,
                    ))
            elif isinstance(item, str):
                _add_emp(_parse_single_text_item(item))
        return employees

    source_path = Path(source) if isinstance(source, (str, Path)) else None
    if source_path and source_path.is_file() and is_supported_excel_file(source_path):
        import tempfile
        with tempfile.TemporaryDirectory() as temp_dir_str:
            working_path = ensure_xlsx_workbook(source_path, Path(temp_dir_str))
            wb = load_workbook(working_path, data_only=True)
            ws = wb.active

            name_col: int | None = None
            id_card_col: int | None = None
            emp_no_col: int | None = None
            dept_col: int | None = None
            phone_col: int | None = None
            material_col: int | None = None
            header_row_idx = 1

            name_synonyms = ("姓名", "员工姓名", "人员姓名", "名字")
            id_card_synonyms = ("身份证号码", "身份证号", "身份证", "证件号码", "证件号")
            emp_no_synonyms = ("工号", "员工编号", "员工号", "人员编号")
            dept_synonyms = ("部门", "项目", "项目部", "所属部门", "归属部门", "所属项目")
            phone_synonyms = ("手机号", "手机号码", "电话", "联系电话", "联系方式")
            material_col_synonyms = ("需要材料", "材料", "需求", "所需材料", "资料类型")

            max_r = ws.max_row or 1
            max_c = ws.max_column or 1

            for r in range(1, min(max_r, 15) + 1):
                for c in range(1, max_c + 1):
                    val = str(ws.cell(r, c).value or "").strip()
                    if not val:
                        continue
                    if name_col is None and val in name_synonyms:
                        name_col = c
                    if id_card_col is None and val in id_card_synonyms:
                        id_card_col = c
                    if emp_no_col is None and val in emp_no_synonyms:
                        emp_no_col = c
                    if dept_col is None and val in dept_synonyms:
                        dept_col = c
                    if phone_col is None and val in phone_synonyms:
                        phone_col = c
                    if material_col is None and any(syn in val for syn in material_col_synonyms):
                        material_col = c
                if name_col is not None:
                    header_row_idx = r
                    break

            if name_col is None:
                first_val = str(ws.cell(1, 1).value or "").strip()
                if _is_valid_person_name(first_val) and not any(noise in first_val for noise in ("花名册", "表", "单", "人员", "员工")):
                    name_col = 1
                    header_row_idx = 0
                else:
                    name_col = 1
                    header_row_idx = 1

            if material_col is None and max_c >= name_col + 1:
                candidate_col = name_col + 1
                if candidate_col not in (id_card_col, emp_no_col, dept_col, phone_col):
                    hits = 0
                    checked = 0
                    for r in range(header_row_idx + 1, min(max_r, header_row_idx + 20) + 1):
                        cell_val = str(ws.cell(r, candidate_col).value or "").strip()
                        if not cell_val:
                            continue
                        checked += 1
                        if any(kw in cell_val for kw in _ALL_MATERIAL_KEYWORDS):
                            hits += 1
                    if checked > 0 and hits / checked >= 0.3:
                        material_col = candidate_col

            for r in range(header_row_idx + 1, max_r + 1):
                name_val = str(ws.cell(r, name_col).value or "").strip()
                if not name_val or not _is_valid_person_name(name_val):
                    continue
                id_card_val = str(ws.cell(r, id_card_col).value or "").strip() if id_card_col else ""
                emp_no_val = str(ws.cell(r, emp_no_col).value or "").strip() if emp_no_col else ""
                dept_val = str(ws.cell(r, dept_col).value or "").strip() if dept_col else ""
                phone_val = str(ws.cell(r, phone_col).value or "").strip() if phone_col else ""

                mat_text = ""
                if material_col:
                    mat_text = str(ws.cell(r, material_col).value or "").strip()
                per_mats = tuple(_resolve_material_text(mat_text))

                _add_emp(TargetEmployee(
                    name=name_val,
                    id_card=id_card_val,
                    phone=phone_val,
                    employee_no=emp_no_val,
                    department=dept_val,
                    per_person_materials=per_mats,
                ))
            wb.close()
        return employees

    raw_text = str(source)
    raw_items = re.split(r"[\n\r;；,，]+", raw_text)
    for item in raw_items:
        _add_emp(_parse_single_text_item(item))

    return employees


# ---------------------------------------------------------------------------
# 文件夹匹配核心算法
# ---------------------------------------------------------------------------

def _match_folder_to_employee(folder_name: str, emp: TargetEmployee) -> str | None:
    """判断一个文件夹名是否属于某员工，返回匹配依据或 None。"""
    f_name = folder_name.strip()
    if not f_name:
        return None

    emp_id = emp.id_card.strip()
    if emp_id and len(emp_id) >= 15 and emp_id in f_name:
        return "id_card"

    emp_phone = emp.phone.strip()
    if emp_phone and len(emp_phone) == 11 and emp_phone in f_name:
        return "phone"

    emp_name = emp.name.strip()
    if not emp_name or not _is_valid_person_name(emp_name):
        return None

    if _ID_CARD_RE.match(emp_name) and emp_name in f_name:
        return "id_card"

    if f_name == emp_name:
        return "exact_name"

    pattern = rf"(?:^|[\d_\s\-\(\)（）\[\]【】#])" + re.escape(emp_name) + rf"(?:[\d_\s\-\(\)（）\[\]【】#]|$)"
    if re.search(pattern, f_name):
        return "name"

    if len(emp_name) >= 2 and emp_name in f_name:
        return "name_sub"

    return None


# ---------------------------------------------------------------------------
# 文档正文提取 & 本地离线 OCR 识图引擎
# ---------------------------------------------------------------------------

def _extract_document_text(file_path: Path) -> str:
    """提取文档内部文本（支持 .docx, .txt, .pdf, .doc 纯文本搜索）。"""
    ext = file_path.suffix.lower()
    if ext == ".docx":
        try:
            with zipfile.ZipFile(file_path) as zf:
                if "word/document.xml" in zf.namelist():
                    xml_content = zf.read("word/document.xml")
                    tree = ET.fromstring(xml_content)
                    return "".join(tree.itertext())
        except Exception:
            pass
    elif ext in (".txt", ".csv"):
        try:
            return file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            pass
    elif ext == ".pdf":
        try:
            with open(file_path, "rb") as f:
                content = f.read(150000)
            texts = re.findall(rb"\((.*?)\)[\s]*Tj", content)
            if texts:
                return b" ".join(texts).decode("utf-8", errors="ignore")
            return content.decode("utf-8", errors="ignore")
        except Exception:
            pass
    elif ext == ".doc":
        try:
            with open(file_path, "rb") as f:
                content = f.read(150000)
            return content.decode("utf-8", errors="ignore")
        except Exception:
            pass
    return ""


_DOC_CONTENT_PATTERNS: dict[str, list[str]] = {
    "劳动合同": ["劳动合同", "用工合同", "劳务合同", "聘用合同", "用工协议", "劳动期限", "工作内容", "劳动报酬", "劳动争议", "解除劳动合同", "劳动法", "甲乙双方根据", "试用期"],
    "个人简历": ["个人简历", "个人履历", "工作经历", "教育经历", "求职意向", "应聘登记", "员工登记表", "基本信息表"],
    "体检报告": ["体检报告", "体格检查", "体检结论", "健康体检", "化验单", "检验报告单", "心电图", "胸透", "体检合格"],
    "学历证明": ["毕业证书", "学位证书", "教育部学历证书", "学信网", "普通高等学校", "学士学位", "硕士学位", "博士学位"],
    "离职证明": ["离职证明", "解除劳动关系证明", "退工证明", "解除劳动合同通知书"],
}


def _classify_by_ocr(file_path: Path) -> tuple[str | None, str, str, str, str]:
    """通过本地离线 OCR 识别图片文字并进行材料分类与信息提取。

    Returns: (material_type or None, match_method, subtype_label, extracted_name, extracted_id)
    例如: ('身份证', 'ocr_id_front', '正面', '姜默蒙', '130527199211020552')
    """
    engine = _get_ocr_engine()
    if engine is None:
        return None, "", "", "", ""

    try:
        with _OCR_LOCK:
            result, _ = engine(str(file_path))
        if not result:
            return None, "", "", "", ""
        texts = [item[1] for item in result]
        full_text = " ".join(texts)
    except Exception:
        return None, "", "", "", ""

    # 实体信息提取：姓名与身份证号
    extracted_name = ""
    extracted_id = ""

    id_match = re.search(r"\b\d{17}[\dxX]\b", full_text)
    if id_match:
        extracted_id = id_match.group(0)

    for idx, t in enumerate(texts):
        m = re.search(r"姓名(?:[/A-Za-z\s:：]*)([\u4e00-\u9fa5]{2,10})", t)
        if m:
            extracted_name = m.group(1).strip()
            break
        if re.search(r"^姓名(?:[/A-Za-z\s:：]*)$", t.strip()):
            if idx + 1 < len(texts) and re.match(r"^[\u4e00-\u9fa5]{2,10}$", texts[idx + 1].strip()):
                extracted_name = texts[idx + 1].strip()
                break

    # 1. 护照（特征极强，优先判断，防止因包含姓名/出生日期被误判为身份证）
    if "PASSPORT" in full_text.upper() or "护照" in full_text or "外交部" in full_text or "国家移民管理局" in full_text:
        return "护照", "ocr_passport", "", extracted_name, extracted_id

    # 2. 身份证反面（国徽面）
    if "居民身份证" in full_text or ("签发机关" in full_text and "有效期限" in full_text):
        return "身份证", "ocr_id_back", "反面", extracted_name, extracted_id

    # 3. 身份证正面（人像面）
    if "公民身份号码" in full_text or ("姓名" in full_text and ("住址" in full_text or "民族" in full_text)) or extracted_id:
        return "身份证", "ocr_id_front", "正面", extracted_name, extracted_id

    # 4. 劳动合同（照片扫描件）
    if "劳动合同" in full_text or "用工合同" in full_text or ("甲方" in full_text and "乙方" in full_text and ("劳动" in full_text or "报酬" in full_text or "工作内容" in full_text)):
        return "劳动合同", "ocr_contract", "", extracted_name, extracted_id

    # 5. 学历证明
    if "毕业证书" in full_text or "学位证书" in full_text or "学信网" in full_text or "学历证书" in full_text or "普通高等学校" in full_text:
        return "学历证明", "ocr_degree", "", extracted_name, extracted_id

    # 6. 体检报告
    if "体检报告" in full_text or "体格检查" in full_text or "检验报告" in full_text or "化验单" in full_text:
        return "体检报告", "ocr_medical", "", extracted_name, extracted_id

    # 7. 资格证书 / 驾驶证
    if "机动车驾驶证" in full_text or "驾驶证" in full_text or "职业资格" in full_text or "特种作业" in full_text:
        return "资格证书", "ocr_certificate", "", extracted_name, extracted_id

    # 8. 银行卡
    if re.search(r"\d{16,19}", full_text) and ("银行" in full_text or "银联" in full_text or "Bank" in full_text):
        return "银行卡", "ocr_bank", "", extracted_name, extracted_id

    # 9. 个人简历
    if "个人简历" in full_text or "个人履历" in full_text or "应聘登记" in full_text:
        return "个人简历", "ocr_resume", "", extracted_name, extracted_id

    return None, "", "", extracted_name, extracted_id


def _classify_material_type(file_path: Path, filename: str, requested_types: list[str]) -> tuple[str | None, str, str, str, str]:
    """Classify file into a material type using filenames, document contents, and local OCR.

    Returns: (matched_material_type or None, match_method, subtype_label, extracted_name, extracted_id)
    """
    stem = Path(filename).stem.lower()

    # 1. 身份证号判断：文件名包含 18 位或 15 位数字
    if re.search(r"\d{17}[\dxX]|\d{15}", stem):
        return "身份证", "id_number_in_filename", "", "", ""

    # 2. 全量同义词库匹配（按优先级互斥判断）
    for mat_type, synonyms in MATERIAL_SYNONYMS.items():
        for syn in synonyms:
            if syn.lower() in stem:
                sub = "正面" if "正面" in stem or "人像" in stem else ("反面" if "反面" in stem or "国徽" in stem else "")
                return mat_type, "filename_keyword", sub, "", ""

    # 3. 文档内部文本内容深度检索（针对文件名如 01.docx, file.pdf 等非标准命名）
    doc_text = _extract_document_text(file_path)
    if doc_text:
        for mat_type, content_keywords in _DOC_CONTENT_PATTERNS.items():
            for kw in content_keywords:
                if kw in doc_text:
                    return mat_type, "doc_content", "", "", ""

    # 4. 本地离线 OCR 视觉图文识别（针对纯哈希/随机命名的图片：如 a5d6e67cd.jpg）
    ext = file_path.suffix.lower()
    if ext in IMAGE_EXTENSIONS:
        ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id = _classify_by_ocr(file_path)
        if ocr_mat:
            return ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id

    return None, "", "", "", ""


def _scan_folder_index(
    lib_path: Path,
    max_depth: int = 1,
    skip_dir: Path | None = None,
) -> dict[str, list[Path]]:
    """扫描资料库，建立"文件夹名 → 文件夹路径列表"的索引，主动跳过输出目录。"""
    folder_index: dict[str, list[Path]] = {}

    def _scan(parent: Path, depth: int) -> None:
        if depth > max_depth:
            return
        if skip_dir and _is_path_nested(parent, skip_dir):
            return
        try:
            entries = list(os.scandir(parent))
        except PermissionError:
            return
        for entry in entries:
            if entry.is_dir(follow_symlinks=False):
                name = entry.name
                if name.startswith(".") or name in _IGNORED_FILENAMES:
                    continue
                path = Path(entry.path)
                if skip_dir and _is_path_nested(path, skip_dir):
                    continue
                folder_index.setdefault(name, []).append(path)
                if depth < max_depth:
                    _scan(path, depth + 1)

    _scan(lib_path, 1)
    return folder_index


# ---------------------------------------------------------------------------
# 核心：收集员工资料
# ---------------------------------------------------------------------------

def collect_employee_materials(
    library_dir: str | Path,
    output_dir: str | Path,
    *,
    roster_source: str | Path | list[dict[str, Any]] | list[str],
    material_types: list[str] | None = None,
    mode: str = MODE_BY_EMPLOYEE,
    create_zip: bool = False,
    generate_report: bool = True,
    collect_all: bool = False,
    scan_depth: int = 1,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> MaterialCollectResult:
    """Search, match, extract, and package employee materials from the repository."""
    lib_path = Path(library_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not lib_path.exists() or not lib_path.is_dir():
        raise FileNotFoundError(f"资料库目录不存在：{lib_path}")

    # P0 防递归死循环：输出目录严禁处于资料库内部
    if _is_path_nested(out_path, lib_path):
        raise ValueError(
            f"保存目录不能在资料库目录内部（会导致循环嵌套复制）：\n"
            f"资料库：{lib_path}\n"
            f"保存目录：{out_path}\n"
            f"请选择一个位于资料库外部的独立文件夹作为保存目录。"
        )

    if mode not in MODES:
        raise ValueError(f"不支持的归类模式：{mode}，可选值：{MODES}")

    employees = parse_employee_roster(roster_source)
    if not employees:
        raise ValueError("未能解析出有效的员工信息，请输入员工姓名/身份证，或上传员工名单表格")

    if material_types is None or len(material_types) == 0:
        global_materials = list(MATERIAL_SYNONYMS.keys())
    else:
        global_materials = list(material_types)

    out_path.mkdir(parents=True, exist_ok=True)

    if progress_callback:
        progress_callback(0, len(employees), "正在扫描资料库文件夹索引...")
    folder_index = _scan_folder_index(lib_path, max_depth=scan_depth, skip_dir=out_path)

    total_steps = len(employees)
    matches: list[MaterialFileMatch] = []
    missing_records: dict[str, list[str]] = {}
    warnings: list[str] = []
    folder_match_counts: dict[str, int] = {}

    for idx, emp in enumerate(employees):
        emp_key = emp.name
        if progress_callback:
            progress_callback(idx + 1, total_steps, f"[{idx + 1}/{total_steps}] 正在检索与匹配：{emp.name}")

        if collect_all:
            emp_materials: list[str] | None = None
        elif emp.per_person_materials:
            emp_materials = list(emp.per_person_materials)
        else:
            emp_materials = global_materials

        matched_folders: list[tuple[Path, str]] = []
        for folder_name, paths in folder_index.items():
            reason = _match_folder_to_employee(folder_name, emp)
            if reason:
                for p in paths:
                    matched_folders.append((p, reason))

        folder_match_counts[emp_key] = len(matched_folders)

        # 同名文件夹防错配检测：如果一个名字匹配到多个不同路径的文件夹
        duplicate_folder_warning = ""
        if len(matched_folders) > 1:
            duplicate_folder_warning = f"⚠️ 资料库中存在 {len(matched_folders)} 个同名文件夹，已全部提取归档，请注意核实！"
            warnings.append(f"员工【{emp.name}】：{duplicate_folder_warning}")

        if not matched_folders:
            if emp_materials is not None:
                missing_records[emp_key] = list(emp_materials) if emp_materials else list(global_materials)
            else:
                missing_records[emp_key] = ["（整个文件夹）"]
            continue

        if emp_materials is None:
            _collect_all_from_folders(
                emp, matched_folders, out_path, mode, matches, warnings, duplicate_folder_warning,
            )
        else:
            emp_missing = _collect_specific_materials(
                emp, matched_folders, out_path, mode, emp_materials, matches, warnings, duplicate_folder_warning,
            )
            if emp_missing:
                missing_records[emp_key] = emp_missing

    zip_path: Path | None = None
    if create_zip:
        zip_path = out_path.parent / f"{out_path.name}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(out_path):
                dirs[:] = [d for d in dirs if not d.startswith(".")]
                for f in files:
                    if _is_junk_or_temp_file(f):
                        continue
                    full_p = Path(root) / f
                    arcname = full_p.relative_to(out_path)
                    zf.write(full_p, arcname=str(arcname))

    report_path: Path | None = None
    if collect_all:
        report_materials = ["全部资料"]
    else:
        report_materials = global_materials
    if generate_report:
        report_path = out_path / "《员工资料提取汇总与缺失清单》.xlsx"
        _write_excel_report(report_path, employees, report_materials, matches, collect_all, warnings)

    return MaterialCollectResult(
        library_dir=lib_path,
        output_dir=out_path,
        zip_path=zip_path,
        report_path=report_path,
        mode=mode,
        target_employees=employees,
        requested_materials=report_materials,
        matches=matches,
        missing_records=missing_records,
        warnings=warnings,
        folder_match_counts=folder_match_counts,
    )


# ---------------------------------------------------------------------------
# 收集策略实现
# ---------------------------------------------------------------------------

def _check_mismatch_warning(emp: TargetEmployee, extracted_name: str, extracted_id: str, duplicate_warning: str = "") -> str:
    """核对识别到的证件人名/号码与目标员工是否一致。"""
    warns: list[str] = []
    if duplicate_warning:
        warns.append(duplicate_warning)
    if extracted_name and emp.name and extracted_name != emp.name and emp.name not in extracted_name:
        warns.append(f"⚠️ 证件姓名【{extracted_name}】与目标【{emp.name}】不一致")
    if extracted_id and emp.id_card and extracted_id != emp.id_card:
        warns.append(f"⚠️ 证件号码【{extracted_id}】与目标【{emp.id_card}】不一致")
    return "；".join(warns)


def _collect_all_from_folders(
    emp: TargetEmployee,
    matched_folders: list[tuple[Path, str]],
    out_path: Path,
    mode: str,
    matches: list[MaterialFileMatch],
    warnings: list[str],
    duplicate_warning: str = "",
) -> None:
    """全部材料模式：将匹配到的文件夹整体拷贝到输出目录。"""
    clean_emp = safe_filename(emp.name)
    seen_hashes: set[tuple[int, str]] = set()

    for folder_idx, (folder_path, match_reason) in enumerate(matched_folders):
        suffix = f"_同名{folder_idx + 1}" if len(matched_folders) > 1 else ""
        dest_name = f"{clean_emp}{suffix}"

        dest_dir = out_path / dest_name
        dest_dir.mkdir(parents=True, exist_ok=True)

        try:
            for root, dirs, files in os.walk(folder_path):
                dirs[:] = [d for d in dirs if not d.startswith(".")]
                rel_root = Path(root).relative_to(folder_path)
                target_root = dest_dir / rel_root
                target_root.mkdir(parents=True, exist_ok=True)

                for f in files:
                    if _is_junk_or_temp_file(f):
                        continue
                    src = Path(root) / f

                    # 同一员工内部跨目录重复文件 Hash 去重
                    sig = _get_file_signature(src)
                    if sig in seen_hashes:
                        continue
                    seen_hashes.add(sig)

                    dst = target_root / f
                    try:
                        shutil.copy2(src, dst)
                    except Exception as e:
                        err_msg = f"复制失败：{src} → {dst}: {e}"
                        warnings.append(err_msg)
                        matches.append(MaterialFileMatch(
                            employee_name=emp.name,
                            material_type="全部",
                            source_path=src,
                            relative_source_path=src.name,
                            matched_by="读取失败",
                            target_filename=f,
                            mismatch_warning=f"⚠️ 文件复制或读取失败: {e}",
                        ))
                        continue

                    try:
                        rel_p = str(src.relative_to(folder_path.parent))
                    except ValueError:
                        rel_p = src.name

                    # 尝试轻量分析图片是否有信息不匹配
                    ocr_name, ocr_id = "", ""
                    if src.suffix.lower() in IMAGE_EXTENSIONS:
                        _, _, _, ocr_name, ocr_id = _classify_by_ocr(src)
                    mismatch = _check_mismatch_warning(emp, ocr_name, ocr_id, duplicate_warning)

                    matches.append(MaterialFileMatch(
                        employee_name=emp.name,
                        material_type="全部",
                        source_path=src,
                        relative_source_path=rel_p,
                        matched_by=match_reason,
                        target_filename=f,
                        target_path=dst,
                        extracted_person_name=ocr_name,
                        extracted_id_card=ocr_id,
                        mismatch_warning=mismatch,
                    ))
        except Exception as e:
            warnings.append(f"无法访问文件夹 {folder_path}: {e}")


def _collect_specific_materials(
    emp: TargetEmployee,
    matched_folders: list[tuple[Path, str]],
    out_path: Path,
    mode: str,
    requested_materials: list[str],
    matches: list[MaterialFileMatch],
    warnings: list[str],
    duplicate_warning: str = "",
) -> list[str]:
    """指定材料模式：在匹配到的文件夹中精准搜集对应材料类型的文件。

    使用 文件名特征 + 文档内容检索 + 离线视觉 OCR 进行三级精准识别，
    如实判定存在与缺失，并自动核对证件姓名/号码一致性。
    """
    clean_emp = safe_filename(emp.name)
    # mat_type -> list of (source_path, rel_source, match_reason, subtype, ocr_name, ocr_id)
    found: dict[str, list[tuple[Path, str, str, str, str, str]]] = {m: [] for m in requested_materials}
    seen_hashes: set[tuple[int, str]] = set()

    for folder_path, folder_reason in matched_folders:
        try:
            for root, dirs, files in os.walk(folder_path):
                dirs[:] = [d for d in dirs if not d.startswith(".")]
                for f in files:
                    if _is_junk_or_temp_file(f):
                        continue
                    f_path = Path(root) / f
                    ext = f_path.suffix.lower()
                    if ext not in SUPPORTED_FILE_EXTENSIONS:
                        continue

                    # 同一员工内部跨目录完全相同文件去重
                    sig = _get_file_signature(f_path)
                    if sig in seen_hashes:
                        continue

                    try:
                        rel_p = str(f_path.relative_to(folder_path.parent))
                    except ValueError:
                        rel_p = f_path.name

                    # 深度分类：三级智能识别（文件名 + 正文 + OCR）
                    try:
                        classified_mat_type, match_method, subtype, ocr_name, ocr_id = _classify_material_type(
                            f_path, f, requested_materials
                        )
                    except Exception as exc:
                        warnings.append(f"文件读取异常 {f_path.name}: {exc}")
                        matches.append(MaterialFileMatch(
                            employee_name=emp.name,
                            material_type="未知",
                            source_path=f_path,
                            relative_source_path=rel_p,
                            matched_by="读取失败",
                            target_filename=f,
                            mismatch_warning=f"⚠️ 文件读取损坏或异常: {exc}",
                        ))
                        continue

                    if classified_mat_type and classified_mat_type in requested_materials:
                        if not any(existing[0] == f_path for existing in found[classified_mat_type]):
                            seen_hashes.add(sig)
                            found[classified_mat_type].append(
                                (f_path, rel_p, match_method or folder_reason, subtype, ocr_name, ocr_id)
                            )
        except Exception as e:
            warnings.append(f"无法访问文件夹 {folder_path}: {e}")

    # 复制匹配到的真实文件到输出目录
    missing_list: list[str] = []
    for mat_type in requested_materials:
        m_list = found[mat_type]
        if not m_list:
            missing_list.append(mat_type)
            continue

        for seq, (src_path, rel_p, match_reason, subtype, ocr_name, ocr_id) in enumerate(m_list, start=1):
            ext = src_path.suffix
            clean_mat = safe_filename(mat_type)

            # 如果 OCR 或文件名识别出具体子类型（如"正面"、"反面"）
            if subtype:
                target_name = f"{clean_emp}_{clean_mat}_{subtype}{ext}"
            else:
                suffix = f"_{seq}" if len(m_list) > 1 else ""
                target_name = f"{clean_emp}_{clean_mat}{suffix}{ext}"

            if mode == MODE_BY_EMPLOYEE:
                dest_dir = out_path / clean_emp
            elif mode == MODE_BY_MATERIAL:
                dest_dir = out_path / clean_mat
            else:  # FLAT
                dest_dir = out_path

            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_file = dest_dir / target_name

            try:
                shutil.copy2(src_path, dest_file)
            except Exception as e:
                warnings.append(f"复制失败：{src_path} → {dest_file}: {e}")
                matches.append(MaterialFileMatch(
                    employee_name=emp.name,
                    material_type=mat_type,
                    source_path=src_path,
                    relative_source_path=rel_p,
                    matched_by="写入失败",
                    target_filename=target_name,
                    mismatch_warning=f"⚠️ 文件复制失败: {e}",
                ))
                continue

            mismatch = _check_mismatch_warning(emp, ocr_name, ocr_id, duplicate_warning)

            matches.append(MaterialFileMatch(
                employee_name=emp.name,
                material_type=mat_type,
                source_path=src_path,
                relative_source_path=rel_p,
                matched_by=match_reason,
                target_filename=target_name,
                target_path=dest_file,
                extracted_person_name=ocr_name,
                extracted_id_card=ocr_id,
                mismatch_warning=mismatch,
            ))

    return missing_list


# ---------------------------------------------------------------------------
# Excel 报告生成
# ---------------------------------------------------------------------------

def _write_excel_report(
    report_path: Path,
    employees: list[TargetEmployee],
    requested_materials: list[str],
    all_matches: list[MaterialFileMatch],
    collect_all: bool,
    warnings: list[str] | None = None,
) -> None:
    """Generate structured summary and missing Excel report with optimized columns and wrap text."""
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "资料提取汇总与缺失清单"

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    normal_font = Font(name="微软雅黑", size=10)
    ok_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ok_font = Font(name="微软雅黑", size=10, color="375623")
    missing_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    missing_font = Font(name="微软雅黑", size=10, bold=True, color="C65911")
    warning_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    warning_font = Font(name="微软雅黑", size=10, bold=True, color="BD8100")
    stat_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    ws_summary["A1"] = "员工资料提取汇总与缺失清单"
    ws_summary["A1"].font = title_font
    ws_summary["A1"].alignment = Alignment(vertical="center")

    emp_file_counts: dict[str, int] = {}
    emp_mismatch_warnings: dict[str, list[str]] = {}
    for m in all_matches:
        emp_file_counts[m.employee_name] = emp_file_counts.get(m.employee_name, 0) + 1
        if m.mismatch_warning:
            emp_mismatch_warnings.setdefault(m.employee_name, []).append(m.mismatch_warning)

    total_emp = len(employees)
    total_files = len(all_matches)
    found_emp = sum(1 for emp in employees if emp_file_counts.get(emp.name, 0) > 0)
    not_found_emp = total_emp - found_emp

    ws_summary["A3"] = "统计概要"
    ws_summary["A3"].font = Font(name="微软雅黑", size=11, bold=True)

    stats = [
        ("目标员工总数", total_emp, "已提取文件总数", total_files),
        ("已找到员工数", found_emp, "未找到员工数", not_found_emp),
    ]

    for row_idx, (k1, v1, k2, v2) in enumerate(stats, start=4):
        ws_summary[f"A{row_idx}"] = k1
        ws_summary[f"B{row_idx}"] = v1
        ws_summary[f"C{row_idx}"] = k2
        ws_summary[f"D{row_idx}"] = v2
        for col_let in ["A", "B", "C", "D"]:
            c = ws_summary[f"{col_let}{row_idx}"]
            c.border = thin_border
            c.font = normal_font
            if col_let in ["A", "C"]:
                c.fill = stat_fill
                c.font = Font(name="微软雅黑", size=10, bold=True)
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")

    # Detail Table
    start_row = 7
    if collect_all:
        headers = ["序号", "员工姓名", "身份证号码", "提取状态", "提取文件数", "信息核对预警 / 备注"]
    else:
        headers = ["序号", "员工姓名", "身份证号码", "提取进度"] + requested_materials + ["信息核对预警 / 备注"]

    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws_summary.cell(start_row, c_idx, h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    current_r = start_row + 1
    for idx, emp in enumerate(employees, start=1):
        ws_summary.cell(current_r, 1, idx).alignment = Alignment(horizontal="center", vertical="center")
        ws_summary.cell(current_r, 2, emp.name).alignment = Alignment(horizontal="center", vertical="center")
        ws_summary.cell(current_r, 3, emp.id_card).alignment = Alignment(horizontal="center", vertical="center")

        file_count = emp_file_counts.get(emp.name, 0)
        warn_list = emp_mismatch_warnings.get(emp.name, [])

        if collect_all:
            status_cell = ws_summary.cell(current_r, 4)
            count_cell = ws_summary.cell(current_r, 5)
            if file_count > 0:
                status_cell.value = "已找到"
                status_cell.fill = ok_fill
                status_cell.font = ok_font
                count_cell.value = file_count
            else:
                status_cell.value = "未找到"
                status_cell.fill = missing_fill
                status_cell.font = missing_font
                count_cell.value = 0
            status_cell.alignment = Alignment(horizontal="center", vertical="center")
            count_cell.alignment = Alignment(horizontal="center", vertical="center")

            warn_cell = ws_summary.cell(current_r, 6)
            if warn_list:
                warn_cell.value = "；".join(sorted(set(warn_list)))
                warn_cell.fill = warning_fill
                warn_cell.font = warning_font
            else:
                warn_cell.value = "正常 (信息一致)" if file_count > 0 else "-"
                warn_cell.fill = ok_fill if file_count > 0 else stat_fill
                warn_cell.font = ok_font if file_count > 0 else normal_font
            warn_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            emp_matches_by_type: dict[str, int] = {}
            for m in all_matches:
                if m.employee_name == emp.name and m.material_type in requested_materials:
                    emp_matches_by_type[m.material_type] = emp_matches_by_type.get(m.material_type, 0) + 1

            found_count = sum(1 for m in requested_materials if emp_matches_by_type.get(m, 0) > 0)
            total_req = len(requested_materials)
            status_cell = ws_summary.cell(current_r, 4, f"{found_count}/{total_req}")
            status_cell.alignment = Alignment(horizontal="center", vertical="center")
            if found_count == total_req:
                status_cell.fill = ok_fill
                status_cell.font = ok_font
            else:
                status_cell.fill = missing_fill
                status_cell.font = missing_font

            for col_offset, mat_type in enumerate(requested_materials, start=5):
                count = emp_matches_by_type.get(mat_type, 0)
                cell = ws_summary.cell(current_r, col_offset)
                cell.border = thin_border
                if count > 0:
                    cell.value = f"已提取({count}份)"
                    cell.fill = ok_fill
                    cell.font = ok_font
                else:
                    cell.value = "缺失"
                    cell.fill = missing_fill
                    cell.font = missing_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            warn_col_idx = 5 + len(requested_materials)
            warn_cell = ws_summary.cell(current_r, warn_col_idx)
            if warn_list:
                warn_cell.value = "；".join(sorted(set(warn_list)))
                warn_cell.fill = warning_fill
                warn_cell.font = warning_font
            else:
                warn_cell.value = "正常 (信息一致)" if found_count > 0 else "-"
                warn_cell.fill = ok_fill if found_count > 0 else stat_fill
                warn_cell.font = ok_font if found_count > 0 else normal_font
            warn_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for c in range(1, len(headers) + 1):
            ws_summary.cell(current_r, c).border = thin_border
            if c <= 3:
                ws_summary.cell(current_r, c).font = normal_font

        current_r += 1

    # 优化列宽：预警列固定 38 并换行，普通列自适应
    for col in ws_summary.columns:
        col_letter = get_column_letter(col[0].column)
        col_name = str(ws_summary.cell(start_row, col[0].column).value or "")
        if "预警" in col_name or "备注" in col_name:
            ws_summary.column_dimensions[col_letter].width = 38
        else:
            max_len = max(len(str(cell.value or "")) for cell in col if cell.row >= start_row)
            ws_summary.column_dimensions[col_letter].width = max(min(max_len * 2 + 2, 28), 12)

    # Sheet 2: Matched File List
    ws_files = wb.create_sheet(title="提取文件明细清单")
    file_headers = ["序号", "员工姓名", "材料类型", "目标文件名", "证件识别姓名", "证件识别号码", "信息匹配校验", "匹配依据", "原始文件路径"]
    for c_idx, h_text in enumerate(file_headers, start=1):
        cell = ws_files.cell(1, c_idx, h_text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for idx, match_item in enumerate(all_matches, start=1):
        r = idx + 1
        ws_files.cell(r, 1, idx).alignment = Alignment(horizontal="center", vertical="center")
        ws_files.cell(r, 2, match_item.employee_name).alignment = Alignment(horizontal="center", vertical="center")
        ws_files.cell(r, 3, match_item.material_type).alignment = Alignment(horizontal="center", vertical="center")
        ws_files.cell(r, 4, match_item.target_filename).alignment = Alignment(horizontal="left", vertical="center")
        ws_files.cell(r, 5, match_item.extracted_person_name or "-").alignment = Alignment(horizontal="center", vertical="center")
        ws_files.cell(r, 6, match_item.extracted_id_card or "-").alignment = Alignment(horizontal="center", vertical="center")

        chk_cell = ws_files.cell(r, 7)
        if match_item.mismatch_warning:
            chk_cell.value = match_item.mismatch_warning
            chk_cell.fill = warning_fill
            chk_cell.font = warning_font
        elif match_item.extracted_person_name or match_item.extracted_id_card:
            chk_cell.value = "✓ 一致"
            chk_cell.fill = ok_fill
            chk_cell.font = ok_font
        else:
            chk_cell.value = "-"
            chk_cell.font = normal_font
        chk_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws_files.cell(r, 8, match_item.matched_by).alignment = Alignment(horizontal="center", vertical="center")
        ws_files.cell(r, 9, match_item.relative_source_path).alignment = Alignment(horizontal="left", vertical="center")

        for c in range(1, len(file_headers) + 1):
            cell = ws_files.cell(r, c)
            cell.border = thin_border
            if c not in (7,):
                cell.font = normal_font

    for col in ws_files.columns:
        col_letter = get_column_letter(col[0].column)
        col_name = str(ws_files.cell(1, col[0].column).value or "")
        if "校验" in col_name or "路径" in col_name:
            ws_files.column_dimensions[col_letter].width = 38
        else:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_files.column_dimensions[col_letter].width = max(min(max_len * 2 + 2, 30), 12)

    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)
    wb.close()
