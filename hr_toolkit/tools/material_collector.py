from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import threading
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..common.excel_compat import ensure_xlsx_workbook, is_supported_excel_file
from ..common.filenames import safe_filename

try:
    from zoneinfo import ZoneInfo
    _BEIJING_TZ = ZoneInfo("Asia/Shanghai")
except Exception:  # pragma: no cover - 兼容性回退
    _BEIJING_TZ = timezone(timedelta(hours=8))


def _beijing_now_str() -> str:
    """统一使用北京时间生成标准格式时间字符串 YYYY-MM-DD HH:MM:SS。"""
    return datetime.now(tz=_BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S")


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


# OCR 智能索引缓存：写入资料库根目录的隐藏 JSON 文件
_OCR_CACHE_FILE_NAME = ".hr_material_index_cache.json"
_OCR_CACHE_VERSION = 3
_OCR_CACHE_TEXT_SNIPPET_MAX = 4096
_OCR_CACHE_FILE_MAX_BYTES = 64 * 1024 * 1024
_OCR_CACHE_FILE_TRIM_BYTES = 48 * 1024 * 1024
_OCR_CACHE_ENTRY_MAX_AGE_DAYS = 90
_OCR_CACHE_HASH_WINDOW = 1 * 1024 * 1024
_OCR_CACHE_HASH_TRIGGER_SIZE = 10 * 1024 * 1024


# ---------------------------------------------------------------------------
# OCR 智能索引缓存层：纯函数（无副作用，便于单测）
# ---------------------------------------------------------------------------


def _get_engine_signature() -> str:
    """获取当前 OCR 引擎的版本签名；用于缓存条目与引擎版本一致性校验。"""
    try:
        import rapidocr_onnxruntime

        version = getattr(rapidocr_onnxruntime, "__version__", "unknown")
        return f"rapidocr_onnxruntime@{version}"
    except Exception:
        return "rapidocr_onnxruntime@unknown"


def _compute_file_fingerprint(file_path: Path) -> tuple[int, float, str] | None:
    """旧文件夹模式的缓存指纹；大文件保持原有的前 1MB hash 行为。"""
    try:
        stat = file_path.stat()
    except (FileNotFoundError, OSError):
        return None

    size = stat.st_size
    mtime = stat.st_mtime
    sha = hashlib.sha256()
    try:
        with open(file_path, "rb") as fp:
            if size <= _OCR_CACHE_HASH_TRIGGER_SIZE:
                while True:
                    chunk = fp.read(64 * 1024)
                    if not chunk:
                        break
                    sha.update(chunk)
            else:
                sha.update(fp.read(_OCR_CACHE_HASH_WINDOW))
    except OSError:
        return None
    return (size, mtime, sha.hexdigest())


def _windows_file_change_time(file_path: Path) -> int | None:
    """Return the Windows file change time, not ``stat().st_ctime``.

    Python 3.12 still exposes file creation time through ``st_ctime`` on
    Windows.  Creation time does not change when an existing file is
    overwritten, so it cannot safely guard the metadata-only cache fast path.
    ``FILE_BASIC_INFO.ChangeTime`` is the NT file change token we need here.
    """
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class _FileBasicInfo(ctypes.Structure):
            _fields_ = [
                ("CreationTime", ctypes.c_longlong),
                ("LastAccessTime", ctypes.c_longlong),
                ("LastWriteTime", ctypes.c_longlong),
                ("ChangeTime", ctypes.c_longlong),
                ("FileAttributes", wintypes.DWORD),
            ]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        create_file = kernel32.CreateFileW
        create_file.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE,
        ]
        create_file.restype = wintypes.HANDLE
        get_file_information = kernel32.GetFileInformationByHandleEx
        get_file_information.argtypes = [
            wintypes.HANDLE,
            ctypes.c_int,
            wintypes.LPVOID,
            wintypes.DWORD,
        ]
        get_file_information.restype = wintypes.BOOL
        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        close_handle.restype = wintypes.BOOL

        file_read_attributes = 0x0080
        file_share_all = 0x0001 | 0x0002 | 0x0004
        open_existing = 3
        file_attribute_normal = 0x0080
        invalid_handle = wintypes.HANDLE(-1).value
        handle = create_file(
            os.path.abspath(os.fspath(file_path)),
            file_read_attributes,
            file_share_all,
            None,
            open_existing,
            file_attribute_normal,
            None,
        )
        if handle == invalid_handle:
            return None
        try:
            info = _FileBasicInfo()
            file_basic_info = 0
            if not get_file_information(
                handle,
                file_basic_info,
                ctypes.byref(info),
                ctypes.sizeof(info),
            ):
                return None
            change_time = int(info.ChangeTime)
            return change_time if change_time > 0 else None
        finally:
            close_handle(handle)
    except Exception:  # pragma: no cover - platform API failure must degrade safely
        # FAT/network shares or restricted handles may not expose ChangeTime.
        # Callers treat None as unsafe for metadata-only reuse and hash instead.
        return None


def _file_change_token(file_path: Path, stat_result: os.stat_result) -> int | None:
    if os.name == "nt":
        return _windows_file_change_time(file_path)
    return stat_result.st_ctime_ns


def _stream_full_sha256(file_path: Path) -> str | None:
    sha = hashlib.sha256()
    try:
        with open(file_path, "rb") as fp:
            while True:
                chunk = fp.read(1024 * 1024)
                if not chunk:
                    break
                sha.update(chunk)
    except OSError:
        return None
    return sha.hexdigest()


def _compute_full_file_fingerprint(file_path: Path) -> tuple[int, float, str] | None:
    """TASK-8 无序索引专用完整指纹，保证大文件后半段变化也能失效。

    返回 None 表示文件无法访问。首次索引流式计算完整 SHA-256；后续同一路径
    会先用 size/mtime/可靠变更标记命中路径索引，避免重复读取大文件。
    """
    initial_metadata = _flat_path_metadata(file_path)
    if initial_metadata is None:
        return None
    size, mtime_ns, change_token = initial_metadata
    sha = _stream_full_sha256(file_path)
    if sha is None:
        return None
    final_metadata = _flat_path_metadata(file_path)
    if final_metadata != initial_metadata:
        return None

    if change_token is None:
        # Without a reliable OS change token, verify the bytes twice.  This
        # fallback is slower but prevents same-size, restored-mtime changes
        # from producing a stable-looking fingerprint on unsupported volumes.
        verified_sha = _stream_full_sha256(file_path)
        if verified_sha != sha or _flat_path_metadata(file_path) != final_metadata:
            return None

    return (size, mtime_ns / 1_000_000_000, sha)


def _compute_cache_key(
    file_path: Path,
    employee_key: str = "",
) -> str | None:
    """根据文件二进制内容哈希 + 文件大小生成唯一指纹（完全脱离文件名与路径）。

    不管文件被重命名为任何名称、移动到任何子目录，只要内容未变，指纹恒定不变。
    """
    fingerprint = _compute_file_fingerprint(file_path)
    if fingerprint is None:
        return None
    size, _mtime, sha = fingerprint
    return f"{sha[:24]}_{size}"


def _mask_id_card(id_card: str) -> str:
    """身份证号脱敏：仅保留前 4 与后 4 位；空串直接返回。"""
    if not id_card:
        return ""
    if len(id_card) <= 8:
        return id_card[:2] + "*" * (len(id_card) - 4) + id_card[-2:]
    return id_card[:4] + "*" * (len(id_card) - 8) + id_card[-4:]


def _hash_id_card(id_card: str) -> str:
    """身份证号 sha256；用于 mismatch 校验但不在缓存文件留明文。"""
    if not id_card:
        return ""
    return hashlib.sha256(id_card.encode("utf-8")).hexdigest()[:16]


def _hash_phone(phone: str) -> str:
    if not phone:
        return ""
    return hashlib.sha256(phone.encode("utf-8")).hexdigest()[:16]


def _new_ocr_cache() -> dict[str, Any]:
    now = _beijing_now_str()
    return {
        "version": _OCR_CACHE_VERSION,
        "engine_signature": _get_engine_signature(),
        "created_at": now,
        "updated_at": now,
        "entries": {},
        "paths": {},
    }


def _load_ocr_cache(cache_path: Path) -> dict[str, Any]:
    """读取缓存 JSON；损坏时返回空结构并由上层决定是否重建。"""
    if cache_path.is_symlink() or not cache_path.is_file():
        return _new_ocr_cache()

    try:
        raw = cache_path.read_text(encoding="utf-8")
    except OSError:
        return _new_ocr_cache()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return _new_ocr_cache()

    if not isinstance(data, dict):
        return _new_ocr_cache()

    data.setdefault("version", _OCR_CACHE_VERSION)
    if not isinstance(data.get("entries"), dict):
        data["entries"] = {}
    if not isinstance(data.get("paths"), dict):
        data["paths"] = {}
    return data


def _save_ocr_cache(cache_path: Path, data: dict[str, Any]) -> bool:
    """原子写缓存：tmp + os.replace；返回是否成功。

    任何 OSError / PermissionError 都被捕获并返回 False，由调用方降级。
    """
    data["version"] = _OCR_CACHE_VERSION
    data["engine_signature"] = _get_engine_signature()
    data["updated_at"] = _beijing_now_str()
    data.setdefault("created_at", data["updated_at"])
    data.setdefault("entries", {})
    data.setdefault("paths", {})

    if cache_path.is_symlink():
        return False
    tmp_path = cache_path.with_name(
        f".{cache_path.name}.{os.getpid()}.{time.time_ns()}.tmp"
    )
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        if cache_path.parent.is_symlink():
            return False
        with tmp_path.open("x", encoding="utf-8") as temp_file:
            temp_file.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
        try:
            tmp_path.chmod(0o600)
        except OSError:
            pass
    except (FileExistsError, OSError):
        return False

    try:
        os.replace(tmp_path, cache_path)
    except OSError:
        # 兜底：清理残留 tmp
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        return False
    return True


def _trim_cache_by_age_and_size(data: dict[str, Any]) -> None:
    """对缓存做 LRU + 体积治理：清理过期与超量条目。

    注：当前 cache key 基于 (content_hash + size)，不依赖文件路径，
因此文件改名/移动不会导致缓存失效；文件被删除后条目会自然过期（90 天未命中后清理）。
    """
    entries: dict[str, Any] = data.get("entries") or {}
    paths: dict[str, Any] = data.get("paths") or {}
    if not entries:
        paths.clear()
        return

    # 1. 清理超过 90 天未验证的条目
    now = datetime.now(tz=_BEIJING_TZ)
    cutoff = now - timedelta(days=_OCR_CACHE_ENTRY_MAX_AGE_DAYS)

    def _parse_ts(ts: str) -> datetime:
        try:
            return datetime.strptime(ts, "%Y-%m-%d %H:%M:%S").replace(tzinfo=_BEIJING_TZ)
        except (TypeError, ValueError):
            try:
                d = datetime.fromisoformat(ts)
                return d if d.tzinfo is not None else d.replace(tzinfo=_BEIJING_TZ)
            except (TypeError, ValueError):
                return now

    stale_keys: list[str] = []
    for key, entry in entries.items():
        ts_str = entry.get("verified_at") if isinstance(entry, dict) else None
        if not ts_str:
            stale_keys.append(key)
            continue
        ts = _parse_ts(ts_str)
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=_BEIJING_TZ)
        if ts < cutoff:
            stale_keys.append(key)

    for key in stale_keys:
        entries.pop(key, None)

    def _drop_orphan_paths() -> None:
        for rel_path, path_entry in list(paths.items()):
            if not isinstance(path_entry, dict) or path_entry.get("cache_key") not in entries:
                paths.pop(rel_path, None)

    _drop_orphan_paths()

    # 2. 估算大小并按 verified_at 升序删除
    serialized_len = sum(
        len(json.dumps(v, ensure_ascii=False).encode("utf-8")) + len(k.encode("utf-8"))
        for k, v in entries.items()
    )
    if serialized_len <= _OCR_CACHE_FILE_MAX_BYTES:
        return

    sorted_keys = sorted(
        entries.keys(),
        key=lambda k: entries[k].get("verified_at", "") if isinstance(entries[k], dict) else "",
    )
    for key in sorted_keys:
        if serialized_len <= _OCR_CACHE_FILE_TRIM_BYTES:
            break
        try:
            serialized_len -= (
                len(json.dumps(entries[key], ensure_ascii=False).encode("utf-8"))
                + len(key.encode("utf-8"))
            )
        except (TypeError, ValueError):
            serialized_len -= 1024
        entries.pop(key, None)
    _drop_orphan_paths()


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

# 资料库的组织形式与上面的“输出归类模式”是两个独立维度。默认值始终走旧逻辑，
# 避免 TASK-8 影响当前已经投入使用的三种输出结构。
LIBRARY_MODE_PERSON_FOLDER = "person_folder"
LIBRARY_MODE_FLAT_OCR = "flat_ocr"
LIBRARY_MODES = {LIBRARY_MODE_PERSON_FOLDER, LIBRARY_MODE_FLAT_OCR}
LIBRARY_MODE_LABELS = {
    "按人员文件夹查找（原模式）": LIBRARY_MODE_PERSON_FOLDER,
    "无序平铺资料库（OCR 索引）": LIBRARY_MODE_FLAT_OCR,
}
LIBRARY_MODE_LABELS_REVERSE = {v: k for k, v in LIBRARY_MODE_LABELS.items()}

# 预置常见材料类型及其多维度别名同义词库（互斥排他分类，绝不串门）
MATERIAL_SYNONYMS: dict[str, list[str]] = {
    "身份证": [
        "身份证", "身分证", "sfz", "idcard", "id_card", "id", "identity",
        "正面", "反面", "人像面", "国徽面", "人像", "国徽", "A面", "B面", "正反面", "正反",
        "zhengmian", "fanmian", "zm", "fm",
        "身份证正面", "身份证反面", "身份证正反面", "身份证复印件", "身份证照片", "身份证件", "证件",
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
        "资格证", "职业资格证", "职称证", "技能证", "驾驶证", "驾照", "上岗证", "从业资格", "资格", "职称", "技能", "证书", "zige", "jineng",
        "certificate", "license",
    ],
    "安全员证": [
        "安全员证", "安全员", "安全员证书", "安全考核合格证", "建安C证", "建安A证", "建安B证",
        "安管人员", "安全生产考核", "C证", "A证", "B证", "安全考核", "安全员合格证", "anquanyuan",
    ],
    "特种证书": [
        "特种证书", "特种作业证", "特种作业操作证", "特种作业证书", "特种作业", "特种设备",
        "特种操作证", "特种工", "高处作业", "电工作业", "焊接作业", "电工证", "焊工证", "登高证",
        "操作证", "tezhong",
    ],
    "证件照片": [
        "一寸照", "二寸照", "证件照", "寸照", "登记照", "蓝底", "白底", "红底", "个人照片", "照片",
        "相片", "头像", "个人照", "photo", "pic", "avatar", "head",
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
    cache_hit: bool = False  # 新增：OCR 缓存命中标记
    employee_identity_key: str = ""

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
            "cache_hit": self.cache_hit,
            "employee_identity_key": self.employee_identity_key,
        }


@dataclass
class MaterialCollectResult:
    library_dir: Path
    output_dir: Path
    zip_path: Path | None = None
    report_path: Path | None = None
    mode: str = MODE_BY_EMPLOYEE
    library_mode: str = LIBRARY_MODE_PERSON_FOLDER
    target_employees: list[TargetEmployee] = field(default_factory=list)
    requested_materials: list[str] = field(default_factory=list)
    matches: list[MaterialFileMatch] = field(default_factory=list)
    missing_records: dict[str, list[str]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    folder_match_counts: dict[str, int] = field(default_factory=dict)
    employee_result_keys: list[str] = field(default_factory=list, repr=False)

    # === 新增：OCR 缓存层指标（默认值保持向后兼容） ===
    ocr_cache_enabled: bool = True
    ocr_cache_hits: int = 0
    ocr_cache_misses: int = 0
    ocr_cache_invalidated: int = 0
    ocr_cache_path: str | None = None
    ocr_cache_skipped_reason: str | None = None

    # === 占位字段：未来隐私开关（本次不实现行为，仅留接口） ===
    # TODO: 当 HR 提报隐私报送需求时启用，本次保持 None / False 以保证报送数据完整
    zip_password: str | None = None
    mask_sensitive: bool = False

    @property
    def total_employees(self) -> int:
        return len(self.target_employees)

    @property
    def matched_file_count(self) -> int:
        return len(self.matches)

    @property
    def complete_employee_count(self) -> int:
        keys = (
            self.employee_result_keys
            if len(self.employee_result_keys) == len(self.target_employees)
            else [employee.name for employee in self.target_employees]
        )
        return sum(1 for key in keys if not self.missing_records.get(key))

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": TOOL_NAME,
            "library_dir": str(self.library_dir),
            "output_dir": str(self.output_dir),
            "zip_path": str(self.zip_path) if self.zip_path else None,
            "report_path": str(self.report_path) if self.report_path else None,
            "mode": self.mode,
            "library_mode": self.library_mode,
            "total_employees": self.total_employees,
            "matched_file_count": self.matched_file_count,
            "complete_employee_count": self.complete_employee_count,
            "requested_materials": self.requested_materials,
            "matches": [m.to_dict() for m in self.matches],
            "missing_records": self.missing_records,
            "warnings": self.warnings,
            # 缓存层指标
            "ocr_cache_enabled": self.ocr_cache_enabled,
            "ocr_cache_hits": self.ocr_cache_hits,
            "ocr_cache_misses": self.ocr_cache_misses,
            "ocr_cache_invalidated": self.ocr_cache_invalidated,
            "ocr_cache_path": self.ocr_cache_path,
            "ocr_cache_skipped_reason": self.ocr_cache_skipped_reason,
            # 占位字段
            "zip_password": self.zip_password,
            "mask_sensitive": self.mask_sensitive,
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


def _extract_pdf_image_bytes(file_path: Path) -> bytes | None:
    """如果 PDF 为纯图片扫描版，从二进制流中提取首张内嵌图片（JPEG/PNG）用于 OCR。"""
    try:
        data = file_path.read_bytes()
        idx = 0
        while True:
            s_idx = data.find(b"stream", idx)
            if s_idx == -1:
                break
            start = s_idx + 6
            if data[start:start+2] == b"\r\n":
                start += 2
            elif data[start:start+1] == b"\n":
                start += 1
            e_idx = data.find(b"endstream", start)
            if e_idx == -1:
                break
            chunk = data[start:e_idx]
            # JPEG 魔数 \xff\xd8\xff 或 PNG 魔数 \x89PNG
            if chunk.startswith(b"\xff\xd8\xff") or chunk.startswith(b"\x89PNG\r\n\x1a\n"):
                return chunk
            idx = e_idx + 9
    except Exception:
        pass
    return None


def _build_doc_format_hint(file_path: Path) -> str | None:
    """针对旧版 .doc 文件给出友好提示，避免静默失败。

    .doc 为二进制 OLE 容器，直接 utf-8 解码几乎全部乱码；本工具不引入外部依赖，
    因此对该格式仅作识别提示，让用户主动另存为 .docx 后重跑。
    """
    if file_path.suffix.lower() != ".doc":
        return None
    return (
        f"⚠️ 旧版 Word 文件 {file_path.name} 为 .doc 格式，"
        "工具无法保证识别准确性；建议另存为 .docx 后重试。"
    )


def _build_employee_key(emp: TargetEmployee) -> str:
    """构造 (姓名|身份证) 维度的员工键，用于缓存与同名员工隔离。"""
    name = (emp.name or "").strip()
    id_card = (emp.id_card or "").strip()
    return f"{name}|{id_card}"


def _lookup_ocr_cache(
    cache: dict[str, Any],
    file_path: Path,
    employee_key: str = "",
    rel_path: str = "",
) -> tuple[str, str, str, str, str] | None:
    """按文件内容哈希指纹查询缓存；命中则返回 (material, method, sub, name, id_hash)。

    只要文件二进制内容没变，无论文件名如何修改、移动到何处，都能 100% 瞬间命中。
    """
    entries: dict[str, Any] = cache.get("entries") or {}
    if not entries:
        return None

    target_key = _compute_cache_key(file_path, employee_key)
    if not target_key:
        return None

    entry = entries.get(target_key)
    if not isinstance(entry, dict):
        return None

    return (
        entry.get("material_type") or "",
        entry.get("match_method") or "ocr_cached",
        entry.get("subtype") or "",
        entry.get("extracted_name") or "",
        entry.get("extracted_id_hash") or "",
    )


def _store_ocr_cache(
    cache: dict[str, Any],
    file_path: Path,
    material_type: str,
    match_method: str,
    subtype: str,
    extracted_name: str,
    extracted_id: str,
    employee_key: str = "",
    rel_path: str = "",
) -> None:
    """OCR 成功后按文件内容哈希指纹回写缓存条目。"""
    fingerprint = _compute_file_fingerprint(file_path)
    if fingerprint is None:
        return
    size, mtime, sha = fingerprint

    cache_key = f"{sha[:24]}_{size}"
    entries: dict[str, Any] = cache.setdefault("entries", {})
    entries[cache_key] = {
        "content_hash": sha[:24],
        "source_size": size,
        "source_mtime": mtime,
        "material_type": material_type,
        "match_method": match_method,
        "subtype": subtype,
        "extracted_name": extracted_name,
        "extracted_id_hash": _hash_id_card(extracted_id),
        "verified_at": _beijing_now_str(),
        "sample_filename": file_path.name,
    }


_DOC_CONTENT_PATTERNS: dict[str, list[str]] = {
    "劳动合同": ["劳动合同", "用工合同", "劳务合同", "聘用合同", "用工协议", "劳动期限", "工作内容", "劳动报酬", "劳动争议", "解除劳动合同", "劳动法", "甲乙双方根据", "试用期"],
    "学历证明": ["毕业证书", "学位证书", "教育部学历证书", "学信网", "普通高等学校", "学士学位", "硕士学位", "博士学位"],
    "安全员证": ["安全生产考核合格证书", "建筑施工企业项目负责人安全生产考核合格证书", "建筑施工企业专职安全生产管理人员安全生产考核合格证书", "安全员考核合格证", "安全员C证", "安全员A证", "安全员B证"],
    "特种证书": ["特种作业操作证", "特种作业人员操作证", "特种设备作业人员证", "特种作业", "特种设备作业"],
    "资格证书": ["职业资格证书", "专业技术职务资格证书", "职称证书", "技能等级证书", "中华人民共和国机动车驾驶证"],
}


def _sanitize_cached_text(text: str) -> str:
    """缓存 OCR 摘要前脱敏证件号、手机号，并限制体积。"""
    compact = re.sub(r"\s+", " ", str(text or "")).strip()
    compact = re.sub(
        r"(?<!\d)(\d{17}[\dXx]|\d{15})(?!\d)",
        lambda match: _mask_id_card(match.group(1)),
        compact,
    )
    compact = re.sub(
        r"(?<!\d)(1[3-9]\d{9})(?!\d)",
        lambda match: match.group(1)[:3] + "****" + match.group(1)[-4:],
        compact,
    )
    return compact[:_OCR_CACHE_TEXT_SNIPPET_MAX]


def _normalize_person_name(name: str) -> str:
    return re.sub(r"[\s·•・]", "", str(name or "")).strip()


def _extract_person_names(text: str) -> list[str]:
    """从证件、合同等正文提取明确带字段标签的人名，不依赖文件名。"""
    normalized = re.sub(r"[\r\n\t]+", " ", str(text or ""))
    labels = (
        "劳动者姓名", "员工姓名", "人员姓名", "持证人姓名", "申请人姓名",
        "姓名", "持证人", "劳动者", "乙方", "受聘人", "开户人",
    )
    names: list[str] = []
    suffix_noise = (
        "性别", "民族", "出生", "住址", "身份证", "公民身份", "证件号码",
        "电话", "手机号", "合同", "岗位", "部门", "文化程度", "有效期",
    )
    for label in labels:
        pattern = rf"{re.escape(label)}\s*(?:名称)?\s*[:：]?\s*([\u3400-\u9fff·•・]{{2,10}})"
        for match in re.finditer(pattern, normalized):
            candidate = match.group(1).strip()
            for suffix in suffix_noise:
                if suffix in candidate:
                    candidate = candidate.split(suffix, 1)[0]
            candidate = candidate.strip()
            if not (2 <= len(_normalize_person_name(candidate)) <= 8):
                continue
            if not _is_valid_person_name(candidate):
                continue
            if candidate not in names:
                names.append(candidate)
    return names


def _extract_id_card(text: str) -> str:
    match = re.search(r"(?<!\d)(\d{17}[\dXx]|\d{15})(?!\d)", str(text or ""))
    return match.group(1) if match else ""


def _extract_phone(text: str) -> str:
    match = re.search(r"(?<!\d)(1[3-9]\d{9})(?!\d)", str(text or ""))
    return match.group(1) if match else ""


def _classify_text_content(
    text: str,
    *,
    requested_types: list[str] | None = None,
    method_prefix: str = "ocr",
) -> tuple[str | None, str, str]:
    """按正文判定材料类型；返回 (类型, 匹配方式, 正反面子类型)。"""
    full_text = str(text or "")
    if not full_text.strip():
        return None, "", ""

    if any(keyword in full_text for keyword in ("特种作业操作证", "特种作业人员", "特种设备作业")):
        return "特种证书", f"{method_prefix}_special_cert", ""
    if any(keyword in full_text for keyword in ("安全生产考核", "安全考核合格", "建安C证", "建安A证", "建安B证")):
        return "安全员证", f"{method_prefix}_safety_cert", ""
    if (
        any(keyword in full_text for keyword in ("劳动合同", "用工合同", "劳务合同", "聘用合同"))
        or ("甲方" in full_text and "乙方" in full_text and any(keyword in full_text for keyword in ("劳动", "报酬", "工作内容")))
    ):
        return "劳动合同", f"{method_prefix}_contract", ""
    if any(keyword in full_text for keyword in ("毕业证书", "学位证书", "学信网", "学历证书", "普通高等学校")):
        return "学历证明", f"{method_prefix}_degree", ""
    if any(keyword in full_text for keyword in ("机动车驾驶证", "驾驶证", "职业资格证书", "职业资格", "职称证书", "技能等级证书")):
        return "资格证书", f"{method_prefix}_certificate", ""
    if re.search(r"\d{16,19}", full_text) and any(keyword in full_text for keyword in ("银行", "银联", "Bank")):
        return "银行卡", f"{method_prefix}_bank", ""
    if "居民身份证" in full_text or (
        "签发机关" in full_text and "有效期限" in full_text
        and "特种" not in full_text and "安全" not in full_text
    ):
        return "身份证", f"{method_prefix}_id_back", "反面"
    extracted_id = _extract_id_card(full_text)
    if (
        "公民身份号码" in full_text
        or ("姓名" in full_text and any(keyword in full_text for keyword in ("住址", "民族", "出生")))
        or extracted_id
    ):
        return "身份证", f"{method_prefix}_id_front", "正面"

    # 自定义材料只接受正文中的完整名称，避免随机文件名造成跨人员误判。
    for requested in sorted(requested_types or [], key=len, reverse=True):
        if requested and requested in full_text:
            return requested, f"{method_prefix}_custom", ""
    return None, "", ""


def _read_ocr_text(file_path: Path) -> tuple[str, list[str], bool]:
    """调用本地 OCR 并兼容 RapidOCR 与测试/旧适配器的两种结果形态。"""
    engine = _get_ocr_engine()
    if engine is None:
        return "", [], False

    target_input: str | bytes = str(file_path)
    if file_path.suffix.lower() == ".pdf":
        img_bytes = _extract_pdf_image_bytes(file_path)
        if not img_bytes:
            return "", [], False
        target_input = img_bytes

    try:
        with _OCR_LOCK:
            result, _ = engine(target_input)
    except Exception:
        return "", [], False
    if not result:
        return "", [], True

    texts: list[str] = []
    for item in result:
        if not isinstance(item, (list, tuple)):
            continue
        if len(item) >= 3 and isinstance(item[1], str):
            texts.append(item[1])
        elif len(item) == 2 and isinstance(item[0], str) and isinstance(item[1], str):
            # 一些本地适配器以 [字段名, 字段值] 返回。
            texts.append(f"{item[0]}：{item[1]}")
        elif len(item) >= 2 and isinstance(item[1], str):
            texts.append(item[1])
    return " ".join(texts), texts, True


def _analyze_ocr_file(
    file_path: Path,
    requested_types: list[str] | None = None,
) -> tuple[str | None, str, str, str, str, str, list[str], bool]:
    full_text, _texts, analysis_complete = _read_ocr_text(file_path)
    names = _extract_person_names(full_text)
    extracted_name = names[0] if names else ""
    extracted_id = _extract_id_card(full_text)
    material, method, subtype = _classify_text_content(
        full_text,
        requested_types=requested_types,
        method_prefix="ocr",
    )
    return material, method, subtype, extracted_name, extracted_id, full_text, names, analysis_complete


def _classify_by_ocr(file_path: Path) -> tuple[str | None, str, str, str, str]:
    """通过本地离线 OCR 识别图片文字并进行材料分类与信息提取。

    Returns: (material_type or None, match_method, subtype_label, extracted_name, extracted_id)
    例如: ('身份证', 'ocr_id_front', '正面', '姜默蒙', '130527199211020552')
    """
    engine = _get_ocr_engine()
    if engine is None:
        return None, "", "", "", ""

    target_input: str | bytes = str(file_path)
    if file_path.suffix.lower() == ".pdf":
        img_bytes = _extract_pdf_image_bytes(file_path)
        if not img_bytes:
            return None, "", "", "", ""
        target_input = img_bytes

    try:
        with _OCR_LOCK:
            result, _ = engine(target_input)
        if not result:
            return None, "", "", "", ""
        texts = [item[1] for item in result]
        full_text = " ".join(texts)
    except Exception:
        return None, "", "", "", ""

    extracted_name = ""
    extracted_id = ""
    id_match = re.search(r"\b\d{17}[\dxX]\b", full_text)
    if id_match:
        extracted_id = id_match.group(0)

    for idx, text in enumerate(texts):
        match = re.search(r"姓名(?:[/A-Za-z\s:：]*)([\u4e00-\u9fa5]{2,10})", text)
        if match:
            extracted_name = match.group(1).strip()
            break
        if re.search(r"^姓名(?:[/A-Za-z\s:：]*)$", text.strip()):
            if idx + 1 < len(texts) and re.match(r"^[\u4e00-\u9fa5]{2,10}$", texts[idx + 1].strip()):
                extracted_name = texts[idx + 1].strip()
                break

    if (
        "特种作业操作证" in full_text or "特种作业" in full_text
        or "特种设备作业" in full_text or "特种作业人员" in full_text
    ):
        return "特种证书", "ocr_special_cert", "", extracted_name, extracted_id
    if (
        "安全生产考核" in full_text or "安全考核合格" in full_text
        or "建安C证" in full_text or "建安A证" in full_text
        or "建安B证" in full_text or "安全员" in full_text
    ):
        return "安全员证", "ocr_safety_cert", "", extracted_name, extracted_id
    if (
        "劳动合同" in full_text or "用工合同" in full_text
        or ("甲方" in full_text and "乙方" in full_text and (
            "劳动" in full_text or "报酬" in full_text or "工作内容" in full_text
        ))
    ):
        return "劳动合同", "ocr_contract", "", extracted_name, extracted_id
    if (
        "毕业证书" in full_text or "学位证书" in full_text or "学信网" in full_text
        or "学历证书" in full_text or "普通高等学校" in full_text
    ):
        return "学历证明", "ocr_degree", "", extracted_name, extracted_id
    if (
        "机动车驾驶证" in full_text or "驾驶证" in full_text
        or "职业资格证书" in full_text or "职业资格" in full_text
    ):
        return "资格证书", "ocr_certificate", "", extracted_name, extracted_id
    if re.search(r"\d{16,19}", full_text) and (
        "银行" in full_text or "银联" in full_text or "Bank" in full_text
    ):
        return "银行卡", "ocr_bank", "", extracted_name, extracted_id
    if "居民身份证" in full_text or (
        "签发机关" in full_text and "有效期限" in full_text
        and "特种" not in full_text and "安全" not in full_text
    ):
        return "身份证", "ocr_id_back", "反面", extracted_name, extracted_id
    if (
        "公民身份号码" in full_text
        or ("姓名" in full_text and ("住址" in full_text or "民族" in full_text or "出生" in full_text))
        or extracted_id
    ):
        return "身份证", "ocr_id_front", "正面", extracted_name, extracted_id

    return None, "", "", extracted_name, extracted_id


def _classify_material_type(
    file_path: Path,
    filename: str,
    requested_types: list[str],
    *,
    employee_key: str = "",
    rel_path: str = "",
    cache: dict[str, Any] | None = None,
    use_cache: bool = True,
    cache_stats: dict[str, int] | None = None,
) -> tuple[str | None, str, str, str, str, bool]:
    """Classify file into a material type using filenames, document contents, and local OCR.

    Returns: (matched_material_type or None, match_method, subtype_label,
              extracted_name, extracted_id, cache_hit)
    cache_hit=True 表示本次分类结果来自 OCR 缓存命中。
    """
    stem = Path(filename).stem.lower()

    # 1. 优先匹配当前请求列表中明确包含在文件名里的材料（标准或自定义，按关键词长度降序最长优先匹配）
    sorted_req_types = sorted(
        [r for r in requested_types if r and r.strip()],
        key=lambda x: len(x),
        reverse=True,
    )
    for req_type in sorted_req_types:
        syns = MATERIAL_SYNONYMS.get(req_type, [req_type])
        # 按同义词长度降序，最长精确匹配优先（例如"保密协议"优先于"协议"）
        for syn in sorted(syns, key=lambda s: len(s), reverse=True):
            if syn.lower() in stem:
                sub = "正面" if "正面" in stem or "人像" in stem else ("反面" if "反面" in stem or "国徽" in stem else "")
                return req_type, "filename_keyword", sub, "", "", False

    # 2b. 全量同义词库匹配（按优先级互斥判断）
    for mat_type, synonyms in MATERIAL_SYNONYMS.items():
        for syn in sorted(synonyms, key=lambda s: len(s), reverse=True):
            if syn.lower() in stem:
                sub = "正面" if "正面" in stem or "人像" in stem else ("反面" if "反面" in stem or "国徽" in stem else "")
                return mat_type, "filename_keyword", sub, "", "", False

    # 3. 文档内部文本内容深度检索（针对文件名如 01.docx, file.pdf 等非标准命名）
    doc_text = _extract_document_text(file_path)
    if doc_text:
        for mat_type, content_keywords in _DOC_CONTENT_PATTERNS.items():
            for kw in content_keywords:
                if kw in doc_text:
                    return mat_type, "doc_content", "", "", "", False

    # 4. 本地离线 OCR 视觉图文识别（针对纯哈希/随机命名的图片或扫描版 PDF）
    ext = file_path.suffix.lower()
    if ext in IMAGE_EXTENSIONS or ext == ".pdf":
        # 4a. 先查缓存（基于文件二进制内容 SHA256 哈希指纹，改名或移动均能 100% 秒级命中）
        if use_cache and cache is not None:
            hit = _lookup_ocr_cache(cache, file_path, employee_key=employee_key, rel_path=rel_path)
            if hit is not None:
                if cache_stats is not None:
                    cache_stats["hits"] = cache_stats.get("hits", 0) + 1
                mat, method, sub, name, _id_hash = hit
                return mat, method, sub, name, "", True
            if cache_stats is not None:
                cache_stats["misses"] = cache_stats.get("misses", 0) + 1

        # 4b. 缓存未命中或不可用 → 真实 OCR
        ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id = _classify_by_ocr(file_path)
        if ocr_mat:
            # 4c. 仅在成功识别时回写缓存（以内容哈希为主键）
            if use_cache and cache is not None:
                _store_ocr_cache(cache, file_path,
                                 ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id,
                                 employee_key=employee_key, rel_path=rel_path)
            return ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id, False

    return None, "", "", "", "", False


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


@dataclass(frozen=True)
class _FlatIndexedFile:
    source_path: Path
    relative_path: str
    cache_key: str
    material_type: str
    match_method: str
    subtype: str
    extracted_names: tuple[str, ...]
    extracted_id_hash: str
    extracted_phone_hash: str = ""
    text_snippet: str = ""
    extracted_id_card: str = ""
    cache_hit: bool = False


def _scan_flat_library_files(lib_path: Path, skip_dir: Path | None = None) -> list[Path]:
    """递归扫描无序资料库中的可识别文件，跳过隐藏项、链接目录和输出目录。"""
    result: list[Path] = []
    for root, dirs, files in os.walk(lib_path, followlinks=False):
        root_path = Path(root)
        dirs[:] = [
            name
            for name in dirs
            if not name.startswith(".")
            and name.lower() not in _IGNORED_FILENAMES
            and not (root_path / name).is_symlink()
            and not (skip_dir and _is_path_nested(root_path / name, skip_dir))
        ]
        for filename in files:
            if _is_junk_or_temp_file(filename):
                continue
            source_path = root_path / filename
            if source_path.is_symlink() or not source_path.is_file():
                continue
            if source_path.suffix.lower() not in SUPPORTED_FILE_EXTENSIONS:
                continue
            result.append(source_path)
    return sorted(result, key=lambda path: path.relative_to(lib_path).as_posix().casefold())


def _extract_flat_document_text(file_path: Path) -> str:
    """提取无序库文档正文；该扩展能力仅用于 TASK-8，不改变原文件夹模式。"""
    ext = file_path.suffix.lower()
    if ext == ".docx":
        try:
            with zipfile.ZipFile(file_path) as archive:
                tree = ET.fromstring(archive.read("word/document.xml"))
                return " ".join(text for text in tree.itertext() if text)
        except Exception:
            return ""
    if ext in DOC_EXTENSIONS:
        return _extract_document_text(file_path)
    if ext == ".xlsx":
        try:
            workbook = load_workbook(file_path, read_only=True, data_only=True)
            chunks: list[str] = []
            character_count = 0
            try:
                for sheet in workbook.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        for value in row:
                            if value is not None:
                                text = str(value)
                                chunks.append(text)
                                character_count += len(text)
                                if character_count >= 20000:
                                    return " ".join(chunks)
            finally:
                workbook.close()
            return " ".join(chunks)
        except Exception:
            return ""
    if ext == ".xls":
        try:
            import xlrd

            workbook = xlrd.open_workbook(str(file_path), on_demand=True)
            chunks: list[str] = []
            character_count = 0
            try:
                for sheet in workbook.sheets():
                    for row_index in range(sheet.nrows):
                        for value in sheet.row_values(row_index):
                            if value not in (None, ""):
                                text = str(value)
                                chunks.append(text)
                                character_count += len(text)
                                if character_count >= 20000:
                                    return " ".join(chunks)
            finally:
                workbook.release_resources()
            return " ".join(chunks)
        except Exception:
            return ""
    if ext == ".pptx":
        try:
            with zipfile.ZipFile(file_path) as archive:
                xml_names = [
                    name for name in archive.namelist()
                    if name.endswith(".xml") and (name.startswith("ppt/slides/") or name == "word/document.xml")
                ]
                chunks: list[str] = []
                for name in xml_names:
                    tree = ET.fromstring(archive.read(name))
                    chunks.extend(text for text in tree.itertext() if text)
                return " ".join(chunks)
        except Exception:
            return ""
    return ""


def _classify_material_from_filename(filename: str, requested_types: list[str]) -> tuple[str | None, str]:
    """正文无法判型时，仅用文件名补充材料类型；绝不以文件名判断人员。"""
    stem = Path(filename).stem.casefold()
    types = list(dict.fromkeys([*requested_types, *MATERIAL_SYNONYMS.keys()]))
    weak_tokens = {"id", "a面", "b面", "正", "反", "照片", "证件"}
    for material in sorted(types, key=len, reverse=True):
        synonyms = MATERIAL_SYNONYMS.get(material, [material])
        for synonym in sorted(synonyms, key=len, reverse=True):
            token = synonym.strip().casefold()
            if not token or token in weak_tokens or len(token) < 2:
                continue
            if token in stem:
                return material, "filename_material_only"
    return None, ""


def _flat_cache_entry_usable(entry: Any) -> bool:
    if not isinstance(entry, dict):
        return False
    if entry.get("analysis_state") == "complete":
        return True
    # 兼容旧缓存：只有明确识别出人员的成功条目才可用于无序库，不能把旧空结果当负缓存。
    return bool(entry.get("material_type") and entry.get("extracted_name"))


def _flat_path_metadata(file_path: Path) -> tuple[int, int, int | None] | None:
    try:
        stat = file_path.stat()
    except OSError:
        return None
    return stat.st_size, stat.st_mtime_ns, _file_change_token(file_path, stat)


def _lookup_flat_index_cache(
    cache: dict[str, Any],
    file_path: Path,
    rel_path: str,
    cache_stats: dict[str, int],
) -> tuple[dict[str, Any] | None, str | None, tuple[int, float, str] | None]:
    """路径元数据先行、内容 hash 兜底；返回缓存条目、key 与指纹。"""
    entries: dict[str, Any] = cache.setdefault("entries", {})
    paths: dict[str, Any] = cache.setdefault("paths", {})
    metadata = _flat_path_metadata(file_path)
    if metadata is None:
        return None, None, None
    size, mtime_ns, change_token = metadata

    previous = paths.get(rel_path)
    if isinstance(previous, dict) and (
        change_token is not None
        and previous.get("source_size") == size
        and previous.get("source_mtime_ns") == mtime_ns
        and previous.get("source_change_token") == change_token
    ):
        previous_key = previous.get("cache_key")
        entry = entries.get(previous_key)
        if _flat_cache_entry_usable(entry):
            entry["verified_at"] = _beijing_now_str()
            cache_stats["hits"] = cache_stats.get("hits", 0) + 1
            return entry, str(previous_key), None

    fingerprint = _compute_full_file_fingerprint(file_path)
    if fingerprint is None:
        return None, None, None
    fp_size, mtime, sha = fingerprint
    cache_key = f"{sha}_{fp_size}"
    entry = entries.get(cache_key)
    if not _flat_cache_entry_usable(entry):
        legacy_key = f"{sha[:24]}_{fp_size}"
        legacy_entry = entries.get(legacy_key)
        if _flat_cache_entry_usable(legacy_entry):
            entry = dict(legacy_entry)
            entry["content_hash"] = sha
            entries[cache_key] = entry
            entries.pop(legacy_key, None)

    old_key = previous.get("cache_key") if isinstance(previous, dict) else None
    if old_key and old_key != cache_key:
        cache_stats["invalidated"] = cache_stats.get("invalidated", 0) + 1

    paths[rel_path] = {
        "cache_key": cache_key,
        "source_size": size,
        "source_mtime_ns": mtime_ns,
        "source_change_token": change_token,
    }
    if _flat_cache_entry_usable(entry):
        entry["verified_at"] = _beijing_now_str()
        cache_stats["hits"] = cache_stats.get("hits", 0) + 1
        return entry, cache_key, fingerprint

    cache_stats["misses"] = cache_stats.get("misses", 0) + 1
    return None, cache_key, fingerprint


def _store_flat_index_entry(
    cache: dict[str, Any],
    file_path: Path,
    rel_path: str,
    cache_key: str,
    fingerprint: tuple[int, float, str],
    *,
    material_type: str,
    match_method: str,
    subtype: str,
    extracted_names: list[str],
    extracted_id: str,
    extracted_text: str,
) -> dict[str, Any]:
    size, mtime, sha = fingerprint
    entry = {
        "content_hash": sha,
        "source_size": size,
        "source_mtime": mtime,
        "material_type": material_type or "其他材料",
        "match_method": match_method or "unrecognized",
        "subtype": subtype,
        "extracted_name": extracted_names[0] if extracted_names else "",
        "extracted_names": list(dict.fromkeys(extracted_names)),
        "extracted_id_hash": _hash_id_card(extracted_id),
        "extracted_phone_hash": _hash_phone(_extract_phone(extracted_text)),
        "ocr_text": _sanitize_cached_text(extracted_text),
        "verified_at": _beijing_now_str(),
        "sample_filename": file_path.name,
        "source_relpath": rel_path,
        "analysis_state": "complete",
        "index_scope": LIBRARY_MODE_FLAT_OCR,
    }
    cache.setdefault("entries", {})[cache_key] = entry
    return entry


def _analyze_flat_source(
    file_path: Path,
    requested_types: list[str],
) -> tuple[str, str, str, list[str], str, str, bool]:
    """分析单个无序库文件，返回类型、方式、子类型、姓名列表、证件号和正文。"""
    text = _extract_flat_document_text(file_path)
    method_prefix = "doc_content"
    analysis_complete = True
    if not text and (file_path.suffix.lower() in IMAGE_EXTENSIONS or file_path.suffix.lower() == ".pdf"):
        material, method, subtype, _name, extracted_id, text, names, analysis_complete = _analyze_ocr_file(
            file_path,
            requested_types=requested_types,
        )
    else:
        names = _extract_person_names(text)
        extracted_id = _extract_id_card(text)
        material, method, subtype = _classify_text_content(
            text,
            requested_types=requested_types,
            method_prefix=method_prefix,
        )

    if not material:
        material, filename_method = _classify_material_from_filename(file_path.name, requested_types)
        if material:
            method = filename_method if not method else f"{method}+{filename_method}"
    return material or "其他材料", method or "unrecognized", subtype, names, extracted_id, text, analysis_complete


def _entry_to_flat_indexed_file(
    file_path: Path,
    rel_path: str,
    cache_key: str,
    entry: dict[str, Any],
    *,
    cache_hit: bool,
    extracted_id_card: str = "",
) -> _FlatIndexedFile:
    raw_names = entry.get("extracted_names")
    if not isinstance(raw_names, list):
        raw_names = [entry.get("extracted_name")] if entry.get("extracted_name") else []
    names = tuple(
        str(name).strip() for name in raw_names
        if str(name or "").strip()
    )
    return _FlatIndexedFile(
        source_path=file_path,
        relative_path=rel_path,
        cache_key=cache_key,
        material_type=str(entry.get("material_type") or "其他材料"),
        match_method=str(entry.get("match_method") or "unrecognized"),
        subtype=str(entry.get("subtype") or ""),
        extracted_names=names,
        extracted_id_hash=str(entry.get("extracted_id_hash") or ""),
        extracted_phone_hash=str(entry.get("extracted_phone_hash") or ""),
        text_snippet=str(entry.get("ocr_text") or ""),
        extracted_id_card=extracted_id_card,
        cache_hit=cache_hit,
    )


def _build_flat_ocr_index(
    lib_path: Path,
    out_path: Path,
    requested_types: list[str],
    cache: dict[str, Any],
    cache_path: Path,
    cache_stats: dict[str, int],
    warnings: list[str],
    progress_callback: Callable[[int, int, str], None] | None,
) -> tuple[list[_FlatIndexedFile], bool]:
    """建立/复用全局 OCR 索引；小批次密集、超大批次分段持久化。"""
    source_files = [
        path for path in _scan_flat_library_files(lib_path, skip_dir=out_path)
        if path.resolve() != cache_path.resolve()
    ]
    indexed: list[_FlatIndexedFile] = []
    active_paths: set[str] = set()
    checkpoint_ok = True
    cache_changed_since_checkpoint = False

    for index, source_path in enumerate(source_files, start=1):
        rel_path = source_path.relative_to(lib_path).as_posix()
        active_paths.add(rel_path)
        if progress_callback:
            progress_callback(index, len(source_files), f"正在建立无序资料 OCR 索引：{index}/{len(source_files)}")

        entry, cache_key, fingerprint = _lookup_flat_index_cache(
            cache, source_path, rel_path, cache_stats,
        )
        if cache_key is None:
            warnings.append(f"无法读取资料文件，已跳过：{rel_path}")
            continue

        cache_hit = entry is not None
        extracted_id = ""
        if entry is None:
            material, method, subtype, names, extracted_id, text, analysis_complete = _analyze_flat_source(
                source_path, requested_types,
            )
            if fingerprint is None:
                fingerprint = _compute_full_file_fingerprint(source_path)
            if fingerprint is None:
                warnings.append(f"文件在索引过程中发生变化，已跳过：{rel_path}")
                continue
            if analysis_complete:
                entry = _store_flat_index_entry(
                    cache,
                    source_path,
                    rel_path,
                    cache_key,
                    fingerprint,
                    material_type=material,
                    match_method=method,
                    subtype=subtype,
                    extracted_names=names,
                    extracted_id=extracted_id,
                    extracted_text=text,
                )
                cache_changed_since_checkpoint = True
            else:
                # OCR 引擎不可用/文件暂时无法识别时不能写成永久负结果，下次查询必须重试。
                entry = {
                    "material_type": material,
                    "match_method": method,
                    "subtype": subtype,
                    "extracted_names": names,
                    "extracted_id_hash": _hash_id_card(extracted_id),
                    "analysis_state": "incomplete",
                }
                if not any("暂时无法完成 OCR" in warning for warning in warnings):
                    warnings.append(
                        "部分图片或扫描版 PDF 暂时无法完成 OCR，未写入负缓存；下次查询会自动重试。"
                    )
        elif entry.get("material_type") == "其他材料" and entry.get("ocr_text"):
            # 后续新增自定义材料时，直接用已缓存 OCR 文本重分类，不重新 OCR。
            material, method, subtype = _classify_text_content(
                str(entry.get("ocr_text") or ""),
                requested_types=requested_types,
                method_prefix="cached_text",
            )
            if material:
                entry["material_type"] = material
                entry["match_method"] = method
                entry["subtype"] = subtype
                cache_changed_since_checkpoint = True

        indexed.append(_entry_to_flat_indexed_file(
            source_path,
            rel_path,
            cache_key,
            entry,
            cache_hit=cache_hit,
            extracted_id_card=extracted_id,
        ))

        should_checkpoint = (
            (index <= 1000 and index % 100 == 0)
            or (index > 1000 and index % 1000 == 0)
        )
        if should_checkpoint and cache_changed_since_checkpoint:
            _trim_cache_by_age_and_size(cache)
            if not _save_ocr_cache(cache_path, cache):
                checkpoint_ok = False
            cache_changed_since_checkpoint = False

    # 移除已经删除/改名的旧路径；内容条目仅在没有任何现存路径引用时清理。
    paths: dict[str, Any] = cache.setdefault("paths", {})
    for rel_path in list(paths):
        if rel_path not in active_paths:
            paths.pop(rel_path, None)
            cache_changed_since_checkpoint = True
    referenced_keys = {
        item.get("cache_key") for item in paths.values()
        if isinstance(item, dict) and item.get("cache_key")
    }
    entries: dict[str, Any] = cache.setdefault("entries", {})
    for key, entry in list(entries.items()):
        if (
            isinstance(entry, dict)
            and entry.get("index_scope") == LIBRARY_MODE_FLAT_OCR
            and key not in referenced_keys
        ):
            entries.pop(key, None)
            cache_changed_since_checkpoint = True

    if cache_changed_since_checkpoint:
        _trim_cache_by_age_and_size(cache)
        if not _save_ocr_cache(cache_path, cache):
            checkpoint_ok = False
    return indexed, checkpoint_ok


def _flat_file_matches_employee(
    indexed_file: _FlatIndexedFile,
    employee: TargetEmployee,
    duplicate_target_names: set[str],
) -> bool:
    target_name = _normalize_person_name(employee.name)
    target_id_hash = _hash_id_card(employee.id_card)
    target_phone_hash = _hash_phone(employee.phone)
    if not target_id_hash and _ID_CARD_RE.fullmatch(employee.name.strip()):
        target_id_hash = _hash_id_card(employee.name.strip())

    if target_id_hash and indexed_file.extracted_id_hash:
        return target_id_hash == indexed_file.extracted_id_hash
    if target_phone_hash and indexed_file.extracted_phone_hash:
        return target_phone_hash == indexed_file.extracted_phone_hash
    if not target_name or target_name in duplicate_target_names:
        return False
    names = {_normalize_person_name(name) for name in indexed_file.extracted_names}
    if target_name in names:
        return True

    # 部分证书只印姓名本身而没有“姓名：”标签；以非中文边界做全文精确匹配，
    # 避免把“张三”错误命中“张三丰”。文件名始终不参与人员判断。
    raw_target = str(employee.name or "").strip()
    if raw_target and indexed_file.text_snippet:
        return bool(re.search(
            rf"(?<![\u3400-\u9fff]){re.escape(raw_target)}(?![\u3400-\u9fff])",
            indexed_file.text_snippet,
        ))
    return False


def _flat_employee_result_key(
    employee: TargetEmployee,
    duplicate_target_names: set[str],
) -> str:
    """同名人员的输出/缺件键使用非敏感限定符，防止文件互相覆盖。"""
    if _normalize_person_name(employee.name) not in duplicate_target_names:
        return employee.name
    if employee.employee_no:
        qualifier = f"工号{employee.employee_no}"
    elif employee.id_card:
        qualifier = f"证件尾号{employee.id_card[-4:]}"
    elif employee.phone:
        qualifier = f"手机尾号{employee.phone[-4:]}"
    else:
        qualifier = hashlib.sha256(employee.identity_key.encode("utf-8")).hexdigest()[:6]
    return f"{employee.name}（{qualifier}）"


def _unique_destination(path: Path) -> Path:
    if not path.exists():
        return path
    sequence = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{sequence}{path.suffix}")
        if not candidate.exists():
            return candidate
        sequence += 1


def _collect_from_flat_index(
    employee: TargetEmployee,
    indexed_files: list[_FlatIndexedFile],
    out_path: Path,
    mode: str,
    requested_materials: list[str] | None,
    matches: list[MaterialFileMatch],
    warnings: list[str],
    duplicate_target_names: set[str],
) -> tuple[list[str], int]:
    person_files = [
        item for item in indexed_files
        if _flat_file_matches_employee(item, employee, duplicate_target_names)
    ]
    person_file_count = len(person_files)
    if requested_materials is not None:
        person_files = [item for item in person_files if item.material_type in requested_materials]

    material_counts: dict[str, int] = {}
    for item in person_files:
        material_counts[item.material_type] = material_counts.get(item.material_type, 0) + 1
    material_sequences: dict[str, int] = {}
    copied_materials: set[str] = set()
    copied_count = 0
    clean_employee = safe_filename(
        _flat_employee_result_key(employee, duplicate_target_names)
    )

    for item in person_files:
        material = item.material_type or "其他材料"
        clean_material = safe_filename(material)
        material_sequences[material] = material_sequences.get(material, 0) + 1
        sequence = material_sequences[material]
        extension = item.source_path.suffix

        if requested_materials is None and mode == MODE_BY_EMPLOYEE:
            destination_dir = out_path / clean_employee
            target_name = item.source_path.name
        else:
            if item.subtype:
                suffix = f"_{item.subtype}"
                if material_counts[material] > 1 and sequence > 1:
                    suffix += f"_{sequence}"
            else:
                suffix = f"_{sequence}" if material_counts[material] > 1 else ""
            target_name = f"{clean_employee}_{clean_material}{suffix}{extension}"
            if mode == MODE_BY_EMPLOYEE:
                destination_dir = out_path / clean_employee
            elif mode == MODE_BY_MATERIAL:
                destination_dir = out_path / clean_material
            else:
                destination_dir = out_path

        destination_dir.mkdir(parents=True, exist_ok=True)
        destination = _unique_destination(destination_dir / target_name)
        current_fingerprint = _compute_full_file_fingerprint(item.source_path)
        if current_fingerprint is None:
            warnings.append(f"资料文件在检索后无法读取，已跳过：{item.relative_path}")
            continue
        current_size, _current_mtime, current_sha = current_fingerprint
        if f"{current_sha}_{current_size}" != item.cache_key:
            warnings.append(f"资料文件在索引后发生变化，已跳过并请重新运行：{item.relative_path}")
            continue
        try:
            shutil.copy2(item.source_path, destination)
        except Exception as exc:
            try:
                destination.unlink(missing_ok=True)
            except OSError:
                pass
            warnings.append(f"复制失败：{item.source_path} → {destination}: {exc}")
            continue

        copied_fingerprint = _compute_full_file_fingerprint(destination)
        if (
            copied_fingerprint is None
            or f"{copied_fingerprint[2]}_{copied_fingerprint[0]}" != item.cache_key
        ):
            try:
                destination.unlink(missing_ok=True)
            except OSError:
                pass
            warnings.append(f"资料文件复制后校验不一致，已移除结果并请重新运行：{item.relative_path}")
            continue

        matches.append(MaterialFileMatch(
            employee_name=employee.name,
            material_type=material,
            source_path=item.source_path,
            relative_source_path=item.relative_path,
            matched_by=item.match_method,
            target_filename=destination.name,
            target_path=destination,
            extracted_person_name=employee.name,
            extracted_id_card=item.extracted_id_card,
            cache_hit=item.cache_hit,
            employee_identity_key=employee.identity_key,
        ))
        copied_materials.add(material)
        copied_count += 1

    if requested_materials is None:
        return ([] if copied_count else ["（OCR 未识别到该人员资料）"], person_file_count)
    return (
        [material for material in requested_materials if material not in copied_materials],
        person_file_count,
    )


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
    library_mode: str = LIBRARY_MODE_PERSON_FOLDER,
    create_zip: bool = False,
    generate_report: bool = True,
    collect_all: bool = False,
    scan_depth: int = 1,
    progress_callback: Callable[[int, int, str], None] | None = None,
    use_ocr_cache: bool = True,
    ocr_cache_path: Path | str | None = None,
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
    if library_mode not in LIBRARY_MODES:
        raise ValueError(f"不支持的资料库形式：{library_mode}，可选值：{LIBRARY_MODES}")
    if library_mode == LIBRARY_MODE_FLAT_OCR and not use_ocr_cache:
        raise ValueError("无序平铺资料库必须启用 OCR 索引缓存，避免每次查询重复识别全部文件")

    employees = parse_employee_roster(roster_source)
    if not employees:
        raise ValueError("未能解析出有效的员工信息，请输入员工姓名/身份证，或上传员工名单表格")

    if material_types is None or len(material_types) == 0:
        global_materials = list(MATERIAL_SYNONYMS.keys())
    else:
        global_materials = list(material_types)

    out_path.mkdir(parents=True, exist_ok=True)

    # === OCR 智能索引缓存层：启动期加载 / 引擎升级全量失效 / 只读目录降级 ===
    ocr_cache: dict[str, Any] | None = None
    cache_stats: dict[str, int] = {"hits": 0, "misses": 0, "invalidated": 0}
    cache_path: Path | None = None
    cache_skipped_reason: str | None = None
    cache_write_ok = True
    current_signature = _get_engine_signature()

    if use_ocr_cache:
        cache_path = (
            Path(ocr_cache_path).expanduser().resolve()
            if ocr_cache_path is not None
            else (lib_path / _OCR_CACHE_FILE_NAME)
        )
        ocr_cache = _load_ocr_cache(cache_path)
        # 引擎升级 → 视为全量失效（彻底重写 entries）
        prev_sig = ocr_cache.get("engine_signature")
        if prev_sig and prev_sig != current_signature:
            ocr_cache["entries"] = {}
            ocr_cache["paths"] = {}
    else:
        cache_skipped_reason = "用户已关闭 OCR 识别缓存"

    total_steps = len(employees)
    matches: list[MaterialFileMatch] = []
    missing_records: dict[str, list[str]] = {}
    warnings: list[str] = []
    folder_match_counts: dict[str, int] = {}
    employee_result_keys: list[str] = []

    # 引擎升级警告：本轮第一次识别时输出一次即可
    if (
        use_ocr_cache
        and ocr_cache is not None
        and ocr_cache.get("engine_signature")
        and ocr_cache.get("engine_signature") != current_signature
    ):
        warnings.append(
            f"OCR 引擎版本变更（{ocr_cache.get('engine_signature')} → {current_signature}），"
            "缓存已全量失效，本次将重新 OCR 识别所有图片。"
        )

    folder_index: dict[str, list[Path]] = {}
    flat_index: list[_FlatIndexedFile] = []
    duplicate_target_names: set[str] = set()
    if library_mode == LIBRARY_MODE_PERSON_FOLDER:
        if progress_callback:
            progress_callback(0, len(employees), "正在扫描资料库文件夹索引...")
        folder_index = _scan_folder_index(lib_path, max_depth=scan_depth, skip_dir=out_path)
    else:
        assert ocr_cache is not None and cache_path is not None
        normalized_name_counts: dict[str, int] = {}
        for employee in employees:
            normalized_name = _normalize_person_name(employee.name)
            normalized_name_counts[normalized_name] = normalized_name_counts.get(normalized_name, 0) + 1
        duplicate_target_names = {
            name for name, count in normalized_name_counts.items() if name and count > 1
        }
        if duplicate_target_names:
            warnings.append(
                "名单中存在同名人员；无身份证号可核对的同名资料不会自动归属，请补充身份证号后重试。"
            )
        flat_index, checkpoint_ok = _build_flat_ocr_index(
            lib_path,
            out_path,
            list(dict.fromkeys([*global_materials, *MATERIAL_SYNONYMS.keys()])),
            ocr_cache,
            cache_path,
            cache_stats,
            warnings,
            progress_callback,
        )
        if not checkpoint_ok:
            cache_write_ok = False
            cache_skipped_reason = "资料库目录只读或无写入权限"
            warnings.append(
                f"OCR 索引缓存写入失败：{cache_path}；本次仍使用内存索引完成检索，下次会重新建立。"
            )

    for idx, emp in enumerate(employees):
        emp_key = (
            _flat_employee_result_key(emp, duplicate_target_names)
            if library_mode == LIBRARY_MODE_FLAT_OCR
            else emp.name
        )
        employee_result_keys.append(emp_key)
        employee_key = _build_employee_key(emp)
        if progress_callback:
            progress_callback(
                idx + 1, total_steps,
                f"[{idx + 1}/{total_steps}] 正在检索与匹配：{emp.name}"
                + (f"（缓存命中 {cache_stats['hits']}）" if cache_stats["hits"] else ""),
            )

        if collect_all:
            emp_materials: list[str] | None = None
        elif emp.per_person_materials:
            emp_materials = list(emp.per_person_materials)
        else:
            emp_materials = global_materials

        if library_mode == LIBRARY_MODE_FLAT_OCR:
            emp_missing, matched_count = _collect_from_flat_index(
                emp,
                flat_index,
                out_path,
                mode,
                emp_materials,
                matches,
                warnings,
                duplicate_target_names,
            )
            folder_match_counts[emp_key] = 1 if matched_count else 0
            if emp_missing:
                missing_records[emp_key] = emp_missing
            continue

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
                employee_key=employee_key,
                ocr_cache=ocr_cache,
                use_ocr_cache=use_ocr_cache,
                cache_stats=cache_stats,
            )
        else:
            emp_missing = _collect_specific_materials(
                emp, matched_folders, out_path, mode, emp_materials, matches, warnings,
                duplicate_folder_warning,
                employee_key=employee_key,
                ocr_cache=ocr_cache,
                use_ocr_cache=use_ocr_cache,
                cache_stats=cache_stats,
            )
            if emp_missing:
                missing_records[emp_key] = emp_missing

    # === OCR 缓存：跑完一轮后汇总写一次（而非每张图片即写）"""
    if (
        use_ocr_cache
        and ocr_cache is not None
        and cache_path is not None
        and (ocr_cache.get("entries") or library_mode == LIBRARY_MODE_FLAT_OCR)
    ):
        _trim_cache_by_age_and_size(ocr_cache)
        if not _save_ocr_cache(cache_path, ocr_cache):
            if cache_write_ok:
                warnings.append(
                    f"OCR 缓存写入失败：{cache_path}（资料库目录可能为只读），本次未持久化识别结果。"
                )
            cache_write_ok = False
            cache_skipped_reason = "资料库目录只读或无写入权限"

    # 缓存指标摘要写进 warnings
    if use_ocr_cache and cache_write_ok and ocr_cache is not None:
        total = cache_stats["hits"] + cache_stats["misses"]
        if total > 0:
            if library_mode == LIBRARY_MODE_FLAT_OCR:
                warnings.append(
                    f"无序资料 OCR 索引：复用 {cache_stats['hits']} 个，新增识别 {cache_stats['misses']} 个"
                    + (f"，缓存文件：{cache_path}" if cache_path else "")
                )
            else:
                warnings.append(
                    f"OCR 智能索引缓存：命中 {cache_stats['hits']} 次，跳过实时识别 {cache_stats['misses']} 次"
                    + (f"，缓存文件：{cache_path}" if cache_path else "")
                )

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
        _write_excel_report(
            report_path, employees, report_materials, matches, collect_all, warnings,
            cache_stats=cache_stats, cache_path=cache_path,
        )

    return MaterialCollectResult(
        library_dir=lib_path,
        output_dir=out_path,
        zip_path=zip_path,
        report_path=report_path,
        mode=mode,
        library_mode=library_mode,
        target_employees=employees,
        requested_materials=report_materials,
        matches=matches,
        missing_records=missing_records,
        warnings=warnings,
        folder_match_counts=folder_match_counts,
        employee_result_keys=employee_result_keys,
        ocr_cache_enabled=use_ocr_cache,
        ocr_cache_hits=cache_stats["hits"],
        ocr_cache_misses=cache_stats["misses"],
        ocr_cache_invalidated=cache_stats["invalidated"],
        ocr_cache_path=str(cache_path) if (use_ocr_cache and cache_write_ok and cache_path) else None,
        ocr_cache_skipped_reason=cache_skipped_reason,
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
    *,
    employee_key: str = "",
    ocr_cache: dict[str, Any] | None = None,
    use_ocr_cache: bool = True,
    cache_stats: dict[str, int] | None = None,
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

                    # 尝试轻量分析图片是否有信息不匹配（优先查缓存）
                    ocr_name, ocr_id = "", ""
                    cache_hit = False
                    if src.suffix.lower() in IMAGE_EXTENSIONS:
                        cached = None
                        if use_ocr_cache and ocr_cache is not None:
                            cached = _lookup_ocr_cache(
                                ocr_cache, src,
                                employee_key=employee_key, rel_path=rel_p,
                            )
                        if cached is not None:
                            _, _, _, ocr_name, _ = cached
                            cache_hit = True
                            if cache_stats is not None:
                                cache_stats["hits"] += 1
                        else:
                            ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id = _classify_by_ocr(src)
                            if cache_stats is not None:
                                cache_stats["misses"] += 1
                            if use_ocr_cache and ocr_cache is not None and ocr_mat:
                                _store_ocr_cache(
                                    ocr_cache, src,
                                    ocr_mat, ocr_method, ocr_sub, ocr_name, ocr_id,
                                    employee_key=employee_key, rel_path=rel_p,
                                )
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
                        cache_hit=cache_hit,
                    ))
        except Exception as e:
            warnings.append(f"无法访问文件夹 {folder_path}: {e}")


def _score_file_candidate(
    filename: str,
    requested_materials: list[str],
) -> int:
    """计算文件对当前请求材料的相关性优先级评分（分数越高越优先做 OCR/内容识别）。

    100分: 文件名明确包含请求的材料名称或同义词（如"身份证"、"特种作业"、"安全员"、"劳动合同"）
     80分: 文件名包含编号特征线索（如 T+身份证号、A/B/C+编号、纯身份证号等高疑似文件名）
     50分: 随机/乱码命名的图片、扫描版 PDF 或普通文件
     10分: 文件名明确属于其他【未请求】的材料类型（降级到最后兜底）
    """
    stem = Path(filename).stem.lower()

    # 1. 文件名直接命中当前请求材料或其同义词 -> 100分
    for req_type in requested_materials:
        if not req_type:
            continue
        syns = MATERIAL_SYNONYMS.get(req_type, [req_type])
        for syn in syns:
            if syn.lower() in stem:
                return 100

    # 2. 文件名包含线索编号特征 -> 80分
    if "特种证书" in requested_materials or "资格证书" in requested_materials:
        if re.search(r"(?:^|_)t\d{17}[\dxX]|(?:^|_)t\d{15}", stem):
            return 80
    if "安全员证" in requested_materials:
        if re.search(r"(?:^|_)[abc]\d{17}[\dxX]|(?:^|_)[abc]\d{15}", stem):
            return 80
    if "身份证" in requested_materials:
        if re.search(r"(?:^|[^a-z0-9])\d{17}[\dxX](?:[^a-z0-9]|$)|(?:^|[^a-z0-9])\d{15}(?:[^a-z0-9]|$)", stem):
            return 80

    # 3. 检查是否明确包含其他【未请求】材料的同义词 -> 10分
    for other_mat, other_syns in MATERIAL_SYNONYMS.items():
        if other_mat not in requested_materials:
            for s in other_syns:
                if s.lower() in stem:
                    return 10

    # 4. 其他普通文件（随机命名图片/PDF等） -> 50分
    return 50


def _is_all_requested_materials_satisfied(
    found: dict[str, list[Any]],
    requested_materials: list[str],
) -> bool:
    """检查当前员工所需材料是否已全部找齐（用于触发短路早停，跳过后续无谓 OCR）。"""
    for mat in requested_materials:
        items = found.get(mat)
        if not items:
            return False
        # 如果是身份证且区分正反面，若只有单侧（只有正面无反面，或只有反面无正面），不算完全找齐，继续找另一侧
        if mat == "身份证":
            subtypes = {it[3] for it in items if len(it) > 3}
            if "" not in subtypes:
                if ("正面" in subtypes and "反面" not in subtypes) or ("反面" in subtypes and "正面" not in subtypes):
                    return False
    return True


def _collect_specific_materials(
    emp: TargetEmployee,
    matched_folders: list[tuple[Path, str]],
    out_path: Path,
    mode: str,
    requested_materials: list[str],
    matches: list[MaterialFileMatch],
    warnings: list[str],
    duplicate_warning: str = "",
    *,
    employee_key: str = "",
    ocr_cache: dict[str, Any] | None = None,
    use_ocr_cache: bool = True,
    cache_stats: dict[str, int] | None = None,
) -> list[str]:
    """指定材料模式：在匹配到的文件夹中精准搜集对应材料类型的文件。

    使用 启发式优先级排序 + 文件名特征 + 文档内容检索 + 离线视觉 OCR 进行识别：
    1. 优先将高疑似度的文件（如明确命名或含证件编号特征的文件）排在前面进行 OCR / 正文识别；
    2. 一旦目标员工所需材料全部找齐，且无后续高置信度同名候选文件，立即短路早停，跳过后续无谓识别，极速省时；
    3. 若前序文件未找齐，绝不遗漏，自动兜底继续逐个识别后续所有文件，直到找齐或扫描完毕。
    """
    clean_emp = safe_filename(emp.name)
    found: dict[str, list[tuple[Path, str, str, str, str, str, bool]]] = {m: [] for m in requested_materials}
    seen_hashes: set[tuple[int, str]] = set()

    # 1. 扫描匹配到的所有文件夹，收集所有候选文件
    raw_candidates: list[tuple[Path, str, str]] = []
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
                    try:
                        rel_p = str(f_path.relative_to(folder_path.parent))
                    except ValueError:
                        rel_p = f_path.name
                    raw_candidates.append((f_path, rel_p, folder_reason))
        except Exception as e:
            warnings.append(f"无法访问文件夹 {folder_path}: {e}")

    # 2. 按线索优先级评分降序排序（高疑似度文件排在最前面优先做 OCR/内容识别）
    scored_candidates: list[tuple[int, Path, str, str]] = [
        (_score_file_candidate(f_path.name, requested_materials), f_path, rel_p, folder_reason)
        for f_path, rel_p, folder_reason in raw_candidates
    ]
    scored_candidates.sort(key=lambda item: item[0], reverse=True)

    # 3. 按优先级顺序逐个进行精准识别（支持短路早停）
    for idx_cand, (cand_score, f_path, rel_p, folder_reason) in enumerate(scored_candidates):
        sig = _get_file_signature(f_path)
        if sig in seen_hashes:
            continue

        doc_hint = _build_doc_format_hint(f_path)
        try:
            classified_mat_type, match_method, subtype, ocr_name, ocr_id, cache_hit = _classify_material_type(
                f_path, f_path.name, requested_materials,
                employee_key=employee_key,
                rel_path=rel_p,
                cache=ocr_cache,
                use_cache=use_ocr_cache,
                cache_stats=cache_stats,
            )
        except Exception as exc:
            warnings.append(f"文件读取异常 {f_path.name}: {exc}")
            if doc_hint:
                warnings.append(doc_hint)
            matches.append(MaterialFileMatch(
                employee_name=emp.name,
                material_type="未知",
                source_path=f_path,
                relative_source_path=rel_p,
                matched_by="读取失败",
                target_filename=f_path.name,
                mismatch_warning=f"⚠️ 文件读取损坏或异常: {exc}",
            ))
            continue

        if classified_mat_type and classified_mat_type in requested_materials:
            if not any(existing[0] == f_path for existing in found[classified_mat_type]):
                seen_hashes.add(sig)
                found[classified_mat_type].append(
                    (f_path, rel_p, match_method or folder_reason, subtype, ocr_name, ocr_id, cache_hit)
                )

        # 4. 短路早停：如果所有请求的材料都已经找齐，且后续没有高置信度同名候选文件（如多页合同/多个证书），立即停止扫描后续文件！
        if _is_all_requested_materials_satisfied(found, requested_materials):
            next_score = scored_candidates[idx_cand + 1][0] if idx_cand + 1 < len(scored_candidates) else 0
            if next_score < 80:
                break

    # 复制匹配到的真实文件到输出目录
    missing_list: list[str] = []
    for mat_type in requested_materials:
        m_list = found[mat_type]
        if not m_list:
            missing_list.append(mat_type)
            continue

        for seq, (src_path, rel_p, match_reason, subtype, ocr_name, ocr_id, cache_hit) in enumerate(m_list, start=1):
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
                cache_hit=cache_hit,
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
    *,
    cache_stats: dict[str, int] | None = None,
    cache_path: Path | None = None,
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
    cache_hit_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    cache_hit_font = Font(name="微软雅黑", size=10, color="1F4E79")

    ws_summary["A1"] = "员工资料提取汇总与缺失清单"
    ws_summary["A1"].font = title_font
    ws_summary["A1"].alignment = Alignment(vertical="center")

    emp_file_counts: dict[str, int] = {}
    emp_cache_hits: dict[str, int] = {}
    emp_mismatch_warnings: dict[str, list[str]] = {}
    uses_identity_keys = any(match.employee_identity_key for match in all_matches)

    def _report_match_key(match: MaterialFileMatch) -> str:
        if uses_identity_keys and match.employee_identity_key:
            return match.employee_identity_key
        return match.employee_name

    def _report_employee_key(employee: TargetEmployee) -> str:
        return employee.identity_key if uses_identity_keys else employee.name

    for m in all_matches:
        match_key = _report_match_key(m)
        emp_file_counts[match_key] = emp_file_counts.get(match_key, 0) + 1
        if m.cache_hit:
            emp_cache_hits[match_key] = emp_cache_hits.get(match_key, 0) + 1
        if m.mismatch_warning:
            emp_mismatch_warnings.setdefault(match_key, []).append(m.mismatch_warning)

    total_emp = len(employees)
    total_files = len(all_matches)
    found_emp = sum(
        1 for emp in employees if emp_file_counts.get(_report_employee_key(emp), 0) > 0
    )
    not_found_emp = total_emp - found_emp

    # OCR 缓存指标
    hits = cache_stats.get("hits", 0) if cache_stats else 0
    misses = cache_stats.get("misses", 0) if cache_stats else 0
    cache_total = hits + misses
    cache_summary = (
        f"{hits}/{cache_total}"
        if cache_total > 0
        else "-"
    )

    ws_summary["A3"] = "统计概要"
    ws_summary["A3"].font = Font(name="微软雅黑", size=11, bold=True)

    stats: list[tuple[str, Any, str, Any]] = [
        ("目标员工总数", total_emp, "已提取文件总数", total_files),
        ("已找到员工数", found_emp, "未找到员工数", not_found_emp),
    ]
    # 仅当启用了缓存且有数据时附加缓存统计行
    if cache_stats is not None:
        stats.append(("OCR 缓存命中", hits, "OCR 实时识别", misses))

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
        headers = ["序号", "员工姓名", "身份证号码", "提取状态", "提取文件数", "OCR 缓存命中", "信息核对预警 / 备注"]
    else:
        headers = ["序号", "员工姓名", "身份证号码", "提取进度", "OCR 缓存命中"] + requested_materials + ["信息核对预警 / 备注"]

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

        employee_report_key = _report_employee_key(emp)
        file_count = emp_file_counts.get(employee_report_key, 0)
        warn_list = emp_mismatch_warnings.get(employee_report_key, [])
        emp_hits = emp_cache_hits.get(employee_report_key, 0)

        if collect_all:
            status_cell = ws_summary.cell(current_r, 4)
            count_cell = ws_summary.cell(current_r, 5)
            cache_cell = ws_summary.cell(current_r, 6)
            warn_cell = ws_summary.cell(current_r, 7)
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

            cache_cell.value = f"{emp_hits}/{file_count}" if file_count > 0 else "-"
            cache_cell.fill = cache_hit_fill if emp_hits > 0 else stat_fill
            cache_cell.font = cache_hit_font if emp_hits > 0 else normal_font
            cache_cell.alignment = Alignment(horizontal="center", vertical="center")

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
                if _report_match_key(m) == employee_report_key and m.material_type in requested_materials:
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

            # OCR 缓存命中列
            cache_cell = ws_summary.cell(current_r, 5)
            cache_cell.value = f"{emp_hits}/{file_count}" if file_count > 0 else "-"
            cache_cell.fill = cache_hit_fill if emp_hits > 0 else stat_fill
            cache_cell.font = cache_hit_font if emp_hits > 0 else normal_font
            cache_cell.alignment = Alignment(horizontal="center", vertical="center")

            for col_offset, mat_type in enumerate(requested_materials, start=6):
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

            warn_col_idx = 6 + len(requested_materials)
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
        elif "缓存命中" in col_name:
            ws_summary.column_dimensions[col_letter].width = 14
        else:
            max_len = max(len(str(cell.value or "")) for cell in col if cell.row >= start_row)
            ws_summary.column_dimensions[col_letter].width = max(min(max_len * 2 + 2, 28), 12)

    # Sheet 2: Matched File List
    ws_files = wb.create_sheet(title="提取文件明细清单")
    file_headers = ["序号", "员工姓名", "材料类型", "目标文件名", "证件识别姓名", "证件识别号码", "信息匹配校验", "匹配依据", "缓存命中", "原始文件路径"]
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

        # 缓存命中列
        hit_cell = ws_files.cell(r, 9)
        if match_item.cache_hit:
            hit_cell.value = "✓ 命中"
            hit_cell.fill = cache_hit_fill
            hit_cell.font = cache_hit_font
        elif match_item.matched_by and "ocr" in match_item.matched_by.lower():
            hit_cell.value = "✗ 实时"
            hit_cell.font = normal_font
        else:
            hit_cell.value = "-"
            hit_cell.font = normal_font
        hit_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws_files.cell(r, 10, match_item.relative_source_path).alignment = Alignment(horizontal="left", vertical="center")

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
        elif "缓存命中" in col_name:
            ws_files.column_dimensions[col_letter].width = 12
        else:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws_files.column_dimensions[col_letter].width = max(min(max_len * 2 + 2, 30), 12)

    # Sheet 3: OCR 缓存指标（如果有缓存数据）
    if cache_stats is not None:
        ws_cache = wb.create_sheet(title="OCR 缓存指标")
        ws_cache["A1"] = "OCR 智能索引缓存指标"
        ws_cache["A1"].font = title_font
        ws_cache["A2"] = "缓存文件"
        ws_cache["B2"] = str(cache_path) if cache_path else "-"
        ws_cache["A3"] = "命中次数"
        ws_cache["B3"] = hits
        ws_cache["A4"] = "实时识别次数"
        ws_cache["B4"] = misses
        ws_cache["A5"] = "命中率"
        ws_cache["B5"] = f"{hits / cache_total * 100:.1f}%" if cache_total > 0 else "-"
        ws_cache["A6"] = "失效次数"
        ws_cache["B6"] = cache_stats.get("invalidated", 0)

        for r in range(1, 7):
            ws_cache.cell(r, 1).font = Font(name="微软雅黑", size=10, bold=True)
            ws_cache.cell(r, 1).fill = stat_fill
            ws_cache.cell(r, 1).alignment = Alignment(horizontal="right", vertical="center")
            ws_cache.cell(r, 1).border = thin_border
            ws_cache.cell(r, 2).font = normal_font
            ws_cache.cell(r, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_cache.cell(r, 2).border = thin_border

        ws_cache.column_dimensions["A"].width = 18
        ws_cache.column_dimensions["B"].width = 60

    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(report_path)
    wb.close()
