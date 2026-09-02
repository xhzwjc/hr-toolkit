"""Shared tutorial navigation and copy for both desktop renderers."""

from __future__ import annotations

from typing import Optional, Tuple

from .desktop_contract import NAV_GROUPS, TOOL_NAV_ITEMS


TutorialLine = Tuple[str, Optional[str]]
TutorialEntry = Tuple[str, Optional[str], str]


def tutorial_entries() -> list[TutorialEntry]:
    """Return the navigation entries used by the original Tk tutorial."""

    entries: list[TutorialEntry] = []
    for tool_id, label in TOOL_NAV_ITEMS:
        if tool_id == "personnel_change_merge":
            entries.append((tool_id, "merge", "异动表汇总"))
            entries.append((tool_id, "roster", "花名册更新"))
        elif tool_id == "archive_import":
            entries.append((tool_id, "import", "档案入库"))
            entries.append((tool_id, "export", "档案表生成"))
        else:
            entries.append((tool_id, None, label))
    return entries


def tutorial_groups() -> list[dict[str, object]]:
    """Return grouped, Qt-friendly navigation with the exact shared copy."""

    entries = tutorial_entries()
    grouped: list[dict[str, object]] = []
    for group_name, tool_ids in NAV_GROUPS:
        items = []
        for tool_id, mode, label in entries:
            if tool_id not in tool_ids:
                continue
            items.append(
                {
                    "toolId": tool_id,
                    "mode": mode or "",
                    "label": label,
                    "lines": [
                        {"text": text, "style": style or ""}
                        for text, style in tutorial_lines(tool_id, mode)
                    ],
                }
            )
        if items:
            grouped.append({"name": group_name, "items": items})
    return grouped


def tutorial_lines(tool_id: str, mode: str | None = None) -> list[TutorialLine]:
    """Return the tutorial text formerly owned only by the Tk renderer."""

    if tool_id == "social_security":
        return [
            ("适用：把各社保账户缴费清单整理成社保明细表和社保汇总表。", "strong"),
            ("步骤：选择单个缴费清单、多个清单、常见压缩包，或包含清单的文件夹；再选择参保人员花名册。", None),
            ("结果：生成“社保明细表.xlsx”和“社保汇总表.xlsx”，汇总表里含基础数据分析和异常提醒。", None),
            ("目前规则：按身份证关联花名册；费用所属期优先读取明细行和原文件名，文件夹或压缩包名称只辅助识别缴纳地和缴纳单位。", None),
            ("注意：公积金、残保金、管理费暂无数据时留空；账单识别结果与花名册不一致时会提醒。", "warning"),
        ]
    if tool_id == "data_statistics":
        return [
            ("适用：把 HR 系统导出的考勤结果、周报记录、月报记录自动整理成统计表。", "strong"),
            ("步骤：选择单个文件、多个文件、常见压缩包，或包含这些文件的文件夹。", None),
            ("如需统计未写周报/月报，请选择“应汇报人员名单”；不选时只能按文件中出现过的人推断。", None),
            ("周报统计日期（可选）：填写如 2026-06-02 至 2026-06-30，只统计范围内周一截止的周报；留空按整月统计。适合 1 号正好是周一的月份，避免把上月最后一周重复统计。", None),
            ("结果：生成“考勤周月报汇总表.xlsx”，包含考勤统计、周月报统计、考勤异常明细、周月报异常明细。", None),
            ("当前规则：考勤公司默认“总部”；周报截止次周一17:00，周二至周四补交算上一期超时（备注写明提交时间），周五起交的算下一期；月报按次月2日17:01及以后算超时。", None),
            ("容易疑惑1：如果某人上一期已经交过周报，周二到周四又交了一份，这份算他提前交的下一期，不记超时，下一期也不会记未写。", None),
            ("容易疑惑2：选了统计日期时，归属期超出范围的周报本次不统计、留给下一次。比如范围选到6.24，6.26（周五）交的属于6.29截止那期，本次不会出现。", None),
            ("注意：周月报异常只统计次数和明细，不计算扣款金额。", "warning"),
        ]
    if tool_id == "insurance_ledger":
        return [
            ("适用：把各保单人员清单整理成保险台账，并根据需求6的人力资源分析表做增减预警。", "strong"),
            ("步骤：选择单个保单清单、多个清单、常见压缩包，或包含清单的文件夹；再选择人力资源分析表。", None),
            ("结果：生成“保险台账.xlsx”，包含保险台账和人员增减预警两个工作表。", None),
            ("当前规则：PZDX保额取“每人伤残死亡限额”，按万元显示；PEAC保额固定按60万元。", None),
            ("注意：人力资源分析表需包含“花名册”工作表；花名册在职但保单没有会提示需加保，保单有但花名册没有或已标记离职会提示需减保。", "warning"),
        ]
    if tool_id == "salary_merge":
        return [
            ("适用：把 1-12 个月工资表合成一张个人应发工资汇总表。", "strong"),
            ("步骤：可选择单个月度工资表、多个工资表、常见压缩包，或包含这些文件的文件夹。", None),
            ("如已有前几月汇总表，再选择“已有汇总表”；不选则新建一张汇总表。", None),
            ("点击“开始合并”后，上传资料和结果会自动保存到当前工作项目。", None),
            ("结果：按姓名、身份证号、月份合并；没有工资的月份填 0；已存在的人员月份不会覆盖。", None),
            ("注意：工资表文件名或表内日期要能识别月份；重复人员或重复月份会在执行结果里提醒。", "warning"),
        ]
    if tool_id == "personnel_change_merge":
        if mode == "roster":
            return [
                ("适用：已有月度异动汇总表时，单独更新人力资源花名册。", "strong"),
                ("步骤：选择单个异动汇总表、多个汇总表，或包含汇总表的文件夹；再选择人力资源花名册。", None),
                ("点击“更新花名册”后，上传资料和结果会自动保存到当前工作项目。", None),
                ("结果：根据汇总表里的增员写入花名册，根据减员在花名册中标记离职。", None),
                ("注意：不会清空原花名册；身份证已存在的增员不会重复写入，找不到的减员会在日志提醒。", "warning"),
            ]
        return [
            ("适用：把项目异动表按记录日期分到对应月份汇总表。", "strong"),
            ("步骤：可选择单个异动表、多个异动表、常见压缩包，或包含这些文件的文件夹。", None),
            ("如已有月度汇总表，可选择单个汇总表或包含多个汇总表的文件夹；工具会按月份追加，原有记录不会清空。", None),
            ("不选择已有汇总表时，工具会按月份新建干净汇总表。缺少某个月份汇总表时也会自动创建。", None),
            ("如果同一文件夹里放了人力资源分析表，工具会自动更新其中的花名册。", None),
            ("点击“开始汇总”后，上传资料和结果会自动保存到当前工作项目。", None),
            ("月份规则：增员看入职日期，减员看离职日期，转正看转正日期，调动看调整日期。", None),
            ("注意：只处理增补表、离职、转正、调整；薪酬、产值和同行对比分析暂不处理。", "warning"),
        ]
    if tool_id == "archive_import":
        if mode == "export":
            return [
                ("适用：把一个或多个档案汇总表写入各公司独立档案表。", "strong"),
                ("步骤：选择档案汇总表文件、多个文件、常见压缩包，或包含汇总表的文件夹。", None),
                ("如已有某个公司的档案表，可选择文件、常见压缩包或文件夹；不选或没匹配到时会按内置干净模板新建。", None),
                ("结果：按公司生成独立 Excel；已有身份证不重复新增，只补充空白字段。", None),
                ("注意：公司档案表会自动改公司名，新增行会补边框、居中和公式。", "warning"),
            ]
        return [
            ("适用：把项目部提交的人事档案移交表写入公司档案汇总表。", "strong"),
            ("步骤：可选择单个移交表、多个移交表、常见压缩包，或包含这些文件的文件夹。", None),
            ("已有档案汇总表可不选；不选时工具会用内置空模板新建一份汇总表。", None),
            ("结果：按“公司”写入对应工作表；身份证已存在时不重复新增，只补充空白材料字段。", None),
            ("注意：编号会从文件名或表头标题识别项目地区，如“茂名项目部”自动填 11；识别不到会留空并提醒。", "warning"),
        ]
    if tool_id == "folder_rename":
        return [
            ("适用：批量修改所选目录下第一层文件夹或文件名称。", "strong"),
            ("按 Excel 人名顺序批量重命名：名单按姓名行顺序，项目按文件名顺序一一对应；文件保留原扩展名。", None),
            ("数量不一致、姓名无效、目标重名或目标已存在时会在预览中明确提醒；未配对或冲突项目不会改名，也不会覆盖。", "warning"),
            ("追加文字：姓名不填就是全部项目追加；填姓名就是只处理这个人。输入内容会原样追加，需要分隔符时请一并输入。", None),
            ("删除结尾文字：输入“_劳动合同”，可删除“张三_劳动合同 / 张三-劳动合同 / 张三劳动合同”的结尾文字。", None),
            ("修改单人名称：填写原姓名和新名称，例如“张三”改为“章五”。", None),
            ("安全说明：确认后会先把所选文件夹复制进当前项目，再在“处理结果”的副本上改名；电脑上的原文件夹不会被修改。", "warning"),
        ]
    if tool_id == "salary_split":
        return [
            ("适用：一个完整工资表按“入职公司”拆成多个公司工资表。", "strong"),
            ("步骤：先打开工作项目，选择工资表文件，再点击“开始拆分”。", None),
            ("点击“打开所在文件夹”可直接查看本次生成的结果目录。", None),
            ("结果：每个入职公司生成一个 Excel，保留表头、格式、公式、小计和底部总计。", None),
            ("注意：源工资表不会被修改；如果模板列名或表结构变化，先发给开发确认。", "warning"),
        ]
    if tool_id == "material_collector":
        return [
            ("适用：根据员工名单从资料库批量提取特定材料（身份证、合同、学历等）并自动打包。", "strong"),
            ("步骤：选择员工资料库根目录，在第二行选择员工名单表格（Excel），勾选需要的材料类型后点击“开始打包”。", None),
            ("资料库形式：已有姓名文件夹选“原模式”；文件无序混放时选“OCR 索引”，首次建立索引后会复用未变化文件。", None),
            ("归类方式：支持“按员工归类”（每人建一个文件夹）、“按材料归类”或“平铺输出”，可选自动生成 ZIP 压缩包。", None),
            ("结果：按指定结构导出文件，并自动生成《员工资料提取汇总与缺失清单.xlsx》。", None),
            ("安全说明：纯本地处理、不上传外网；源文件不会修改，OCR 索引模式只会新增隐藏缓存文件。", "warning"),
        ]
    return [
        ("该工具暂未实现。", "strong"),
        ("请选择左侧已完成的工具：需求1、需求2、需求4、需求5、需求6、需求7、需求8、需求9。", None),
    ]


__all__ = ["tutorial_entries", "tutorial_groups", "tutorial_lines"]
