# 《HR 工具箱技术架构与深度实现方案》

> **文档性质**：本文档为基于当前 `HR Toolkit` 项目真实源码、工程结构、构建流水线与业务逻辑深度反编译与提炼的标准技术架构方案。可直接作为新部门（如财务、运营、行政、供应链等）开发同类型本地自动化工具箱的技术方案、架构规范与工程实施蓝本。

---

## 1. 项目概述

*   **项目定位**：一款跨平台（Windows x64 / macOS）的**纯本地桌面端批量数据与文件自动化处理工具箱**。
*   **核心解决场景**：面向企业职能部门（如人事、财务、运营等），解决多源异构 Excel 报表合并与拆分、历史台账比对、周月报与考勤数据清洗统计、扫描件/图片 OCR 提取与自动归档、人员资料批量规范化重命名等重复度高、易出错、数据量大且格式敏感的业务场景。
*   **系统运行方式**：
    *   **纯本地单体执行**：无需部署中心服务器，无需依赖外部云端 API，双击直接运行。
    *   **独立项目工作区机制（Workspace Sandbox）**：采用“工作区隔离”代替传统的“原地文件修改”，保障原始资料 100% 不被污染或误删。
    *   **CLI / GUI 双模运行**：核心业务代码完全解耦，既可以通过图形界面（Tkinter）交互操作，也可以通过命令行（CLI）在脚本系统或 CI 环境中无头（Headless）执行。

---

## 2. 总体技术架构

### 2.1 架构本质说明
*   **【当前项目实际实现】**：本项目采用经典的**分层单体桌面应用架构（Desktop Layered Monolith）**。
*   **【明确边界】**：
    *   **无前后端分离**：不存在 Web 前端（Vue/React）与后端（FastAPI/Spring Boot）的划分，不监听任何本地 HTTP/TCP 端口。
    *   **无网络数据库**：不依赖 MySQL/PostgreSQL。仅在本地通过 SQLite（记录元数据历史索引）和 JSON（项目清单 Manifest）进行持久化，配合 OS 文件级互斥锁保证多实例并发安全。
    *   **无外部云调用**：所有 Excel 解析、数据计算、OCR 识别均在本地 CPU/内存中完成。

### 2.2 系统分层架构图

```text
┌─────────────────────────────────────────────────────────────────────────┐
│                              表现层 (Presentation)                       │
│  ┌─────────────────────────────────┐   ┌──────────────────────────────┐ │
│  │   GUI 界面 (gui.py)             │   │   CLI 命令行 (cli.py)        │ │
│  │   - Tkinter / ttk 三栏工作台     │   │   - argparse 子命令解析      │ │
│  │   - 响应式抽屉 / 进度弹窗       │   │   - JSON / 控制台格式化输出   │ │
│  └─────────────────────────────────┘   └──────────────────────────────┘ │
├─────────────────────────────────────────────────────────────────────────┤
│                             调度与异步层 (Async & Dispatch)             │
│  - threading.Thread (后台守护工作线程，杜绝 UI 假死)                     │
│  - queue.Queue / root.after() (线程安全状态传递与 UI 刷新)               │
├─────────────────────────────────────────────────────────────────────────┤
│                             业务逻辑层 (Business Tools)                 │
│  - tools/social_security.py (社保明细与汇总)                            │
│  - tools/data_statistics.py (考勤与周月报统计)                          │
│  - tools/salary_split.py (工资表按公司拆分)                             │
│  - tools/salary_merge.py (多月薪资汇总)                                 │
│  - tools/personnel_change_merge.py (异动表汇总与花名册同步)              │
│  - tools/archive_import.py (档案移交表入库与导出)                       │
│  - tools/material_collector.py (员工资料提取与 OCR 索引)                │
│  - tools/folder_rename.py (人员文件夹/文件批量重命名)                    │
├─────────────────────────────────────────────────────────────────────────┤
│                             核心数据处理层 (Core Data Engine)           │
│  - common/excel.py (openpyxl 样式快照、StyleArray 缓存、公式平移)       │
│  - common/excel_compat.py (三级 .xls 兼容引擎: xlrd -> COM -> LibreOffice)│
│  - common/inputs.py (ZIP 安全解压、编码自动识别 cp437/GBK、防 Zip-Bomb)  │
│  - rapidocr_onnxruntime (本地离线 OCR 推理)                            │
├─────────────────────────────────────────────────────────────────────────┤
│                             存储与基础设施层 (Storage & Infra)          │
│  - project_store.py (项目工作区管理、批次清单 Manifest、原子提交机制)    │
│  - history_store.py (全局历史索引、SQLite WAL、文件去重 SHA-256)        │
│  - runlog.py (业务脱敏运行日志、自动截断轮转)                           │
│  - app_update.py / update_runner.py (Gitee/GitHub 双源增量自更新)       │
└─────────────────────────────────────────────────────────────────────────┘
```

### 2.3 数据流转与操作全链路

```text
用户操作 (GUI 选择文件 / CLI 传参)
       │
       ▼
【项目工作区拦截 (project_store.py)】
  1. 校验输入路径安全性（拒绝非法链接/符号链接）
  2. 计算文件 SHA-256 校验和
  3. 将原始文件复制至沙盒目录（`.hrtoolkit/staging/`）
       │
       ▼
【异步任务派发 (gui.py)】
  1. 收集界面表单参数，构造强类型参数字典
  2. 创建 `threading.Thread(target=worker, daemon=True)`
  3. 启动 UI 进度条与状态监听
       │
       ▼
【核心业务执行 (tools/xxx.py)】
  1. 调用 `common.inputs` 安全解压 ZIP（如有）
  2. 调用 `common.excel_compat` 将 `.xls` 转换为 `.xlsx`
  3. 通过 `openpyxl` 读取数据并执行清洗、匹配、聚合
  4. 基于 `templates/*.xlsx` 内置模板克隆结构与公式
  5. 写入目标结果文件至当前批次 `处理结果/` 目录
       │
       ▼
【原子落盘与状态同步 (project_store.py)】
  1. 校验生成文件的完整性并记录 Manifest
  2. 将临时批次状态标记为 `success`
  3. 记录脱敏运行日志 (`runlog.py`)
       │
       ▼
【UI 响应刷新 (主线程)】
  1. 后台线程通过队列或回调通知主线程
  2. 刷新右侧“项目文件树”视图，展示新生成的 Excel
  3. 弹出操作成功摘要或异常 Warnings 提醒
```

---

## 3. 技术栈详细清单

| 层次 / 维度 | 选用技术 / 依赖 | 版本要求 | 选型理由与核心用途 |
| :--- | :--- | :--- | :--- |
| **基础运行环境** | Python | `>=3.9` | 兼顾现代 Python 特性（类型注解、zoneinfo）与旧版 Windows 兼容性。 |
| **GUI 表现层** | Tkinter / ttk | Python 标准库 | **无外部重型依赖**（对比 Qt/Electron 包体积骤降 100MB+），跨平台原生支持，启动毫秒级。 |
| **现代 Excel 读写** | `openpyxl` | `>=3.1,<4` | 支持 `.xlsx` 底层 XML 级控制，可完美读写单元格样式、公式、边框、合并单元格及冻结窗格。 |
| **旧版 Excel 兼容** | `xlrd` | `>=2.0,<3` | 纯 Python 内存只读解析老旧 `.xls` 格式，零外部依赖，毫秒级快速转换。 |
| **离线 OCR 识别** | `rapidocr_onnxruntime` | `>=1.3.0` | 纯本地 ONNX 推理，无须外网 API，高精度提取扫描件/图片中的文字及证件信息。 |
| **Windows 系统交互** | `pywin32` | `>=306` | （仅 Win 生效）用于通过 COM 调用 Microsoft Excel/WPS 导出复杂旧格式表格。 |
| **打包分发引擎** | `PyInstaller` | `>=6.0` | 将 Python 代码及二进制依赖打包为单目录（`--onedir`）及单文件更新器（`--onefile`）。 |
| **安装包制作** | Inno Setup / WiX | 命令行调用 | Windows 下构建企业级安装程序（`.exe` / `.msi`），安装至 `%LOCALAPPDATA%` 免管理员提权。 |
| **macOS 镜像** | `hdiutil` / `lipo` | macOS 内置 | 自动组装 `.dmg` 磁盘镜像，支持构建 Intel/Apple Silicon 双架构 `universal2`。 |
| **网络与证书** | `certifi` | `>=2024.8.30` | 打包环境下提供权威根证书，确保自更新 HTTPS 通信安全。 |
| **发布自动化** | Node.js / GitHub Actions | 18+ / CI | 本地 `npm run release` 执行版本校验与原子 Tag 推送，远端 Actions 自动化构建全平台产物。 |

---

## 4. 项目真实目录结构与模块职责

```text
hr-toolkit/
├── .github/workflows/          # [自动化流水线]
│   ├── ci.yml                  # 单元测试、编译检查、静态代码分析
│   └── release.yml             # 多平台构建 Windows/macOS 并发布至 GitHub & Gitee
├── packaging/                  # [打包静态资产]
│   ├── macos/                  # macOS Info.plist, entitlements, 图标
│   └── windows/                # Windows 应用程序清单 HRToolkit.manifest, 图标
├── scripts/                    # [构建与发布脚本集]
│   ├── build_windows.py        # Windows PyInstaller 打包与架构白名单校验
│   ├── build_windows_installers.py # 调用 Inno Setup / WiX 生成安装包
│   ├── build_update_assets.py  # 提取增量更新 ZIP 与 SHA256 校验清单
│   ├── build_macos.py          # macOS 构建 .app、通用二进制 universal2 及打包 .dmg
│   ├── release.py              # 本地发布控制台（严格 SemVer 校验、原子 Git Tag）
│   └── generate_app_icons.py   # 根据矢量算法自动渲染多尺寸应用图标
├── hr_toolkit/                 # [核心 Python 源码包]
│   ├── __init__.py             # 定义版本号 __version__ = "0.2.3"
│   ├── __main__.py             # 模块启动入口 (python -m hr_toolkit)
│   ├── cli.py                  # CLI 命令行入口与参数解析器
│   ├── gui.py                  # Tkinter UI 实现（三栏布局、组件交互、异步工作流）
│   ├── project_store.py        # ⭐️ 核心：项目工作区、沙盒存储与 Manifest 清单管理
│   ├── history_store.py        # ⭐️ 核心：历史索引数据库 (SQLite)、去重快照与回收站
│   ├── app_update.py           # ⭐️ 核心：Gitee/GitHub 双源更新检查与下载器
│   ├── update_runner.py        # ⭐️ 核心：独立更新器进程逻辑（文件安全替换与回滚）
│   ├── runlog.py               # 安全脱敏运行日志
│   ├── runtime_checks.py       # 无头自检指令 (--version, --smoke-test)
│   ├── _icon_data.py           # Base64 内嵌多分辨率窗口图标数据
│   ├── common/                 # [底层公共核心库]
│   │   ├── excel.py            # 高性能 openpyxl 样式快照 (RowSnapshot)、缓存与公式处理
│   │   ├── excel_compat.py     # .xls 兼容引擎（xlrd 内存流 -> COM -> LibreOffice）
│   │   ├── inputs.py           # 安全输入校验、ZIP 编码修复与防解压炸弹
│   │   ├── filenames.py        # 跨平台文件名脱敏与非法字符清理
│   │   └── resources.py        # 运行时内置模板路径定位器 (兼容开发/打包态)
│   ├── templates/              # [内置标准 Excel 模板库]
│   │   ├── social_security_summary_template.xlsx
│   │   ├── data_statistics_template.xlsx
│   │   ├── insurance_ledger_template.xlsx
│   │   └── archive_summary_template.xlsx
│   └── tools/                  # ⭐️ [业务逻辑插件目录 (按需求完全解耦)]
│       ├── social_security.py         # 需求1：社保账单解析、花名册匹配与分公司拆分
│       ├── data_statistics.py         # 需求2：打卡记录、周报/月报超时统计
│       ├── insurance_ledger.py        # 需求3：商保台账与加保/减保预警比对
│       ├── salary_split.py            # 需求4：工资表按入职公司与部门智能拆分
│       ├── salary_merge.py            # 需求5：跨月度工资合并与人员累计应发统计
│       ├── personnel_change_merge.py  # 需求6：各项目异动表按月归档与花名册更新
│       ├── archive_import.py          # 需求7：人事档案移交表按公司入库与生成
│       ├── folder_rename.py           # 需求8：员工档案文件夹/文件批量重命名
│       └── material_collector.py      # 需求9：员工证件资料批量归集与 OCR 自动命名
├── tests/                      # [30+ 自动化测试套件 (覆盖存储/业务/打包/UI)]
├── pyproject.toml              # 项目打包与元数据配置
├── requirements.txt            # Python 核心依赖清单
└── README.md                   # 详细使用说明与发布指引
```

---

## 5. 核心模块深度设计

### 5.1 业务模块标准范式 (`hr_toolkit/tools/`)
为保证架构可被其他团队轻松复用，业务逻辑层必须严格遵守**“输入纯粹、过程无 GUI、输出结构化”**三大原则：

```python
# 示例：tools/salary_split.py 核心设计结构
@dataclass
class CompanyOutput:
    company: str
    employee_count: int
    sections: list[str]
    file_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "company": self.company,
            "employee_count": self.employee_count,
            "sections": self.sections,
            "file_path": self.file_path,
        }

@dataclass
class SalarySplitResult:
    input_path: Path
    output_dir: Path
    dry_run: bool
    outputs: list[CompanyOutput] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "tool_name": "需求4-工资表按入职公司拆分",
            "input_path": str(self.input_path),
            "output_dir": str(self.output_dir),
            "dry_run": self.dry_run,
            "company_count": len(self.outputs),
            "employee_count": sum(item.employee_count for item in self.outputs),
            "outputs": [item.to_dict() for item in self.outputs],
            "warnings": self.warnings,
        }

def split_salary_by_company(
    input_path: str | Path,
    output_dir: str | Path,
    *,
    dry_run: bool = False,
) -> SalarySplitResult:
    """核心纯业务函数：
    1. 接收基础路径和参数（不依赖任何 UI 控件）
    2. 执行业务计算与 Excel 读写
    3. 返回强类型 Result 对象（内含结构化指标与 warnings）
    """
    ...
```

### 5.2 存储与工作区隔离机制 (`project_store.py`)
为了杜绝数据覆盖事故，本项目设计了**自包含的工作区文件模型**：
1.  **物理隔离**：
    *   用户在软件中选择一个本地目录建立“工作项目”。
    *   项目根目录下包含：`业务分类/工具名称/处理批次/`（内含 `上传资料/`、`处理结果/`、`补充资料/`）和 `共用资料/`。
    *   隐藏目录 `.hrtoolkit/` 存储：`project.json`（项目元数据）、`manifests/`（批次快照清单）、`staging/`（临时暂存区）、`trash/`（项目回收站）、`project-write.lock`（进程排他锁）。
2.  **原子批次提交**：
    *   外部资料在导入时，先拷贝到 `staging/` 并进行 SHA-256 校验。
    *   处理成功后，原子性地将文件移动到正式批次目录，并落盘批次 Manifest。如果中途崩溃或用户强退，未完成的暂存文件将在下次启动时安全回滚清理。

---

## 6. 数据处理与 Excel 性能优化方案

### 6.1 openpyxl 样式快照与索引缓存 (`common/excel.py`)
在处理数万行带有复杂样式的 Excel 表格时，如果简单使用 `cell.font = other_cell.font`，会导致 `openpyxl` 对每个单元格进行递归深拷贝（Deep Copy）和哈希计算，导致处理极其缓慢（甚至需要数分钟）。

**【当前项目的优化实现】**：
1.  **共享样式池直接引用**：直接提取工作簿共享样式表（`workbook._fonts`, `workbook._fills`, `workbook._borders`）的整数索引（`StyleArray`），避免对象构造。
2.  **`RowSnapshot` 行级快照**：封装单行的值、高度、单元格样式索引。
3.  **样式转换缓存（Style Translation Cache）**：跨工作簿复制时，缓存源工作簿与目标工作簿的样式下标映射表，成千上万行仅需转换一次样式。

### 6.2 旧版 `.xls` 文件的三级降级转换引擎 (`common/excel_compat.py`)
企业实际场景中充斥着各类 ERP/HR 系统导出的旧版 `.xls` 文件。本项目实现了全自动无感转换：
1.  **第一级（默认首选）**：基于 `xlrd` 纯 Python 流式解析，在内存中直接重建成 `.xlsx`。速度极快（0.01 秒级），跨平台通用，且 100% 避免改动源文件元数据。
2.  **第二级（Windows 兜底）**：若 `xlrd` 遇到特殊私有格式失败，Windows 环境通过 `win32com.client` 启动静默后台 Excel/WPS 进程转换为标准 `.xlsx`。
3.  **第三级（macOS/Linux 兜底）**：通过探测系统全局与 Homebrew 路径下的 `soffice` / `LibreOffice` 进行无头转换。

### 6.3 ZIP 安全解压与编码自适应 (`common/inputs.py`)
1.  **编码纠正**：Windows 下压缩软件经常使用 GBK 编码且不打 UTF-8 标记，导致 Python `zipfile` 解压出乱码。代码中通过探测 `flag_bits & 0x0800`，对非 UTF-8 编码自动执行 `cp437 -> gbk` 转码。
2.  **安全防护**：限制单文件上限（512MB）、总解压上限（2GB）、压缩比阈值（200倍），严防 Zip-Bomb；严格过滤 `..` 相对路径和符号链接（Symlink），防止路径穿越攻击。

---

## 7. 并发、异步与界面防假死方案

### 7.1 线程模型与调度
*   **【问题】**：Tkinter 属于单线程事件驱动模型（Main Loop）。如果直接在按钮点击事件中解析 10MB 的 Excel 或执行 OCR，界面会立刻“未响应”。
*   **【当前实现】**：
    *   **守护工作线程**：所有耗时计算一律包装在 `threading.Thread(target=worker, daemon=True).start()` 中执行。
    *   **线程安全通知**：工作线程内**严禁直接操作 Tkinter 控件**。工作线程通过线程安全队列 `queue.Queue` 发送消息，或者在主线程中使用 `root.after(100, poll_callback)` 周期性拉取状态，或通过 `root.after(0, update_ui_func)` 将更新调度回主线程。

```text
┌───────────────────────────┐         ┌───────────────────────────┐
│     GUI 主线程 (Tk)        │         │   Background Worker 线程   │
│                           │         │                           │
│  [用户点击执行]            │         │                           │
│       │                   │         │                           │
│       ├─ 禁用按钮/显示进度 ─┼─ 启动 ──>│  调用 tools/xxx.py        │
│       │                   │         │  读取/解析/计算 Excel      │
│       │                   │<─ 进度 ──┤  `queue.put(progress)`    │
│  [after 定时轮询队列]      │         │                           │
│  刷新 ProgressBar         │         │                           │
│                           │<─ 完成 ──┤  `queue.put(result)`      │
│  恢复按钮/展示结果         │         │                           │
└───────────────────────────┘         └───────────────────────────┘
```

---

## 8. 异常处理与安全审计日志

### 8.1 业务级容错 (Fault Tolerance)
*   **非中断式警告机制**：业务处理过程中遇到某一行缺失字段、身份证校验失败或格式微小异常时，不抛出致命异常中断流程，而是记录到 `result.warnings` 列表中。
*   **最终汇总呈现**：任务结束后在界面弹出清晰的中文提示框（如“处理完成：成功 120 人，发现 3 处异常需人工核对”），并列出具体原因。

### 8.2 业务数据脱敏日志 (`runlog.py`)
*   **脱敏铁律**：**严禁在日志中记录任何 Excel 表格内部的具体业务数据**（如员工姓名、身份证号、薪资数字、银行卡号等）。
*   **允许记录的内容**：
    *   程序启动与版本号。
    *   工具调用名称、输入文件名称及大小（如 `南昌工资表.xlsx(1.2MB)`）、所选参数。
    *   耗时、处理记录行数、预警条数。
    *   发生未捕获异常时的完整 Python Traceback。
*   **日志存储与轮转**：
    *   日志写入可执行文件同级的 `HRToolkit_app.log`。
    *   每次写入前检查文件大小，若超过 1MB 则自动截断保留最新 256KB，防止撑满磁盘。

---

## 9. 打包、分发与自更新体系 (重点)

### 9.1 PyInstaller 构建配置 (`scripts/build_windows.py`)
项目放弃了启动缓慢且不利于热更新的 `--onefile` 单文件模式，选用了**`--onedir` 目录模式**。

```python
# 核心打包命令参数配置
main_command = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--distpath", "dist/windows",
    "--name", "HRToolkit",
    "--onedir",                       # 单目录模式：启动极快，便于增量更新替换
    "--windowed",                     # 无黑色控制台窗口
    "--icon", "packaging/windows/HRToolkit.ico",
    "--manifest", "packaging/windows/HRToolkit.manifest", # Windows 现代 DPI 适配
    "--version-file", "build/HRToolkit.version.txt",      # 注入 PE 详细版本信息
    "--add-data", "README.md;.",                         # 打包说明文档
    "--add-data", "hr_toolkit/templates/*.xlsx;hr_toolkit/templates", # 打包内置模板
    "--hidden-import", "xlrd",                           # 强制包含动态导入模块
    "--hidden-import", "win32com.client",
    "--collect-all", "rapidocr_onnxruntime",             # 收集 OCR ONNX 模型与二进制
    "--exclude-module", "pytest",                        # 排除测试相关冗余包
    "--exclude-module", "unittest",
    "hr_toolkit_app.py"
]
```

### 9.2 产物安全与白名单校验 (`verify_windows_payload`)
打包完成后，构建脚本会对输出目录进行**严苛的安全扫描**：
*   **严禁项拦截**：遍历打包目录，若发现 `.db`、`.sqlite`、`.log`、测试文件、以及带有“上传资料/处理结果”字样的文件夹，构建立即报错中断，**物理杜绝开发者本地测试数据被误打包分发给用户**。
*   **架构合规检查**：读取所有 `.exe` / `.dll` / `.pyd` 的 PE Header 签名，确保全部为 `PE_MACHINE_AMD64` (0x8664)，防止混入 32 位二进制。

### 9.3 双更新源与全自动热更新机制 (`app_update.py` & `update_runner.py`)

1.  **双源更新发现（国内加速优先）**：
    *   第一更新源：Gitee API（`gitee.com/api/v5/.../releases/latest`，国内网络访问极速）。
    *   第二备用源：GitHub Releases（`github.com/.../releases/latest/download/latest.json`）。
    *   客户端启动时后台静默请求，优先读取 Gitee；若超时或被拦截，自动回退到 GitHub。
2.  **更新配置文件规范 (`latest.json`)**：
    ```json
    {
      "version": "0.2.3",
      "windows": {
        "version": "0.2.3",
        "file_url": "https://gitee.com/.../HRToolkit-0.2.3-win-update.zip",
        "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        "mandatory": true,
        "notes": ["1. 新增员工资料批量 OCR 提取工具", "2. 修复多月工资合并样式问题"]
      },
      "macos": {
        "version": "0.2.3",
        "file_url": "https://gitee.com/.../HRToolkit_0.2.3_universal.dmg",
        "sha256": "...",
        "update_mode": "manual"
      }
    }
    ```
3.  **独立更新器执行流程 (`HRToolkitUpdater.exe`)**：
    *   主程序下载更新 ZIP 并校验 SHA-256。
    *   主程序唤起轻量级独立更新器：`HRToolkitUpdater.exe --zip <path> --app-dir <dir> --wait-pid <pid> --relaunch --ui`。
    *   主程序退出释放文件句柄。
    *   更新器等待主进程彻底退出后，将原程序目录更名为 `_backup`，解压新版覆盖到目标目录。
    *   更新成功后清理备份和临时包，拉起新版主程序；若替换失败，自动恢复 `_backup` 原目录并记录更新日志。

---

## 10. 自动化测试与工程验证体系

本项目构建了四层质量防护网，所有变更必须通过完整验证方可发布：

```text
┌────────────────────────────────────────────────────────┐
│ 1. 单元与组件测试 (tests/test_*.py)                     │
│    - 300+ 项测试用例，覆盖各工具算法、Excel 快照、样式平移│
├────────────────────────────────────────────────────────┤
│ 2. 编译与语法严格检查                                  │
│    - python -m compileall (全模块无语法错误)            │
│    - git diff --check (无残余冲突标记与非法空白字符)     │
├────────────────────────────────────────────────────────┤
│ 3. 无头运行时烟测 (Runtime Smoke Tests)                │
│    - 打包后可执行文件执行 HRToolkit.exe --version       │
│    - 执行 HRToolkit.exe --smoke-test (验证内置模板与沙盒)│
│    - 执行 HRToolkit.exe --update-smoke-test (验证更新源) │
├────────────────────────────────────────────────────────┤
│ 4. 平台发布前演练 (Dry-Run Release)                    │
│    - npm run release -- 0.2.4 --dry-run               │
│    - 校验 SemVer、分支纯净度、Tag 冲突与版本一致性     │
└────────────────────────────────────────────────────────┘
```

---

## 11. 新部门项目复用落地方案 (核心实施指南)

如果要基于本套架构，为**财务部（Finance Toolkit）**、**运营部（Ops Toolkit）**或**供应链部**开发一套全新的自动化工具，研发团队应当按照以下标准化步骤实施：

### 11.1 资产复用与重构对照表

| 模块类别 | 涉及文件路径 | 复用策略 | 实施说明 |
| :--- | :--- | :--- | :--- |
| **基础设施** | `hr_toolkit/project_store.py`<br>`hr_toolkit/history_store.py`<br>`hr_toolkit/runlog.py` | **100% 原样保留** | 仅需全局搜索将类名或常量标识中的 `HRToolkit` 替换为新系统名（如 `FinanceToolkit`）。 |
| **通用处理层** | `hr_toolkit/common/*` | **100% 原样保留** | 保留 Excel 样式快照引擎、`.xls` 兼容转换、ZIP 安全解压、文件名清理。 |
| **打包与 CI/CD**| `scripts/*`<br>`.github/workflows/*` | **95% 保留** | 修改 `APP_NAME = "FinanceToolkit"`，并在 GitHub Actions 中配置对应的 Gitee 仓库 Token。 |
| **自动更新体系**| `hr_toolkit/app_update.py`<br>`hr_toolkit/update_runner.py` | **95% 保留** | 修改其中的 `GITEE_REPOSITORY` 和 `GITHUB_REPOSITORY` 地址。 |
| **业务逻辑层** | `hr_toolkit/tools/*` | **100% 替换** | 清空原有 HR 业务，新建新部门的业务脚本（如 `invoice_audit.py`, `tax_summary.py`）。 |
| **模板文件** | `hr_toolkit/templates/*` | **100% 替换** | 放入新部门的标准输出 Excel 空模板。 |
| **UI 页面与表单**| `hr_toolkit/gui.py` | **针对性适配** | 保留整体三栏框架、窗口拖拽、文件树，修改左侧功能导航菜单和中间表单参数项。 |
| **CLI 入口** | `hr_toolkit/cli.py` | **针对性适配** | 重新定义 `argparse` 子命令，与新的 `tools` 模块一一绑定。 |

---

### 11.2 推荐的新项目脚手架模板 (以财务工具箱为例)

```text
finance-toolkit/
├── .github/workflows/
│   ├── ci.yml                          # CI 流水线
│   └── release.yml                     # 跨平台自动打包与发布流水线
├── packaging/
│   ├── macos/FinanceToolkit.icns
│   └── windows/FinanceToolkit.ico
├── scripts/
│   ├── build_windows.py                # 全局常量 APP_NAME = "FinanceToolkit"
│   ├── build_windows_installers.py
│   ├── build_update_assets.py
│   ├── build_macos.py
│   └── release.py
├── src/finance_toolkit/
│   ├── __init__.py                     # __version__ = "1.0.0"
│   ├── __main__.py
│   ├── cli.py                          # 财务命令行入口
│   ├── gui.py                          # 财务桌面 UI (三栏式)
│   ├── project_store.py                # 项目工作区沙盒 (直接沿用)
│   ├── history_store.py                # 历史索引库 (直接沿用)
│   ├── app_update.py                   # 双源自更新 (指向 finance-toolkit 仓库)
│   ├── update_runner.py                # 独立更新器 (直接沿用)
│   ├── runlog.py                       # 运行日志 (直接沿用)
│   ├── runtime_checks.py               # 烟测自检 (直接沿用)
│   ├── common/                         # 底层通用能力 (直接沿用)
│   │   ├── excel.py                    # openpyxl 样式快照引擎
│   │   ├── excel_compat.py             # .xls 三级兼容转换
│   │   ├── inputs.py                   # ZIP 安全解压与编码修复
│   │   ├── filenames.py
│   │   └── resources.py
│   ├── templates/                      # 财务专属模板
│   │   ├── invoice_summary_template.xlsx
│   │   └── tax_report_template.xlsx
│   └── tools/                          # ⭐️ 财务业务开发区
│       ├── invoice_validation.py       # 业务1：发票真伪批量核验与汇总
│       ├── tax_splitter.py             # 业务2：个税明细按项目拆分
│       └── bank_reconciliation.py      # 业务3：银行对账单自动平账
├── tests/                              # 针对财务模块编写单元测试
├── pyproject.toml                      # name = "finance-toolkit"
├── requirements.txt                    # 依赖清单
└── README.md
```

---

### 11.3 研发落地五步走实施法

1.  **第一步：脚手架克隆与基础重命名（第 1 天）**
    *   拉取本工程代码，重命名包名（如 `hr_toolkit` -> `finance_toolkit`）。
    *   在 `scripts/build_windows.py`、`gui.py`、`app_update.py` 中更新软件名称、图标和更新源仓库地址。
    *   执行 `pytest` 与 `python -m compileall`，确保基线脚手架测试全部通过。
2.  **第二步：明确业务规则与输出模板（第 2 天）**
    *   收集业务方真实的数据源样例（含 `.xlsx`, `.xls`, `.zip`）和标准输出 Excel 模板。
    *   将空模板放入 `templates/` 目录，并在 `build_windows.py` 的 `RELEASE_TEMPLATE_NAMES` 白名单中登记。
3.  **第三步：独立开发业务模块 `tools/`（第 3~4 天）**
    *   在 `tools/` 下编写独立的纯 Python 业务处理模块。
    *   使用 `common.excel.snapshot_row` 和 `apply_row_snapshot` 保证生成报表的格式、边框、公式与模板 100% 吻合。
    *   在 `cli.py` 中增加对应子命令，编写单元测试并确保通过命令行测试成功。
4.  **第四步：GUI 界面装配与异步接入（第 5 天）**
    *   在 `gui.py` 左侧导航增加功能入口卡片。
    *   中间主操作区绑定文件选择器与业务参数项。
    *   在按钮事件中通过 `threading.Thread` 调用新编写的 `tools` 模块，并通过队列向界面回传处理进度和预警日志。
5.  **第五步：本地验证、CI 打包与全自动交付（第 6 天）**
    *   本地执行 `python -m finance_toolkit` 进行界面完整链路验证。
    *   执行 `npm run release -- 1.0.0 --dry-run` 演练发布流程。
    *   推送到 GitHub 触发 Actions，自动生成 Windows `.exe`/`.msi` 安装包及 macOS `.dmg` 镜像，并自动镜像至 Gitee Release 供用户下载体验。

---

## 12. 总结与架构评审意见

本方案所提炼的架构体系已经在生产环境中经历了严苛的数据安全性检验、旧格式兼容性检验和跨平台自动化构建考验：
1.  **高安全性**：通过“本地项目工作区 + 沙盒隔离 + 严格脱敏日志”，杜绝了核心数据外泄与原文件误破坏的风险。
2.  **高工程度**：具备从开发、测试、无头自检、资源过滤、跨平台打包到国内双源热更新的完整工业级生命周期闭环。
3.  **高复用性**：表现层、调度层、存储层与通用数据层高度沉淀，新部门业务逻辑完全以插件形式在 `tools/` 扩展，使得开发一套全新的企业级工具箱仅需 3~5 个工作日，具备极高的复制与推广价值。
