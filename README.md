# HR Toolkit

人事 Excel 自动化工具箱。当前已落地：**需求1：社保明细与汇总**、**需求2：考勤与周月报统计**、**需求3：保险台账与增减预警**、**需求4：工资表按入职公司拆分**、**需求5：多月工资合并个人薪资汇总**、**需求6：异动表汇总与花名册更新**、**需求7：档案移交表入库**、**需求8：人员资料文件夹改名**。

## 已实现工具

### 需求1-社保明细与汇总

输入社保账户缴费清单、zip 压缩包，或一个包含多个社保清单/压缩包的文件夹，再选择参保人员花名册，按身份证关联人员信息，输出社保明细表和社保汇总表。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 自动识别老 `.xls` 长表、单险种明细表、宽表缴费清单
- 按身份证匹配参保人员花名册
- 自动识别账单期，优先使用账单文件夹或文件名月份；账单内跨月人员也计入本次上传账单月份
- 社保缴纳地、缴纳单位优先按账单文件夹或文件名识别；与花名册不一致时会提醒
- 生成总的 `社保明细表.xlsx`
- 按参保单位/参保地额外拆分明细表，例如 `唐人四川-社保明细表.xlsx`、`唐人长春-社保明细表.xlsx`
- 生成 `社保汇总表.xlsx`，包含总汇总、按公司汇总、按缴纳单位/参保地/险种/项目的数据分析和异常提醒
- 未匹配花名册的人员、姓名不一致、未识别账单期会在日志和异常提醒中列出

公积金、残保金、管理费暂无数据时留空；后续人事提供单独数据后可继续补充。

### 需求2-考勤与周月报统计

输入 HR 系统导出的考勤结果、周报记录、月报记录，支持单个文件、多个文件、zip 压缩包，或包含这些文件/压缩包的文件夹，自动生成考勤和周月报统计表。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 自动识别 `考勤结果`、`周报记录`、`月报记录`
- 生成 `考勤统计`，汇总事假、病假、带薪休假、调休、各月加班、旷工、迟到/早退、漏打卡和备注
- 生成 `考勤异常明细`，列出漏打卡、加班、调休、迟到/早退、旷工等明细
- 可选 `应汇报人员名单`，用于准确统计未写周报、未写月报人员
- 生成 `周月报统计`，统计未写周报、周报超时、未写月报、月报超时，不计算扣款金额
- 生成 `周月报异常明细`，列出异常人员、周期、截止时间、实际汇报时间和来源文件
- 周报截止为次周一 `17:00:59` 前正常、`17:01` 起超时；月报按次月 2 日同样规则判断
- 可选周报统计日期范围（如 `2026-06-02` 至 `2026-06-30`），只统计范围内周一截止的周报；适合 1 号正好是周一的月份，避免把上月最后一周重复统计。界面提供“本月/上月/本周/上周”快捷填充，填充后仍可手动修改
- 输出 `考勤周月报汇总表.xlsx`

未提供应汇报人员名单时，未写周/月报只能按文件中可推断人员统计。

**周报算哪一期，按提交时间判断**（以截止时间周一 6.15 17:00 为例）：

| 提交时间 | 算哪一期 | 统计表显示 |
| --- | --- | --- |
| 上周五 0:00 ～ 周一 17:00:59 | 6.15 这期 | 正常 |
| 周一 17:01 ～ 周一 23:59 | 6.15 这期 | 超时（如“17:30提交”） |
| 周二 ～ 周四 | 6.15 这期的补交 | 超时（如“6月17日9:05提交”） |
| 周五 0:00 起 | 下一期（6.22 截止） | 下一期正常 |

两个容易疑惑的情况：

- **上一期已经交过，周二到周四又交了一份**：这份视为提前交的下一期周报，不记超时，下一期也不会记未写。例如小王 6.15 周一按时交了，6.18 周四因为要请假提前把本周的交了，那 6.22 那期就算他已交。
- **归属期超出所选日期范围**：选了统计日期时，归属期不在范围内的周报本次不统计，留给下一次，避免重复统计或错记超时。例如范围选到 6.24，某人 6.26（周五）交的周报属于 6.29 截止那期，本次统计里不会出现，下次统计到 6.29 那期时才会算。

### 需求3-保险台账与增减预警

输入各保单人员清单、zip 压缩包，或一个包含多个保单清单/压缩包的文件夹，再选择需求6的人力资源分析表，自动生成保险台账。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 自动识别保单号和保单人员
- PZDX 类保单的保额取 `每人伤残死亡限额`，按万元显示
- PEAC 类保单没有明显保额字段，固定按 `60` 万元显示
- `项目/部门` 从人力资源分析表的 `花名册` 工作表补充，优先取 `部门/项目` 列
- 生成人员增减预警：花名册在职但保单没有提示 `需加保`，保单有但花名册没有或已标记离职提示 `需减保`
- 输出 `保险台账.xlsx`，包含 `保险台账` 和 `人员增减预警` 两个工作表
- 如存在 `需加保` 人员，会额外输出 `人力资源分析表_保险预警.xlsx`，在 `花名册` 中标记 `保险预警`

人事已确认岗位保险规则暂取消，当前只做台账明细和人员增减预警。

### 需求4-工资表按入职公司拆分

输入一个包含 `汇总表`、`明细表` 的工资表，按明细表中的 `入职公司` 字段拆分为多个 Excel 工作簿。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 每个入职公司一个独立 `.xlsx`
- 保留原工资表的主要样式、表头、公式结构
- 明细表只保留对应公司的员工行
- 明细表保留原模板分段小计和底部总计文案
- 汇总表引用拆分后明细表中的分段小计

### 需求5-多月工资合并个人薪资汇总

输入单个月度工资表、多个工资表、zip 压缩包，或一个包含多个月度工资表/压缩包的文件夹，按 `身份证号码` 合并，输出每个人一行的个人应发工资汇总表。也可以同时选择已有汇总表，工具会把新月份追加进去。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 自动识别每张工资表的月份
- 按身份证号码合并同一员工
- 已有汇总表中已经存在的人员月份不覆盖，避免重复写入
- 新月份中出现的新员工会自动新增一行
- 人员在某个月没有工资时自动填 `0`
- 输出 `个人薪资汇总表.xlsx`

### 需求6-异动表汇总

输入单个项目异动表、多个项目异动表、zip 压缩包，或一个包含多个项目异动表/压缩包的文件夹，将各项目填写的 `增补表`、`离职`、`转正`、`调整` 按记录日期分到对应月份汇总表。文件夹里如果同时放入人力资源分析表，工具会同步更新其中的 `花名册`。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 支持项目表中的 `增补表`、`离职`、`转正`、`调整`
- zip 会自动解压后读取，文件夹内的 zip 也会自动处理
- 不选择已有汇总表时，会按月份新建干净汇总表
- 选择已有汇总表文件或汇总表文件夹时，会按月份追加新记录，不会清空原记录
- 对应月份没有已有汇总表时，会自动创建该月份汇总表
- 已存在的异动记录不会重复追加，只会补充已有行中的空白字段
- 月份规则：增员看 `入职日期/入职时间`，减员看 `离职日期`，转正看 `转正日期`，调动看 `调整日期`
- 忽略模板中只有预填序号、没有填写内容的空行
- 汇总后重新编排各 sheet 序号
- 保留模板工作簿样式
- 按记录月份输出，例如 `2026年4月异动汇总表.xlsx`
- `增员` 会插入花名册对应项目后方，`减员` 会在花名册中标黄
- GUI 中可切换到 `花名册更新` 页，单独选择异动汇总表和人力资源花名册进行更新

薪酬、产值和同行对比分析暂不处理，等需求6第三部分数据源确认后再单独实现。

### 需求7-档案移交表入库

输入项目部提交的人事档案移交表，按 `公司` 写入档案汇总表；也可以从一份或多份档案汇总表生成各公司独立档案表。

输出内容：

- 支持单个 `.xlsx/.xls` 移交表、多个移交表、`.zip` 压缩包，或包含移交表/压缩包的文件夹
- 按 `公司` 自动写入档案汇总表对应工作表
- 已有档案汇总表可选；不选择时使用内置空模板新建汇总表
- 身份证已存在时不重复新增，只补充原汇总表中为空的材料字段
- 档案汇总表缺少公司工作表时，会按第一个工作表样式自动创建
- `编号` 从文件名、表头标题或公司名识别项目地区，例如 `茂名项目部` 自动填 `11`
- `档案号` 使用模板公式按 `编号-入职公式-出生年月公式-序号` 生成
- 档案表中有、汇总表没有的字段会汇总到 `其他`
- 档案入库输出 `档案表汇总表.xlsx`
- 档案表生成支持选择已有公司档案表；匹配到公司就追加，没匹配到就用内置干净模板新建
- 生成公司档案表时会自动改标题公司名，并为新增人员补边框、居中和公式
- 档案表生成会按公司输出 `公司名-档案表.xlsx`

### Excel 旧格式兼容

- 已实现的 Excel 类工具均支持上传 `.xlsx` 和 `.xls`
- 文件夹和 zip 压缩包中也会识别 `.xls`
- 输出文件统一为 `.xlsx`
- 需求1的老 `.xls` 社保清单和参保花名册会用内置依赖直接读取
- 其他工具遇到 `.xls` 会先自动转换为 `.xlsx` 再处理；Windows 电脑需要安装 Excel 或 WPS 表格，Mac/Linux 需要安装 LibreOffice 才能自动转换

### 需求8-人员资料文件夹改名

选择一个人员资料目录，对目录下的人员文件夹或指定类型的文件（PDF、图片、文档）做批量改名。执行前会先预览，并二次确认。

支持内容：

- 选择项目类型：文件夹、PDF、图片、文档，或全部
- 批量追加后缀，例如 `张三` -> `张三-劳动合同`（追加文字会按原样写入，建议以 `-` 或 `_` 开头）
- 批量删除结尾文字，例如 `张三_劳动合同`、`李四劳动合同` -> `张三`、`李四`
- 指定单个人员/单文件处理，例如只处理 `张三` 或 `张三.pdf`
- 替换单个名称，例如 `张三` -> `章五`；替换文件时会自动补全原扩展名

## 桌面版使用

无参数启动时会打开图形界面：

```bash
python3 -m hr_toolkit
```

界面操作流程：

1. 在左侧选择工具
2. 选择工资表文件、工资表文件夹、异动表文件夹、档案移交表文件夹或人员资料目录
3. 保存位置默认在桌面结果目录下，也可以手动选择
4. 点击 `开始拆分`、`开始合并`、`开始汇总`、`开始入库` 或 `预览`
5. 程序会在保存位置下自动创建 `结果_年月日_时分秒` 子文件夹
6. 点击 `打开所在文件夹` 查看结果

## Mac 本机验证

当前目录执行：

```bash
python3 -m pip install -r requirements.txt
python3 -m hr_toolkit social-security \
  --input "问题1-3相关数据及模板/1.社保类模板" \
  --roster "问题1-3相关数据及模板/1.社保类模板/参保人员花名册.xlsx" \
  --output "outputs/social_security_demo"
```

考勤与周月报统计：

```bash
python3 -m hr_toolkit data-statistics \
  --input "问题1-3相关数据及模板/2.数据统计类模板" \
  --output "outputs/data_statistics_demo"
```

如有人事提供的应汇报人员名单，可追加：

```bash
python3 -m hr_toolkit data-statistics \
  --input "问题1-3相关数据及模板/2.数据统计类模板" \
  --staff "应汇报人员名单.xlsx" \
  --output "outputs/data_statistics_demo"
```

只统计指定日期范围内周一截止的周报（两个日期需同时提供）：

```bash
python3 -m hr_toolkit data-statistics \
  --input "问题1-3相关数据及模板/2.数据统计类模板" \
  --week-start 2026-06-02 \
  --week-end 2026-06-30 \
  --output "outputs/data_statistics_demo"
```

保险台账：

```bash
python3 -m hr_toolkit insurance-ledger \
  --input "问题1-3相关数据及模板/3.保险类模板" \
  --roster "问题6-2026年4月人力资源分析.xlsx" \
  --output "outputs/insurance_ledger_demo"
```

工资表拆分：

```bash
python3 -m hr_toolkit salary-split \
  --input "附件/问题4-薪资表模板(1).xlsx" \
  --output "outputs/salary_split_demo"
```

预览模式，不生成文件：

```bash
python3 -m hr_toolkit salary-split \
  --input "附件/问题4-薪资表模板(1).xlsx" \
  --output "outputs/salary_split_demo" \
  --dry-run
```

系统集成时建议使用 JSON 输出：

```bash
python3 -m hr_toolkit salary-split \
  --input "附件/问题4-薪资表模板(1).xlsx" \
  --output "outputs/salary_split_demo" \
  --json
```

多月工资合并：

```bash
python3 -m hr_toolkit salary-merge \
  --input-dir "某项目工资表文件夹" \
  --output "outputs/salary_merge_demo"
```

已有汇总表追加新月份：

```bash
python3 -m hr_toolkit salary-merge \
  --input-dir "第三月工资表文件夹" \
  --summary "已有个人薪资汇总表.xlsx" \
  --output "outputs/salary_merge_demo"
```

异动表汇总：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --output "outputs/change_merge_demo"
```

单个异动表也可以直接处理：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "问题6-2026年4月南昌分公司异动表.xlsx" \
  --output "outputs/change_merge_demo"
```

追加到已有异动汇总表：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "已有异动汇总表.xlsx" \
  --output "outputs/change_merge_demo"
```

追加到一个包含多个月份汇总表的文件夹：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "已有月度汇总表文件夹" \
  --output "outputs/change_merge_demo"
```

zip 压缩包也可以直接处理：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "项目部异动表.zip" \
  --output "outputs/change_merge_demo"
```

指定人力资源分析表并同步更新花名册：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "已有异动汇总表.xlsx" \
  --analysis-template "问题6-2026年4月人力资源分析.xlsx" \
  --output "outputs/change_merge_demo"
```

只用已有异动汇总表单独更新花名册：

```bash
python3 -m hr_toolkit roster-update \
  --input "已有月度汇总表文件夹" \
  --roster "人力资源花名册.xlsx" \
  --output "outputs/roster_update_demo"
```

档案移交表入库：

```bash
python3 -m hr_toolkit archive-import \
  --input "档案移交表文件夹" \
  --output "outputs/archive_import_demo"
```

追加到已有档案汇总表：

```bash
python3 -m hr_toolkit archive-import \
  --input "档案移交表文件夹" \
  --target "档案表汇总表.xlsx" \
  --output "outputs/archive_import_demo"
```

档案入库预览，不生成文件：

```bash
python3 -m hr_toolkit archive-import \
  --input "档案移交表文件夹" \
  --output "outputs/archive_import_demo" \
  --dry-run
```

按公司生成独立档案表：

```bash
python3 -m hr_toolkit archive-export \
  --summary "档案表汇总表.xlsx" \
  --output "outputs/archive_export_demo"
```

追加到已有公司档案表：

```bash
python3 -m hr_toolkit archive-export \
  --summary "档案表汇总表.xlsx" \
  --existing "已有公司档案表文件夹" \
  --output "outputs/archive_export_demo"
```

人员资料文件夹改名预览：

```bash
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode append \
  --text "劳动合同"
```

确认执行改名：

```bash
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode append \
  --text "劳动合同" \
  --apply
```

删除结尾文字：

```bash
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode remove \
  --text=_劳动合同 \
  --apply
```

替换单个文件夹名：

```bash
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode replace \
  --target "张三" \
  --replacement "章五" \
  --apply
```

## 后续扩展约定

每个需求独立成一个工具模块：

- `hr_toolkit/tools/social_security.py`：需求1，社保明细/汇总
- `hr_toolkit/tools/data_statistics.py`：需求2，考勤与周月报统计
- `hr_toolkit/tools/insurance_ledger.py`：需求3，保险台账
- `hr_toolkit/tools/salary_split.py`：需求4，工资表拆分
- `hr_toolkit/tools/salary_merge.py`：需求5，工资表合并
- `hr_toolkit/tools/personnel_change_merge.py`：需求6，异动表汇总
- `hr_toolkit/tools/archive_import.py`：需求7，档案移交表入库
- `hr_toolkit/tools/folder_rename.py`：需求8，人员资料文件夹改名

CLI 只是入口，核心函数可以直接被 ScriptHub 或 Web 后端调用。

## 自动构建与发布

日常发布只在本地 Mac 做版本检查、测试、版本提交、annotated Tag 和原子推送；Windows、macOS 构建与 GitHub Release 发布全部交给 GitHub Actions。正式发布命令为：

```bash
npm run release -- 0.2.1
```

首次使用先安装 Python 依赖和 Node.js。npm 入口会优先使用 `.venv/bin/python`，不存在时再使用 `python3`；两者之一必须能运行完整测试。发布前可执行不修改版本文件、commit、Tag 或远端的演练：

```bash
npm run release -- 0.2.1 --dry-run
```

无人值守环境审核完版本后可追加 `--yes`。发布脚本会严格检查 stable SemVer、clean `main`、`HEAD == origin/main`、本地/远端 Tag 冲突以及全部版本字段，然后运行 `unittest`、`compileall` 和 `git diff --check`。正式执行只会暂存 `hr_toolkit/__init__.py`、`package.json`、`package-lock.json`，不会运行本地跨平台构建，也不会使用 `git add .`。

脚本创建单一版本提交和 `v<version>` annotated Tag，再通过一次 atomic push 同时推送 `main` 与 Tag。推送失败时只有在确认远端两个引用都未变化后才自动回滚；远端状态不明确时会保留现场，要求人工核对。

> 不要为了测试发布脚本在正式仓库创建 `v0.2.1`。先使用 `--dry-run`；正式命令必须等发布负责人确认。

### GitHub Actions 产物

普通 push 和 pull request 由 `.github/workflows/ci.yml` 运行 Python 3.9+ 测试、编译和静态发布检查。只有 `v*` Tag 会触发 `.github/workflows/release.yml`：先校验 Tag 与 `hr_toolkit.__version__` 完全一致，再分别构建 Windows 与 macOS；两个平台全部成功后才创建并发布 GitHub Release。

`v0.2.1` 的直接下载资产为：

```text
HRToolkit_0.2.1_universal.dmg
HRToolkit_0.2.1_x64-setup.exe
HRToolkit_0.2.1_x64.msi
HRToolkit-0.2.1-win-update.zip
latest.json
SHA256SUMS.txt
```

macOS 优先构建 universal2，并对 Bundle 中的 Mach-O 使用 `file`/`lipo` 验证 `arm64` 与 `x86_64`。如果 universal2 构建或验证失败，发布资产会改为两个真实架构文件：

```text
HRToolkit_0.2.1_arm64.dmg
HRToolkit_0.2.1_x64.dmg
```

不会把单架构程序改名伪装成 `universal`。DMG 内包含标准 `HRToolkit.app` 和指向 `/Applications` 的快捷方式。

正式 PyInstaller 数据资源只允许 `README.md` 与 `hr_toolkit/templates/*.xlsx`。构建验证会拒绝附件、真实 Excel、outputs、日志、缓存和测试数据。打包后的主程序支持无界面检查：

```bash
HRToolkit --version
HRToolkit --smoke-test
```

### Windows 三阶段构建

Windows runner 按三个独立阶段执行：

```powershell
python scripts\build_windows.py --version 0.2.1 --output-dir dist\windows
python scripts\build_windows_installers.py --version 0.2.1 --app-dir dist\windows\HRToolkit --updater dist\windows\HRToolkitUpdater.exe --output-dir artifacts\windows
python scripts\build_update_assets.py --version 0.2.1 --app-dir dist\windows\HRToolkit --updater dist\windows\HRToolkitUpdater.exe --output-dir artifacts\windows
```

兼容入口 `scripts/release_windows.py` 只负责依次调用这三个阶段，不再递增版本、不提交代码、不上传发布物。主程序是 PyInstaller onedir，Updater 是 onefile；更新 ZIP 根目录保持 `HRToolkit.exe`、`HRToolkitUpdater.exe` 和 `_internal/`，继续使用现有备份、替换与失败回滚逻辑。

EXE 安装器是普通用户的主要下载项，MSI 用于企业部署。两者使用当前用户可写的 `%LOCALAPPDATA%\Programs\HRToolkit`，程序 payload 位于其 `app` 子目录，安装器/卸载器元数据留在外层，保证自更新器可以只替换 `app`。

### macOS 本地检查

当前 Mac 可以在不发布的情况下构建并检查当前版本 DMG：

```bash
python scripts/build_macos.py \
  --version 0.1.32 \
  --architecture arm64 \
  --output-dir dist/release-assets
```

使用 universal2 Python 时把架构改为 `universal2`。构建脚本会生成 `.app`、创建 DMG、验证 Applications 快捷方式、Bundle 版本、资源白名单、无界面启动以及所有 Mach-O 的真实架构。

### 更新地址迁移与 v0.2.1 桥接

新版本默认读取公开 GitHub Release：

```text
https://github.com/xhzwjc/hr-toolkit/releases/latest/download/latest.json
```

Windows 继续下载 `HRToolkit-<version>-win-update.zip` 并使用现有 Updater 自动替换。`v0.2.1` 发布成功后，Windows 构建还会提供一次性 `legacy-server-latest.json`：先确认其中的 GitHub 下载地址和 SHA256，再把它部署到旧服务器原 `latest.json` 的位置。旧 Windows 客户端由服务器发现 `v0.2.1`，下载 GitHub Release 的更新 ZIP；安装后的 `update_url.txt` 已转向 GitHub，以后不再依赖旧服务器。不要在 GitHub Release 成功前部署桥接文件。

macOS 第一阶段只支持 DMG 手动更新。新版客户端发现更新后只打开 DMG 下载地址，不会调用 ZIP 替换器，也不应宣传为 Mac 自动更新。旧 Mac 客户端不能安全消费标准 `.app` DMG，应通过人工通知和 DMG 安装迁移，旧服务器桥接清单只提供 Windows 条目。

### 签名预留

当前版本不做 Windows Authenticode、Apple Developer ID 签名或 notarization，因此正式构建不要求额外签名 Secret。Windows 安装器保留 `--inno-sign-tool-name`，macOS 构建保留 `--codesign-identity` / `MACOS_CODESIGN_IDENTITY` 与 entitlements 入口；未来启用时再配置证书、密码、Developer ID、Apple ID/App Store Connect 凭据，并在 Release 发布门禁前加入完整签名与公证验证。

应用图标（窗口/任务栏/exe）由 `scripts/generate_app_icons.py` 生成：品牌绿圆角方块加白色 “HR” 字标，与侧栏标识一致。调整图标后运行该脚本，重新生成 `hr_toolkit/_icon_data.py` 与 `packaging/windows/HRToolkit.ico`；macOS 构建会基于同一图形生成 `.icns`。

## 自动更新行为

启动时的检查在后台静默进行，只有发现新版本才会弹窗提示；主界面右上角也有“检查更新”，可手动检查。更新提示遵循 `latest.json` 中的 `mandatory` 字段：强制更新（默认）只能“立即更新”或退出程序；非强制更新（`"mandatory": false`）可选择“稍后再说”，下次启动时再次提醒。

Windows 下载完成后会启动 `HRToolkitUpdater.exe`，关闭主程序，替换整个 payload 目录后自动重新打开新版本。更新成功会清理下载包和临时文件；失败时会保留或恢复备份，并在安装目录同级写入 `HRToolkit_update.log`。

如果旧版本更新失败后只剩 `HRToolkit_backup_时间`，先把该文件夹改名回 `HRToolkit`，再打开 `HRToolkit.exe` 检查更新。正常版本不需要单独下载更新器，更新器已包含在更新 ZIP 中。

## 运行日志

程序在 `HRToolkit_update.log` 同一位置写入 `HRToolkit_app.log`（开发环境写到当前目录，可用环境变量 `HR_TOOLKIT_APP_LOG` 指定路径）。远程排查问题时，让使用者把这两个日志文件发过来即可。

记录内容：程序启动（版本号）、每次工具运行的开始（工具名、输入文件名和大小、参数）、完成（耗时、提醒条数）、失败（耗时和完整错误堆栈）、用户手动停止，以及界面和后台线程的未捕获异常。**只记录文件名和统计数字，不记录任何表格内容**（身份证、工资等敏感数据不会进日志），日志文件可以放心外发。日志超过 1MB 时自动截断，只保留最近的内容。
