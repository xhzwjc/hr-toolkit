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
- 周报按自然周的下周一 `17:00:59` 前正常、`17:01` 起异常；月报按次月 2 日同样规则判断
- 输出 `考勤周月报汇总表.xlsx`

未提供应汇报人员名单时，未写周/月报只能按文件中可推断人员统计。

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

选择一个人员资料目录，对目录下的人员文件夹做批量改名。执行前会先预览，并二次确认。

支持内容：

- 批量追加后缀，例如 `张三` -> `张三-劳动合同`
- 批量删除结尾文字，例如 `张三_劳动合同`、`李四劳动合同` -> `张三`、`李四`
- 指定单个人员处理，例如只处理 `张三`
- 替换单个文件夹名，例如 `张三` -> `章五`

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

## Windows 打包

命令行版调试包：

```powershell
python -m PyInstaller --name HRToolkit --onedir --console --clean --add-data "README.md;." --add-data "hr_toolkit/templates;hr_toolkit/templates" --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --hidden-import win32timezone --hidden-import xlrd hr_toolkit_app.py
```

给人事双击使用的桌面版：

```powershell
python -m PyInstaller --name HRToolkit --onedir --windowed --clean --add-data "README.md;." --add-data "hr_toolkit/templates;hr_toolkit/templates" --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32com.client --hidden-import win32timezone --hidden-import xlrd hr_toolkit_app.py
```

自动更新程序：

```powershell
python -m PyInstaller --name HRToolkitUpdater --onefile --windowed --clean hr_toolkit_updater.py
Copy-Item dist\HRToolkitUpdater.exe dist\HRToolkit\ -Force
```

Mac 打包时把 `;` 改成 `:`，更新程序复制无后缀文件：

```bash
python -m PyInstaller --name HRToolkit --onedir --windowed --clean --add-data "README.md:." --add-data "hr_toolkit/templates:hr_toolkit/templates" --hidden-import xlrd hr_toolkit_app.py
python -m PyInstaller --name HRToolkitUpdater --onefile --windowed --clean hr_toolkit_updater.py
cp dist/HRToolkitUpdater dist/HRToolkit/
```

打包完成后，把整个目录发给使用者：

```text
dist/
  HRToolkit/
    HRToolkit.exe
    HRToolkitUpdater.exe
    update_url.txt
    _internal/
```

不要只发送 `.exe`，`_internal` 目录也是程序运行必需的。

## 自动更新发布

程序启动后会检查：

```text
http://hr.seedlingintl.com/api/static/hr-toolkit/latest.json
```

也可以在 `HRToolkit.exe` 同目录放一个 `update_url.txt`，第一行写 Gitee 上的 `latest.json` 地址。程序会优先读取这个文件，例如：

```text
http://hr.seedlingintl.com/api/static/hr-toolkit/latest.json
```

如果发现新版本，用户必须点击更新；取消会直接退出程序。下载完成后会启动 `HRToolkitUpdater.exe`，关闭主程序，替换整个 `HRToolkit` 目录，再自动重新打开。更新过程不会弹出黑色命令窗口。主界面右上角也有“检查更新”，可手动检查。

更新失败时，程序会在 `HRToolkit` 文件夹同级目录写入 `HRToolkit_update.log`。这个日志会记录下载包路径、解压目录、备份目录、替换步骤和具体错误，方便直接定位是哪一步失败。

发布步骤推荐使用一键脚本。脚本会自动递增版本号、打包、计算 SHA256，并生成服务器静态目录，不需要手动改 `latest.json`。

### 通过 Gitee 中转发布

更新包最终放在 ScriptHub 静态目录中。Gitee 只用于让 Windows 打包机把发布文件同步到 Mac。推荐用一键发布脚本，不要手写 `version`、`file_url` 和 `sha256`。

Windows 版必须在 Windows 上执行。日常发布用补丁版本：

```powershell
python scripts\release_windows.py --bump patch --notes "本次更新说明"
```

版本变化：

```text
0.1.0 -> 0.1.1 -> 0.1.2
```

阶段性小版本用：

```powershell
python scripts\release_windows.py --bump minor --notes "完成一批新需求"
```

版本变化：

```text
0.1.9 -> 0.2.0
```

大版本用：

```powershell
python scripts\release_windows.py --bump major --notes "正式版发布"
```

版本变化：

```text
0.9.9 -> 1.0.0
```

一键发布脚本会自动完成这些事：

- 先递增 `hr_toolkit/__init__.py` 里的版本号
- 打包 `HRToolkit.exe`
- 打包 `HRToolkitUpdater.exe`
- 把 `HRToolkitUpdater.exe` 复制进 `dist\HRToolkit\`
- 在 `dist\HRToolkit\` 写入 `update_url.txt`
- 把 `dist\HRToolkit\*` 压缩到 `release/scripthub_static/hr-toolkit/releases/HRToolkit-版本号-win.zip`
- 计算 zip 的 SHA256
- 更新 `release/scripthub_static/hr-toolkit/latest.json` 里的 `version`、`file_url`、`sha256`
- 自动清理历史发布文件，每个平台只保留当前最新版本

发布脚本生成后，会得到：

```text
release/scripthub_static/hr-toolkit/
  latest.json
  releases/
    HRToolkit-版本号-win.zip
```

把 `release/scripthub_static/hr-toolkit/` 这个文件夹复制到 ScriptHub 项目的：

```text
fastApiProject/static/
```

最终目录应为：

```text
fastApiProject/static/hr-toolkit/latest.json
fastApiProject/static/hr-toolkit/releases/HRToolkit-版本号-win.zip
```

如果旧版本更新失败后只剩 `HRToolkit_backup_时间`，先把这个文件夹改名回 `HRToolkit`，再重新打开 `HRToolkit.exe` 检查更新。正常版本不需要单独下载更新器，更新器已经包含在 zip 更新包里。

如果打包机上也有 ScriptHub 项目，可以直接让脚本复制过去：

```powershell
python scripts\release_windows.py --bump patch --notes "本次更新说明" --publish-dir "E:\path\to\Nexus-Scripts\fastApiProject\static\hr-toolkit"
```

4. 提交并推送：

```bash
git add hr_toolkit/__init__.py release/scripthub_static/hr-toolkit/latest.json release/scripthub_static/hr-toolkit/releases/
git commit -m "发布 HR工具箱 0.2.0"
git push gitee main
```

推送 Gitee 后，Mac 拉取仓库，再把 `release/scripthub_static/hr-toolkit/` 复制到服务器的 `fastApiProject/static/` 下。客户端真正下载的 zip 会走 ScriptHub：

```text
http://hr.seedlingintl.com/api/static/hr-toolkit/releases/HRToolkit-版本号-win.zip
```

只推源码不会触发客户端更新。客户端只看 `latest.json` 里的 `version` 是否大于当前程序版本，并按 `file_url` 下载 zip。
