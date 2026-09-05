# HR Toolkit

人事 Excel 自动化工具箱。当前已落地：
- **需求1：社保明细与汇总**
- **需求2：考勤与周月报统计**
- **需求3：保险台账与增减预警**
- **需求4：工资表按入职公司拆分**
- **需求5：多月工资合并个人薪资汇总**
- **需求6：异动表汇总与花名册更新**
- **需求7：档案移交表入库与导出**
- **需求8：人员资料文件夹改名**
- **需求9：员工资料收集与本地离线 OCR 打包**

---

## 已实现工具

### 需求1-社保明细与汇总

输入社保账户缴费清单、常见压缩包，或一个包含多个社保清单/压缩包的文件夹，再选择参保人员花名册，按身份证关联人员信息，输出社保明细表和社保汇总表。输入支持 `.xlsx`、`.xls` 以及 ZIP、RAR、7Z、TAR 压缩包。

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

---

### 需求2-考勤与周月报统计

输入 HR 系统导出的考勤结果、周报记录、月报记录，支持单个文件、多个文件、常见压缩包，或包含这些文件/压缩包的文件夹，自动生成考勤和周月报统计表。输入支持 `.xlsx`、`.xls` 以及 ZIP、RAR、7Z、TAR 压缩包。

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

---

### 需求3-保险台账与增减预警

输入各保单人员清单、常见压缩包，或一个包含多个保单清单/压缩包的文件夹，再选择需求6的人力资源分析表，自动生成保险台账。输入支持 `.xlsx`、`.xls` 以及 ZIP、RAR、7Z、TAR 压缩包。

输出内容：

- 自动识别保单号和保单人员
- PZDX 类保单的保额取 `每人伤残死亡限额`，按万元显示
- PEAC 类保单没有明显保额字段，固定按 `60` 万元显示
- `项目/部门` 从人力资源分析表的 `花名册` 工作表补充，优先取 `部门/项目` 列
- 生成人员增减预警：花名册在职但保单没有提示 `需加保`，保单有但花名册没有或已标记离职提示 `需减保`
- 输出 `保险台账.xlsx`，包含 `保险台账` 和 `人员增减预警` 两个工作表
- 如存在 `需加保` 人员，会额外输出 `人力资源分析表_保险预警.xlsx`，在 `花名册` 中标记 `保险预警`

人事已确认岗位保险规则暂取消，当前只做台账明细和人员增减预警。

---

### 需求4-工资表按入职公司拆分

输入一个包含 `汇总表`、`明细表` 的工资表，按明细表中的 `入职公司` 字段拆分为多个 Excel 工作簿。输入支持 `.xlsx` 和 `.xls`。

输出内容：

- 每个入职公司一个独立 `.xlsx`
- 保留原工资表的主要样式、表头、公式结构
- 明细表只保留对应公司的员工行
- 明细表保留原模板分段小计和底部总计文案
- 汇总表引用拆分后明细表中的分段小计

---

### 需求5-多月工资合并个人薪资汇总

输入单个月度工资表、多个工资表、常见压缩包，或一个包含多个月度工资表/压缩包的文件夹，按 `身份证号码` 合并，输出每个人一行的个人应发工资汇总表。也可以同时选择已有汇总表，工具会把新月份追加进去。输入支持 `.xlsx`、`.xls` 以及 ZIP、RAR、7Z、TAR 压缩包。

明细表的应发金额支持 `应发小计` 和 `本月应发工资`（同时存在时优先使用前者）。优先读取 Excel 已保存的计算结果；新格式拆分后无缓存时，支持按原公式 `ROUND(SUM(本行收入区域)-SUM(本行扣款区域),2)` 计算，包括输入项中的本行引用和基本四则运算。无法安全计算的公式会提示用 Excel/WPS 重新计算并保存，不会套用旧模板列号。月份仍优先从文件名识别，文件名无月份时读取表内日期，请保持两者一致。

Qt 桌面端支持客户自行设置列头对应关系：

1. 选择工资表后点击开始合并。已支持的模板自动处理；列名变化、重名列或已保存模板结构变化时，先显示“确认工资列头对应关系”。
2. 在下拉列表中选择原表的姓名、身份证号码、应发工资列，并核对列字母、列头和样例。例如可将 `应发小记`、`本月应发合计` 对应到“应发工资”。已有汇总表可单独设置姓名和身份证号码列，月份列仍按原有规则识别。
3. 同类模板只需确认一次。勾选“记住当前项目的对应关系”后，下次相同模板复用设置；不同项目分别保存。只保存列头规则，不保存样例或工资数据。取消勾选则仅在本次使用。
4. 列头列表不正确时，可展开“选择工作表与表头位置”重新读取。支持前 200 行内连续最多 6 行的表头，展示前 512 列的非空列名。
5. 需要调整或删除规则时，选择相应模板文件，点击主界面的“列头对应设置”。可修改对应列，或点击“恢复自动识别”，再“保存设置并关闭”；也可保存后继续合并。

桌面端不会静默跳过未确认的模板。确实不需要合并的文件，应明确勾选“本次不合并”，结果会增加“未合并文件”工作表供核对。列头确认后原文件发生变化时，需要重新确认。自定义对应关系只决定读取哪一列，不修改原表，也不改变金额含义；没有有效计算结果且无法安全计算的公式，仍需先在 Excel/WPS 中重新计算并保存。命令行保留原有调用方式，不读取桌面端的项目设置。

输出内容：

- 自动识别每张工资表的月份
- 按身份证号码合并同一员工
- 已有汇总表中已经存在的人员月份不覆盖，避免重复写入
- 新月份中出现的新员工会自动新增一行
- 人员在某个月没有工资时自动填 `0`
- 输出 `个人薪资汇总表.xlsx`

---

### 需求6-异动表汇总

输入单个项目异动表、多个项目异动表、常见压缩包，或一个包含多个项目异动表/压缩包的文件夹，将各项目填写的 `增补表`、`离职`、`转正`、`调整` 按记录日期分到对应月份汇总表。文件夹里如果同时放入人力资源分析表，工具会同步更新其中的 `花名册`。输入支持 `.xlsx`、`.xls` 以及 ZIP、RAR、7Z、TAR 压缩包。

输出内容：

- 支持项目表中的 `增补表`、`离职`、`转正`、`调整`
- ZIP、RAR、7Z、TAR 会自动识别并解压，文件夹内的压缩包也会自动处理
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

---

### 需求7-档案移交表入库

输入项目部提交的人事档案移交表，按 `公司` 写入档案汇总表；也可以从一份或多份档案汇总表生成各公司独立档案表。

输出内容：

- 支持单个 `.xlsx/.xls` 移交表、多个移交表、ZIP/RAR/7Z/TAR 压缩包，或包含移交表/压缩包的文件夹
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

---

### 需求8-人员资料文件夹改名

选择一个人员资料目录，对目录下的人员文件夹或指定类型的文件（PDF、图片、文档）做批量改名。执行前会先预览，并二次确认。

支持内容：

- 选择项目类型：文件夹、PDF、图片、文档，或全部
- 批量追加后缀，例如 `张三` -> `张三-劳动合同`（追加文字会按原样写入，建议以 `-` 或 `_` 开头）
- 批量删除结尾文字，例如 `张三_劳动合同`、`李四劳动合同` -> `张三`、`李四`
- 指定单个人员/单文件处理，例如只处理 `张三` 或 `张三.pdf`
- 替换单个名称，例如 `张三` -> `章五`；替换文件时会自动补全原扩展名

---

### 需求9-员工资料收集与本地离线 OCR 打包

选择待整理的原始混乱员工资料总目录，自动按目标员工名单（支持上传花名册 Excel 或手动输入人员信息）进行材料识别、分类和归档打包。

核心功能与特性：

- **本地离线深度学习 OCR 识别**：内置 RapidOCR + ONNXRuntime 推理引擎，在本地 CPU 上直接识别图片（JPG/PNG/WEBP/BMP）、扫描件 PDF 及 Word 文档，**不依赖外部网络与云端 API，保障敏感身份信息绝对安全**；
- **智能双重匹配**：
  - 优先按 18 位身份证号码精准匹配；
  - 结合姓名模糊匹配、拼音容错匹配，从复杂的合同、体检表、证件照片中提取所属员工；
- **多类别材料自动分类**：自动归类身份证（正反面）、体检报告、劳动合同、银行卡、学历证书、离职证明等材料；
- **规范输出**：生成按员工姓名命名的独立文件夹或统一 ZIP 归档包，并输出详细的识别与匹配报告日志。

---

### Excel 旧格式兼容

- 已实现的 Excel 类工具均支持上传 `.xlsx` 和 `.xls`
- 文件夹和 ZIP/RAR/7Z/TAR 压缩包中也会识别 `.xls`
- 输出文件统一为 `.xlsx`
- 需求1的老 `.xls` 社保清单和参保花名册会用内置依赖直接读取
- 其他工具遇到 `.xls` 会先自动转换为 `.xlsx` 再处理；Windows 电脑需要安装 Excel 或 WPS 表格，Mac/Linux 需要安装 LibreOffice 才能自动转换

---

## 桌面版使用

无参数启动时会打开图形界面：

```bash
python3 -m hr_toolkit
```

界面操作流程：

1. 首次使用时通过“新建工作项目”填写名称并确认保存位置，或打开电脑上已有的项目文件夹
2. 在左侧选择工具，再从电脑添加文件、压缩包或文件夹
3. 外部资料会先复制到当前项目的“上传资料”，工具只处理项目内的副本
4. 点击 `开始拆分`、`开始合并`、`开始汇总`、`开始入库` 或 `开始提取`
5. 正式结果直接保存到当前项目的“处理结果”，不再额外保存一份隐藏结果副本
6. 在右侧项目文件区查看上传资料、处理结果和以前的处理批次

### 项目文件工作区与追溯

一个项目可以使用全部人事工具。项目内会先按业务、工具和处理批次整理，再在每个批次内分别保存“上传资料”和“处理结果”；没有使用过的工具不会预先创建空文件夹。

```text
项目名称/
├─ 业务分类/
│  └─ 工具名称/
│     └─ 处理批次/
│        ├─ 上传资料/
│        ├─ 处理结果/
│        └─ 补充资料/   # 有补充文件时才创建
├─ 共用资料/
└─ .hrtoolkit/    # 隐藏的项目标记、批次清单、锁和回收站
```

- 新建项目弹窗会同时显示最终位置、留存内容和位置风险；活动项目只允许放在本机普通文件夹，网盘、共享盘、NAS、U 盘和移动磁盘请仅在退出工具后用于备份
- 项目位置由用户选择，项目资料与程序安装目录分开，正常更新或重新安装不会自动删除项目
- 从电脑其他位置选择的源文件默认复制进项目，不会长期引用外部路径；源文件被移动或删除后，项目仍可追溯
- 项目内的“处理结果”是正式版本，工具不再向隐藏目录重复复制；如需发送，可从右侧打开后复制文件
- 右侧“添加”支持从电脑导入文件、导入文件夹和新建文件夹；业务分组和工具目录由系统维护，普通人工资料放在“共用资料”；切到“当前功能”时需先选中具体批次的“上传资料”或“补充资料”
- 导入时会依次显示“检查资料、复制并校验、完成保存”；前两阶段可以安全取消，进入最后登记阶段后会短暂禁止取消，避免留下半批资料
- 已开始或已完成批次的“上传资料”和“处理结果”保持不变；后续合同、签字版等人工材料放进按需创建的“补充资料”
- 项目内以前的上传资料、补充资料、共用资料和已完成结果，都可以复制为新批次的输入；批次资料必须仍与原清单一致，每次处理会再保留一份独立快照
- 每个已登记文件都会记录大小和 SHA-256 校验值；复制未完成、源文件中途变化、空间不足或程序异常退出时，不会显示为成功
- 文件复制、校验和索引在后台分批进行；项目文件区每次启动默认收起，展开后文件树只读取当前层级，避免大项目拖住界面
- 每个批次和普通资料导入都有独立恢复记录；如果复制、登记、批次改名、移入回收站或从回收站恢复时程序意外退出，下次打开项目会按记录安全补完或明确停止，不会覆盖冲突文件；更高版本的项目只读打开，不会擅自改写
- 所有受支持压缩包在解压前都会检查文件数、单文件大小、总大小、重复路径、链接和异常压缩比例，超过安全限制会停止处理
- 输入压缩包后缀包括 `.zip`、`.rar`（RAR/RAR5）、`.7z`、`.tar`、`.tar.gz/.tgz`、`.tar.bz2/.tbz2`、`.tar.xz/.txz`；当前按普通单卷、无密码压缩包处理，加密包会明确提示而不会产生不完整结果
- `移到回收站` 只对完整处理批次生效，会把该批次的上传资料、处理结果和补充资料一起移入项目内部回收站，不会立即永久删除；可在右侧“项目回收站”中搜索并恢复，同名位置会使用新名称且不覆盖现有资料
- 项目功能不会主动上传资料；如需把完整项目备份到 OneDrive、企业网盘、NAS 或 U 盘，请先退出 HR Toolkit，再按公司的数据安全规定复制项目文件夹

需要迁移或备份时，请先退出 HR Toolkit，再完整复制项目文件夹（包括隐藏的 `.hrtoolkit`）；不要单独编辑或复制正在使用的项目元数据和批次清单。

---

## 本机 CLI 验证与调用

当前目录执行：

```bash
python3 -m pip install -r requirements.txt
```

现代版正式 CI 与安装包使用已经验证的 Python 3.12 精确依赖集合；Windows 7 SP1 x64 兼容包使用单独冻结的 Python 3.8 依赖集合，不会改变现代版依赖。需要在本机复现现代生产环境时执行：

```bash
python3.12 -m pip install \
  -r requirements.txt \
  -c constraints/python312-production.txt
```

从源码运行桌面界面还必须安装 Qt 依赖；只安装 `requirements.txt` 仅支持 CLI，不会再静默打开性能较低的旧版 Tk 界面：

```bash
python3.12 -m pip install \
  -r requirements.txt \
  -r requirements-gui.txt \
  -c constraints/python312-production.txt
python3.12 hr_toolkit_app.py
```

Windows 7 SP1 x64 源码环境必须使用 Python 3.8，并把约束文件替换为 `constraints/python38-win7.txt`。正式安装包已内置对应 Qt 运行库，终端用户无需安装 Python 或执行上述命令。

Python 3.13 本地源码环境使用新版 `rapidocr==3.9.2`、ONNX Runtime 和对应 Qt 运行库：

```bash
python -m pip install -r requirements.txt -r requirements-gui.txt
python -m hr_toolkit
```

此环境不要套用 `constraints/python312-production.txt`。Python 3.12 正式构建与 Win7 继续使用原有依赖及 OCR 后端。新版接口使用轻量 PP-OCRv4 模型，首次缺少模型时由引擎下载，之后本地复用；切换 OCR 后端会重新建立识别缓存。无法识别人名的合同仍需核对归属。运行记录支持鼠标选择、系统复制快捷键、右键复制及“复制全部”（复制当前列表保留的最多 1000 条记录）。

资料确认：开启缓存后，在《资料待确认.xlsx》的“确认员工”列填姓名、“已核对归属及完整性”列填“是”，同组填写一行，保存关闭后重新打包。重试生成新表后，近期旧表中的填写也会读取；空白新表不会覆盖确认。记录最多保留 16 份清单、合计 20000 条资料引用，达到上限时会提示使用最新清单。程序只读取登记过的清单，跳过生成后未修改的表；归属冲突或原图变化仍须重新核对。

### 1. 社保明细与汇总：
```bash
python3 -m hr_toolkit social-security \
  --input "社保原始清单文件夹" \
  --roster "参保人员花名册.xlsx" \
  --output "outputs/social_security_demo"
```

### 2. 考勤与周月报统计：
```bash
python3 -m hr_toolkit data-statistics \
  --input "考勤数据文件夹" \
  --output "outputs/data_statistics_demo"
```

如有人事提供的应汇报人员名单，可追加：
```bash
python3 -m hr_toolkit data-statistics \
  --input "考勤数据文件夹" \
  --staff "应汇报人员名单.xlsx" \
  --output "outputs/data_statistics_demo"
```

只统计指定日期范围内周一截止的周报（两个日期需同时提供）：
```bash
python3 -m hr_toolkit data-statistics \
  --input "考勤数据文件夹" \
  --week-start 2026-06-02 \
  --week-end 2026-06-30 \
  --output "outputs/data_statistics_demo"
```

### 3. 保险台账：
```bash
python3 -m hr_toolkit insurance-ledger \
  --input "保单文件夹" \
  --roster "人力资源分析表.xlsx" \
  --output "outputs/insurance_ledger_demo"
```

### 4. 工资表拆分：
```bash
python3 -m hr_toolkit salary-split \
  --input "月度薪资表.xlsx" \
  --output "outputs/salary_split_demo"
```

预览模式，不生成文件：
```bash
python3 -m hr_toolkit salary-split \
  --input "月度薪资表.xlsx" \
  --output "outputs/salary_split_demo" \
  --dry-run
```

系统集成时建议使用 JSON 输出：
```bash
python3 -m hr_toolkit salary-split \
  --input "月度薪资表.xlsx" \
  --output "outputs/salary_split_demo" \
  --json
```

### 5. 多月工资合并：
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

### 6. 异动表汇总：
```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --output "outputs/change_merge_demo"
```

追加到已有异动汇总表：
```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "已有异动汇总表.xlsx" \
  --output "outputs/change_merge_demo"
```

指定人力资源分析表并同步更新花名册：
```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "已有异动汇总表.xlsx" \
  --analysis-template "人力资源分析表.xlsx" \
  --output "outputs/change_merge_demo"
```

只用已有异动汇总表单独更新花名册：
```bash
python3 -m hr_toolkit roster-update \
  --input "已有月度汇总表文件夹" \
  --roster "人力资源花名册.xlsx" \
  --output "outputs/roster_update_demo"
```

### 7. 档案移交表入库：
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

按公司生成独立档案表：
```bash
python3 -m hr_toolkit archive-export \
  --summary "档案表汇总表.xlsx" \
  --output "outputs/archive_export_demo"
```

### 8. 人员资料文件夹改名：
```bash
# 预览
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode append \
  --text "劳动合同"

# 执行
python3 -m hr_toolkit folder-rename \
  --root "人员资料目录" \
  --mode append \
  --text "劳动合同" \
  --apply
```

### 9. 员工资料离线 OCR 提取打包：

按人员文件夹查找（默认，保持原有行为）：
```bash
python3 -m hr_toolkit material-collector \
  --library "按人员整理的资料库" \
  --roster "目标员工名单.xlsx" \
  --output "outputs/material_collector_demo"
```

无序平铺资料库按 OCR 内容建立全局索引：
```bash
python3 -m hr_toolkit material-collector \
  --library "无序混放的资料库" \
  --roster "目标员工名单.xlsx" \
  --output "outputs/material_collector_flat_demo" \
  --library-mode flat_ocr
```

`flat_ocr` 首次运行会在资料库根目录写入隐藏索引缓存；后续查询会复用内容未变化的文件。源文件不会被修改。

### 精确 A/B 回归验收

修改 Excel 处理或文件处理实现前，先保存“修改前输出 + 客户源附件”的本地基线：

```bash
python3 scripts/compare_regression_outputs.py capture \
  --outputs "/path/to/before-output" \
  --sources "/path/to/source-copy" \
  --manifest "/tmp/hr-toolkit-baseline.json"
```

使用同一批输入生成修改后结果，再做验收：

```bash
python3 scripts/compare_regression_outputs.py verify \
  --baseline "/tmp/hr-toolkit-baseline.json" \
  --outputs "/path/to/after-output" \
  --sources "/path/to/source-copy" \
  --report "/tmp/hr-toolkit-regression-report.json"
```

工具会逐项比较工作表顺序、单元格值、公式及缓存计算值、格式、批注、链接、行列尺寸、合并区域、命名区域、打印区域与页眉页脚、验证规则、条件格式、表格、图表和嵌入资源；非 Excel 产物及全部源附件按字节大小和 SHA-256 验证。工具不会修改被检查的目录。基线会包含 Excel 业务值及附件哈希，应只保存在受控临时目录，不要提交到 Git 或对外发送。

---

## 自动构建与发布

日常发布只在本地 Mac 做版本检查、测试、版本提交、annotated Tag 和原子推送；Windows、macOS 构建与 GitHub Release 发布全部交给 GitHub Actions。正式发布命令为：

```bash
npm run release -- 0.3.5
```

首次使用先安装 Python 依赖和 Node.js。npm 入口会优先使用 `.venv/bin/python`，不存在时再使用 `python3`；两者之一必须能运行完整测试。发布前可执行不修改版本文件、commit、Tag 或远端的演练：

```bash
npm run release -- 0.3.5 --dry-run
```

无人值守环境审核完版本后可追加 `--yes`。发布脚本会严格检查 stable SemVer、clean `main`、`HEAD == origin/main`、本地/远端 Tag 冲突以及全部版本字段，并通过已登录的 GitHub CLI (`gh`) 确认当前精确提交的完整 CI 已成功；CI 尚未开始、运行中、失败或取消时都不会创建版本提交和 Tag。通过远端门禁后再运行 `unittest`、`compileall` 和 `git diff --check`。正式执行只会暂存 `hr_toolkit/__init__.py`、`package.json`、`package-lock.json`，不会运行本地跨平台构建，也不会使用 `git add .`。

脚本创建单一版本提交和 `v<version>` annotated Tag，再通过一次 atomic push 同时推送 `main` 与 Tag。推送失败时只有在确认远端两个引用都未变化后才自动回滚；远端状态不明确时会保留现场，要求人工核对。

> 不要为了测试发布脚本在正式仓库创建 Tag。先使用 `--dry-run`；正式命令必须等发布负责人确认。

### GitHub Actions 产物

普通 push 和 pull request 由 `.github/workflows/ci.yml` 运行测试、编译和静态发布检查：现代 Windows 使用 Python 3.12，Win7 兼容栈使用 Python 3.8，并分别实际加载 OCR 与压缩包运行时。`.github/workflows/test-build.yml` 每周一北京时间 02:00 同时构建现代 Windows 与 Win7 SP1 兼容包，也可手动选择单个平台；选择 macOS 时会在真实 Intel 与 Apple Silicon runner 上分别构建和验收。现代构建使用 `constraints/python312-production.txt`，Win7 构建使用 `constraints/python38-win7.txt`，两套依赖互不替换。

只有 `v*` Tag 会触发 `.github/workflows/release.yml`：先校验 Tag 与 `hr_toolkit.__version__` 完全一致，再独立构建现代 Windows、Win7 SP1 x64 和 macOS；全部成功后才创建并发布 GitHub Release。GitHub Release 发布成功后，独立的 `mirror-gitee` job 才会把同一份源码、annotated Tag，以及 `SHA256SUMS.txt`、`latest.json` 和两种 Windows `setup.exe` 同步到 Gitee。MSI 与两个 DMG 按镜像精简策略仅保留在 GitHub；所有正式二进制资产发布前都必须严格小于 100,000,000 字节。

每个版本的直接下载资产为（以下用 `<version>` 表示版本号）：

```text
HRToolkit_<version>_arm64.dmg
HRToolkit_<version>_x64.dmg
HRToolkit_<version>_x64-setup.exe
HRToolkit_<version>_x64.msi
HRToolkit_<version>_win7_x64-setup.exe
HRToolkit_<version>_win7_x64.msi
latest.json
SHA256SUMS.txt
```

当前 OCR 运行时包含架构专属原生二进制，因此正式发布直接并行构建上表中的两个真实架构文件。构建脚本仍保留经 `file`/`lipo` 严格验证的 universal2 能力，但发布流水线不会先执行注定失败的跨架构尝试，也不会把单架构程序改名伪装成 `universal`。DMG 内包含标准 `HRToolkit.app` 和指向 `/Applications` 的快捷方式。

### Windows 安装与自动更新架构

Windows 构建输出：
- `HRToolkit_<version>_x64-setup.exe`：原现代版安装包，继续使用 Python 3.12，支持 Windows 8.1 x64 及更高版本，文件名和更新入口保持不变；
- `HRToolkit_<version>_x64.msi`：WiX v4 生成的 per-user MSI 安装器，供企业 IT 批量分发；
- `HRToolkit_<version>_win7_x64-setup.exe`：供 Windows 7 SP1 x64 使用的 Python 3.8 兼容安装包（同时覆盖不能运行 Python 3.12 的 Windows 8/Server 2012），OCR 固定使用 Win7 兼容候选 ONNX Runtime 1.11.1，并内含固定为 Windows SDK 10.0.14393.795 的完整 x64 app-local UCRT、Win7 可用的 Visual C++ 2015-2019 运行库和固定版本的官方 7-Zip；UCRT/VC DLL 会按微软的旧系统部署要求放在主 EXE 同目录；
- `HRToolkit_<version>_win7_x64.msi`：Win7 兼容版的企业分发安装器；
- `HRToolkitUpdater.exe`：内置独立更新器。

两种 EXE 都安装到当前用户目录 `%LOCALAPPDATA%\Programs\HRToolkit`，普通用户双击安装无需管理员提权。同一台电脑只安装其中一种。Win7 用户首次切换时必须手动安装一次带 `win7_x64` 的安装包；之后程序会固定使用 `windows-x64-win7` 更新通道。不要再从网上单独下载或复制 `api-ms-win-core-path-l1-1-0.dll` 到程序目录，兼容包已经携带经过校验的运行时文件。

Win7 与 Python 3.8 均已停止官方维护，因此兼容分支只加载程序随包附带的可信 OCR 模型，不提供用户自定义 ONNX 模型入口；任何原生依赖升级都必须重新通过 PE 扫描和干净 Win7 虚拟机验收，现代分支不跟随该冻结版本降级。

发布前除自动执行现代环境全量测试、Python 3.8 核心兼容测试、两种 Windows 安装测试和所有 PE 导入表检查外，还必须在干净的 Windows 7 SP1 x64 虚拟机中运行 `scripts/smoke_win7_installer.ps1 -InstallerPath <安装包> -ExpectedVersion <版本号>`。该脚本会完成静默安装、主程序与更新器真实启动、OCR/压缩包/项目运行检查、运行时白名单校验和卸载；未取得这一步的通过结果，不能把“静态兼容”表述为“已在 Win7 实机验证”。

**自动更新机制**：
- Windows 客户端在检测到新版本后，按当前系统选择现代版或 Win7 专用 `setup.exe`；Win7 通道不会回退到现代包；
- 下载完成后唤起 `HRToolkitUpdater.exe` 并退出主程序；
- 更新器在后台通过静默指令 `setup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART` 秒级完成程序目录覆盖并自动重新打开新版本；
- 保留对旧版 ZIP 解压更新的向后兼容。

### 国内更新源、GitHub 回退与容灾

客户端按以下顺序检查公开更新源：

```text
1. https://gitee.com/api/v5/repos/optimistic-little-sunspot/hr-toolkit/releases/latest
2. https://github.com/xhzwjc/hr-toolkit/releases/latest/download/latest.json
```

- Gitee 最新 Release 接口会返回公开附件列表，客户端从中找到 `latest.json`。只要 Gitee 能返回有效配置，就不会访问 GitHub；
- Gitee 连接失败、超时或配置异常时，自动回退尝试 GitHub；
- `latest.json` 同时提供 `windows`/`windows-x64-modern` 与 `windows-x64-win7`；两种 Windows EXE 均优先使用 Gitee 国内 CDN，下载失败时自动重试 GitHub；
- macOS 的 DMG 清单使用 GitHub 下载地址；Gitee Release 按镜像精简策略只保留校验文件、更新清单和两种 Windows EXE；
- 下载弹窗实时显示进度与速率，支持点击右上角 `×` 或按 `Esc` 随时安全取消，并自动清理未完成的临时文件。

---

## 运行日志与数据隐私

- **全离线计算**：所有数据统计、表格拆分合并、OCR 图像识别均在用户本机内存与本地磁盘中运行，不连接任何第三方 AI 云端或外部数据服务器。
- **脱敏日志**：
  - 运行日志统一输出至同级目录下的 `HRToolkit_app.log` 与 `HRToolkit_update.log`；
  - **仅记录工具启动、耗时、处理文件数量与错误堆栈，严禁记录任何真实表格内容（身份证号、手机号、薪资金额等敏感信息绝不上报与外显）**；
  - 日志设置自动滚动截断限制（最大 1MB），防止长期占用磁盘空间。
