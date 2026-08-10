# HR Toolkit 项目工作区设计验收

**Source visual truth**

- `/private/tmp/hr-toolkit-requirements-audit/01-workspace.png`
- 原型像素：1440 × 900

**Rendered implementation**

- `/var/folders/ng/s5d0j44s06z6zdc_qr9ylpv80000gn/T/com.openai.sky.CUAService/Python Screenshot 2026-08-03 at 6.27.59 PM.jpeg`
- 实现截图像素：1187 × 768
- 对照图：`/private/tmp/hr-toolkit-ui-comparison.png`

**Viewport and normalization**

- Tk 窗口请求尺寸：1440 × 900；受当前 Mac 可用屏幕区域限制，实际整窗截图为 1187 × 768。
- 对照时将原型按比例缩小后居中到 1187 × 768 白底面板；实现截图保持原始 1187 × 768，未拉伸。
- 桌面原生 Tk 界面没有 CSS viewport 或浏览器 `deviceScaleFactor`；本次按最终屏幕像素比较。

**State**

- 当前功能：社保明细与汇总。
- 当前项目：已打开可写项目。
- 右侧范围：全部文件。
- 实现截图中的目录树为根目录折叠态；原型为示例资料已选中、目录树展开态。因此不把文件行数量和展开层级差异计为视觉缺陷，只比较三栏结构、层级、间距、颜色、文字和控制位置。

## Findings

- 没有遗留 P0、P1 或 P2 问题。
- [P3] 右侧原型为文件和文件夹提供了更丰富的图标，当前原生 `Treeview` 主要依赖展开箭头和文字。
  - Location: 右侧“项目文件”目录树。
  - Evidence: 原型为文件夹和文件显示独立图标；实现保留系统展开箭头、结果色标和文字层级。
  - Impact: 不影响查找、展开、打开或定位，视觉识别速度略低。
  - Follow-up: 后续如提供正式跨平台图标资产，可为目录、Excel、ZIP、PDF 增加统一图标；本次不使用字符或手绘图标替代。

## Required fidelity surfaces

- Fonts and typography: 延续现有系统中文字体，标题、分组、正文和辅助信息层级清楚；修正后表单标签完整显示，无裁切。项目长路径仅在右侧显示，左侧改为“当前项目 · N 次处理”。
- Spacing and layout rhythm: 三栏结构、固定左导航、主操作区和右项目区成立；卡片间距、圆角和边界沿用现有 HR Workbench 设计语言。右区支持 270–430 宽度调节及折叠轨道。
- Colors and visual tokens: 暖纸底色、白色卡片、深绿色主色和灰色辅助文字与原型一致；成功、警告和禁用状态使用现有语义色。
- Image quality and asset fidelity: 原型没有照片、插画或品牌图片资产；实现没有用占位图片替换目标资产。目录图标差异列为 P3。
- Copy and content: “结果位置”显示为“当前项目 / 本次处理结果”，避免向人事暴露长路径；“处理结果”“上传资料”“补充资料”“共用资料”等名称与需求一致。
- Responsiveness and accessibility: 主区变窄时表单切换纵向布局；右区在窄窗口进入抽屉模式。按钮均保留可读文字，禁用状态可见，正文对比度符合现有桌面主题。
  抽屉支持 `Esc` 关闭，折叠轨道支持鼠标、回车和空格重新展开。

## Focused region evidence

- 对照图中重点检查了中心表单和右侧项目区。第一轮发现中心“参保人员花名册”“结果位置”标签被压缩为单字，且长磁盘路径降低可读性；修正后的实现截图显示完整标签和业务化结果位置。
- 右侧区域在全视图中已经能清楚判断标题、项目身份、范围切换、搜索、添加、刷新、目录树和详情区，无需额外裁切图。

## Comparison history

1. 第一轮实现截图：`Python Screenshot 2026-08-03 at 6.24.21 PM.jpeg`
   - [P2] 非紧凑表单第一列宽度不足，两个字段标签只显示首字。
   - [P2] 结果位置直接显示很长的绝对路径，不符合人事用户的阅读方式。
2. 修正：
   - 为非紧凑表单标签列增加 116 像素最小宽度，并提高紧凑布局切换阈值。
   - 结果位置改为业务文案“当前项目 / 本次处理结果”，实际路径仍在内部用于运行。
   - 左侧项目卡片改为处理次数摘要，完整路径保留在右侧项目区。
3. 第二轮实现截图：`Python Screenshot 2026-08-03 at 6.27.59 PM.jpeg`
   - 表单标签和业务化结果位置已完整显示，未发现新的 P0/P1/P2 视觉问题。
   - 后续收起态和窄窗口追加截图因 Mac 锁屏未能继续；相关状态已由 Tk 隐藏窗口 smoke、方法绑定和响应式分支检查覆盖，作为残余测试缺口记录，不阻塞主视图验收。
4. 发布前安全与交互复核：
   - 资料导入期间禁止切换项目或开始处理；退出时先取消并等待安全落盘，再释放项目写锁。
   - 重复打开当前项目保持原可写实例；切换项目会清空上一项目的输入和“打开结果”状态。
   - 链接输入、未登记或已改动的历史结果、处理期间插入的链接均会被拒绝。
   - Windows 与 macOS 打包检查同时阻止 `.hrtoolkit`、上传资料、处理结果、补充资料和共用资料进入安装包。
   - 317 项自动化测试、强制编译检查和差异格式检查全部通过。

## Implementation checklist

- [x] 三栏工作区与右侧项目文件区。
- [x] 项目切换、打开文件夹、全部文件/当前功能、搜索、添加、刷新。
- [x] 右侧收起/展开、宽度记忆和窄窗口抽屉逻辑。
- [x] 表单标签裁切修复和人事化结果位置文案。
- [x] 上传资料、正式结果、补充资料和项目复用链路。
- [ ] P3：获得正式跨平台文件类型图标资产后再增强目录树图标。

**final result: passed**

---

## 项目弹窗追加验收（2026-08-10）

**Source visual truth**

- 新建工作项目：`/tmp/hr-toolkit-figma-reference/new-project.png`
- 项目回收站：`/tmp/hr-toolkit-figma-reference/recycle.png`
- 检查资料：`/tmp/hr-toolkit-figma-reference/import-checking.png`
- 复制并校验：`/tmp/hr-toolkit-figma-reference/import-progress.png`
- 完成保存：`/tmp/hr-toolkit-figma-reference/import-finalizing.png`

**Rendered implementation**

- 新建工作项目：`/tmp/hr-toolkit-final-visual-qa.P0L0WO/01-new-project-current.png`
- 检查资料：`/tmp/hr-toolkit-final-visual-qa.P0L0WO/02-import-checking-current.png`
- 复制并校验：`/tmp/hr-toolkit-final-visual-qa.P0L0WO/03-import-copying-current.png`
- 正在完成保存：`/tmp/hr-toolkit-final-visual-qa.P0L0WO/04-import-finalizing-current.png`
- 修复后的项目回收站：`/tmp/hr-toolkit-figma-reference/current-recycle-fixed.png`
- 回收站并排对照：`/tmp/hr-toolkit-figma-reference/recycle-comparison-fixed.png`

**Findings**

- 没有遗留 P0、P1 或 P2 问题。
- 新建项目弹窗完整展示名称、保存位置、最终路径、留存内容、位置风险和操作按钮；长路径会换行，内容较高时可滚动。
- 导入弹窗的实际内容区为 640 × 510，检查、复制和最后保存三种状态均无裁切；前两阶段可取消，最后安全登记阶段禁用取消。
- 回收站曾出现底部操作按钮被裁切的发布阻断问题。修复后弹窗为 760 × 487，关闭和恢复按钮完整显示，底部保留 20 像素空间；小屏只压缩可滚动的中间列表区。
- 回收站使用业务事项作为主标题，单列业务分组和功能；恢复位置不显示带时间戳的内部目录，同名保护使用提醒色。
- macOS 原生标题栏和无背景蒙层与 Figma 有视觉差异，属于桌面平台惯例，不影响操作和信息层级。

**Scope decision**

- 原型中的“直接修改原文件夹”风险确认页没有实现。当前冻结范围明确不直接修改外部原文件；资料文件夹改名只处理项目内经过校验的结果副本，因此该风险分支已被产品方案取消，不属于漏做弹窗。

**Verification**

- 317 项自动化测试在严格 `ResourceWarning` 模式下全部通过。
- 项目创建、回收站、导入取消、最后保存恢复、只读、同名不覆盖、链接拒绝和打包防资料泄漏均有定向覆盖。
- 全模块强制编译、运行烟测和 `git diff --check` 全部通过。
- macOS 源码版已逐页实机截图验收；Windows 安装包仍需在真实 Windows 100% / 125% / 150% 缩放下完成发布前人工烟测。

**final result: passed for source and macOS visual QA; Windows packaged smoke pending**

---

## 滚动条样式追加验收（2026-08-03）

**Source visual truth**

- 目标滚动条：`/var/folders/ng/s5d0j44s06z6zdc_qr9ylpv80000gn/T/codex-clipboard-467e6e0b-bb06-4bfc-aebf-f2994e9888a4.png`（63 × 424）
- 原问题全图：`/var/folders/ng/s5d0j44s06z6zdc_qr9ylpv80000gn/T/codex-clipboard-17129da5-7631-40db-a2f6-66cd59a7a6dc.png`（1636 × 855）

**Rendered implementation**

- 最终 Tk 组件截图：`/private/tmp/hrtoolkit-scrollbar-final-window.png`（328 × 640）
- 同尺寸聚焦对照：`/private/tmp/hrtoolkit-scrollbar-final-comparison.png`（126 × 424；左侧为目标，右侧为实现）

**Viewport and normalization**

- 实现使用 macOS Tk 9.0.3、1× UI 比例，组件窗口为 260 × 540 逻辑像素；系统窗口阴影使最终截图为 328 × 640 像素。
- 聚焦对照只保留滚动轨道，将实现轨道纵向归一到 424 像素，横向厚度保持 1×，再各置于 63 × 424 白底区域。
- 两侧内容量和当前位置不同，因此滑块长度与纵向位置不作为样式差异；本次只比较轨道宽度、滑块宽度、圆角、颜色和箭头可见性。

**State**

- 纵向滚动条处于接近顶部状态；横向滚动条同时显示在组件底部。
- 默认、悬停和按下使用三档浅灰色；截图记录默认状态。

**Findings**

- 没有遗留 P0、P1 或 P2 问题。
- 目标约为 12 像素白色轨道、8 像素浅灰圆角滑块；实现采用相同尺寸关系，端部为圆角胶囊，没有可见的旧式三角箭头。
- 原箭头点击区域仍以透明样式保留，拖拽、轨道翻页、逐行点击、滚轮以及内容回调没有改变。

**Required fidelity surfaces**

- Fonts and typography: 本次目标不包含文字，现有字体配置未改。
- Spacing and layout rhythm: 轨道厚度 12 像素、可见滑块约 8 像素；纵向和横向使用同一尺寸体系。
- Colors and visual tokens: 白色轨道；默认滑块 `#EAEAEB`，悬停 `#DCDCDD`，按下 `#CCCCCE`，与目标浅灰层级一致。
- Image quality and asset fidelity: 滑块按当前 UI 比例在内存生成，并使用可拉伸边框保留圆角端帽；1× 下边缘清晰，没有缩放模糊或透明边缘黑线。
- Copy and content: 没有改动任何界面文案或业务内容。

**Focused region evidence**

- 聚焦对照直接并排展示了目标轨道与最终实现轨道；宽度、浅灰色和圆角端帽一致，没有旧式箭头残留。
- 全图中的主区、运行记录和项目文件区均使用未命名的全局 `ttk.Scrollbar`；同一全局样式覆盖 8 个纵向和 1 个横向实例，无需逐个改控件。

**Comparison history**

1. 初始实现使用 `clam` 默认滚动条，14–15 像素宽，带可见三角箭头和立体边框，与目标存在 P2 视觉差异。
2. 修正为全局白色细轨道和浅灰圆角滑块；箭头命中区域透明化，保留原绑定。
3. Tk 8.6.14 与 Tk 9.0.3 实测纵向宽度和横向高度均为 12 像素；逐行点击与拖拽后内容位置均按原逻辑变化。
4. 最终聚焦对照未发现新的 P0/P1/P2 视觉问题。

**Implementation checklist**

- [x] 8 个纵向滚动条统一样式。
- [x] 1 个横向滚动条统一样式。
- [x] 保留箭头点击、拖拽、轨道点击、滚轮和内容联动。
- [x] Tk 8.6、Tk 9.0 运行验证。
- [x] 目标图与最终实现同图对照。

**final result: passed**
