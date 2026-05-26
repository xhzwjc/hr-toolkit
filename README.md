# HR Toolkit

人事 Excel 自动化工具箱。当前已落地：**需求4：工资表按入职公司拆分**、**需求5：多月工资合并个人薪资汇总**、**需求6：异动表汇总**、**需求8：人员资料文件夹改名**。

## 已实现工具

### 需求4-工资表按入职公司拆分

输入一个包含 `汇总表`、`明细表` 的工资表，按明细表中的 `入职公司` 字段拆分为多个 Excel 工作簿。

输出内容：

- 每个入职公司一个独立 `.xlsx`
- 保留原工资表的主要样式、表头、公式结构
- 明细表只保留对应公司的员工行
- 明细表保留原模板分段小计和底部总计文案
- 汇总表引用拆分后明细表中的分段小计

### 需求5-多月工资合并个人薪资汇总

输入一个包含多个月度工资表的文件夹，按 `身份证号码` 合并，输出每个人一行的个人应发工资汇总表。也可以同时选择已有汇总表，工具会把新月份追加进去。

输出内容：

- 自动识别每张工资表的月份
- 按身份证号码合并同一员工
- 已有汇总表中已经存在的人员月份不覆盖，避免重复写入
- 新月份中出现的新员工会自动新增一行
- 人员在某个月没有工资时自动填 `0`
- 输出 `个人薪资汇总表.xlsx`

### 需求6-异动表汇总

输入一个包含多个项目异动表的文件夹，将各项目填写的 `增员`、`减员`、`转正`、`调动`、`奖罚扣补` 汇总到一份异动汇总表。

输出内容：

- 自动读取五类异动 sheet
- 忽略模板中只有预填序号、没有填写内容的空行
- 汇总后重新编排各 sheet 序号
- 保留模板工作簿样式
- 输出 `异动汇总表.xlsx`

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
2. 选择工资表文件、工资表文件夹、异动表文件夹或人员资料目录
3. 保存位置默认在桌面结果目录下，也可以手动选择
4. 点击 `开始拆分`、`开始合并` 或 `开始汇总`
5. 程序会在保存位置下自动创建 `结果_年月日_时分秒` 子文件夹
6. 点击 `打开所在文件夹` 查看结果

## Mac 本机验证

当前目录执行：

```bash
python3 -m pip install -r requirements.txt
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

指定异动表模板：

```bash
python3 -m hr_toolkit change-merge \
  --input-dir "各项目异动表文件夹" \
  --template "问题6-2026年4月异动汇总表（模板）.xlsx" \
  --output "outputs/change_merge_demo"
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

- `hr_toolkit/tools/salary_split.py`：需求4，工资表拆分
- `hr_toolkit/tools/salary_merge.py`：需求5，工资表合并
- `hr_toolkit/tools/personnel_change_merge.py`：需求6，异动表汇总
- `hr_toolkit/tools/folder_rename.py`：需求8，人员资料文件夹改名
- `hr_toolkit/tools/archive_import.py`：需求7，档案移交表入库
- `hr_toolkit/tools/social_security.py`：需求1，社保明细/汇总

CLI 只是入口，核心函数可以直接被 ScriptHub 或 Web 后端调用。

## Windows 打包

命令行版调试包：

```powershell
python -m PyInstaller --name HRToolkit --onedir --console --clean --add-data "README.md;." hr_toolkit_app.py
```

给人事双击使用的桌面版：

```powershell
python -m PyInstaller --name HRToolkit --onedir --windowed --clean --add-data "README.md;." hr_toolkit_app.py
```

自动更新程序：

```powershell
python -m PyInstaller --name HRToolkitUpdater --onefile --console --clean hr_toolkit_updater.py
Copy-Item dist\HRToolkitUpdater.exe dist\HRToolkit\ -Force
```

Mac 打包时把 `;` 改成 `:`，更新程序复制无后缀文件：

```bash
python -m PyInstaller --name HRToolkit --onedir --windowed --clean --add-data "README.md:." hr_toolkit_app.py
python -m PyInstaller --name HRToolkitUpdater --onefile --console --clean hr_toolkit_updater.py
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
https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main/release/latest.json
```

也可以在 `HRToolkit.exe` 同目录放一个 `update_url.txt`，第一行写 Gitee 上的 `latest.json` 地址。程序会优先读取这个文件，例如：

```text
https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main/release/latest.json
```

如果发现新版本，用户必须点击更新；取消会直接退出程序。下载完成后会启动 `HRToolkitUpdater.exe`，关闭主程序，替换整个 `HRToolkit` 目录，再自动重新打开。主界面右上角也有“检查更新”，可手动检查。

发布步骤：

1. 修改 `hr_toolkit/__init__.py` 里的 `__version__`。
2. 按上面的命令打包 `HRToolkit` 和 `HRToolkitUpdater`。
3. 如果使用 Gitee，把 `release/update_url.txt.example` 复制为 `dist\HRToolkit\update_url.txt`，并把里面的地址改成你的 Gitee `latest.json` 地址。
4. 把 `dist\HRToolkit\*` 压缩成 zip，例如 `HRToolkit-0.2.0-win.zip`。
5. 计算 zip 的 SHA256：

```powershell
Get-FileHash .\HRToolkit-0.2.0-win.zip -Algorithm SHA256
```

6. 上传 zip 到 ScriptHub 或 Gitee 可公开访问的位置，例如：

```text
https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main/release/downloads/HRToolkit-0.2.0-win.zip
```

7. 按 `release/latest.json.example` 生成 `latest.json`，填好 `version`、`file_url`、`sha256`，上传到：

```text
https://gitee.com/optimistic-little-sunspot/hr-toolkit/raw/main/release/latest.json
```

### 只用 Gitee 发布

可以不用 ScriptHub，直接用 Gitee。推荐用一键发布脚本，不要手写 `version`、`file_url` 和 `sha256`。

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
- 把 `dist\HRToolkit\*` 压缩到 `release/downloads/HRToolkit-版本号-win.zip`
- 计算 zip 的 SHA256
- 更新 `release/latest.json` 里的 `version`、`file_url`、`sha256`

4. 提交并推送：

```bash
git add hr_toolkit/__init__.py release/latest.json release/downloads/
git commit -m "发布 HR工具箱 0.2.0"
git push gitee main
```

只推源码不会触发客户端更新。客户端只看 `latest.json` 里的 `version` 是否大于当前程序版本，并按 `file_url` 下载 zip。
