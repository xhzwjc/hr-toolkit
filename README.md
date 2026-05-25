# HR Toolkit

人事 Excel 自动化工具箱。当前先落地一个可维护样板工具：**需求4：工资表按入职公司拆分**。

## 已实现工具

### 需求4-工资表按入职公司拆分

输入一个包含 `汇总表`、`明细表` 的工资表，按明细表中的 `入职公司` 字段拆分为多个 Excel 工作簿。

输出内容：

- 每个入职公司一个独立 `.xlsx`
- 保留原工资表的主要样式、表头、公式结构
- 明细表只保留对应公司的员工行
- 汇总表按项目重新生成汇总行
- 输出 `_salary_split_manifest.json`，便于系统集成和测试核对

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

## 后续扩展约定

每个需求独立成一个工具模块：

- `hr_toolkit/tools/salary_split.py`：需求4，工资表拆分
- `hr_toolkit/tools/salary_merge.py`：需求5，工资表合并
- `hr_toolkit/tools/archive_import.py`：需求7，档案移交表入库
- `hr_toolkit/tools/social_security.py`：需求1，社保明细/汇总

CLI 只是入口，核心函数可以直接被 ScriptHub 或 Web 后端调用。

