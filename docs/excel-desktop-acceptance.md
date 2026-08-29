# Microsoft Excel 桌面验收

本流程用于补齐自动化 Open XML/LibreOffice 测试不能替代的 Microsoft Excel 客户端证据。只有自动检查和独立人工检查都通过，才可把开发文档 DoD 12 的 Excel 部分标记为完成。

## 输入与安全边界

- 使用安装了受支持 Microsoft Excel 桌面版的隔离 Windows 验收机。
- 至少选择一个固定渲染 Golden `.xlsx` 和一个包含真实业务宏、签名及实际 ActiveX/VML 控件的渲染后 `.xlsm`。
- 输入必须是待发布 Git SHA 产生的输出文件；记录其来源任务和下载制品。
- 自动脚本设置 `msoAutomationSecurityForceDisable`、禁用事件和链接更新，不执行工作簿宏，也不修改原文件。
- 不把业务工作簿或其中的数据提交到 Git；证据只记录文件名、SHA-256、工作表名、Excel 版本和检查结果。

## 自动打开与另存回归

在仓库根目录使用 Windows PowerShell 执行，证据文件必须是尚不存在的新路径：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/Invoke-ExcelDesktopAcceptance.ps1 `
  -WorkbookPath C:\acceptance\golden-output.xlsx,C:\acceptance\business-output.xlsm `
  -EvidencePath C:\acceptance\excel-automated-evidence.json
```

脚本会完成以下检查：

1. 以只读、强制禁用宏方式通过 Excel COM 打开每个文件；
2. 获取工作表名称并 `SaveCopyAs` 到独立临时目录；
3. 再次用 Excel 打开另存副本，并核对工作表集合；
4. 比较 VBA 工程、VBA 签名、ActiveX、控件属性、嵌入对象和 VML 关键包部件的 SHA-256；
5. 写入不可覆盖的结构化 JSON 证据，失败时保留失败结果并返回非零退出码。

## 独立人工检查

由非本次实现操作者在 Excel 中交互式打开同一批文件，确认没有修复提示或受保护视图之外的异常，并逐项检查：

- 工作表名称、顺序、冻结窗格、隐藏状态和布局正确；
- 插入列、标色、批注和“核验报告”内容正确；
- 公式、Table、筛选、数据验证和命名区域行为正确；
- 真实宏、数字签名、ActiveX/VML 控件符合业务预期；
- 不存在静默丢失、修复日志或意外外部链接更新。

复制 `docs/excel-acceptance-review.example.json` 到验收目录，填写自动证据文件的 SHA-256、审核人、UTC 时间和结论。不得把示例中的 `pending`/`false` 直接改成通过而不执行人工检查。

## 证据校验与归档

```powershell
uv run --frozen python -m excel_auditor.excel_acceptance `
  C:\acceptance\excel-automated-evidence.json `
  C:\acceptance\excel-human-review.json
```

校验器要求自动证据同时包含 `.xlsx`/`.xlsm`、所有自动检查为真，人工检查逐项批准，并使用 SHA-256 把人工签核绑定到不可变的自动证据。将两个 JSON、校验器输出、发布 Git SHA、Excel 版本和受控工作簿来源一起归档到组织的审计存储；业务工作簿仍按敏感数据策略保存，不上传公共仓库。
