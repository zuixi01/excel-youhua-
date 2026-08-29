# Excel Renderer

.NET 8 / Open XML SDK 渲染器。命令行只接受严格 JSON manifest，不执行表达式、宏或任意脚本：

```text
ExcelRenderer --input input.xlsx --output result.xlsx --manifest render-manifest.json
ExcelRenderer --input input.xlsx --output result.xlsx --manifest render-manifest.json --dry-run
```

渲染器先校验输入 SHA-256，在同目录临时文件中处理，完成 Open XML 和关键结构验证后才原子替换输出；dry-run 不触碰已有输出。结果包含原始文件 SHA、排除自引用元数据工作表后的内容哈希，以及带操作索引、差异 ID、状态和错误信息的逐操作结果。

当前支持：单元格/整行标色、批注、before/after 插列、类型化写值、追加行、报告工作表、veryHidden 元数据工作表、样式去重、数字格式/验证/公式模板、Table/筛选/验证/命名区域/合并区域/内部超链接及简单 A1 公式引用维护。

安全边界：插列遇到图表或其他 drawing、透视表、外部链接、复杂/共享/数组公式时返回 `UNSUPPORTED_FEATURE`；含旧式控件或宏形状的 VML 不会被重写。带业务时区的 datetime 写值在 manifest 中携带 IANA `timezone`，Renderer 会跨 Windows/Linux 解析时区并拒绝不存在、重复或 offset 不匹配的墙上时间，避免 Excel 日期序列丢失 DST 语义。`.xlsm` 只允许保留 VBA 部件，不修改或执行宏。新增 Open XML 特性必须先增加固定 Golden 或 Renderer 合约测试。

发布构建同时锁定 `win-x64`、`linux-x64` 和 `linux-musl-x64` RuntimeIdentifier；Alpine API 镜像使用 `linux-musl-x64` 自包含二进制，并安装 IANA/Windows 时区映射所需的 ICU 运行库。
