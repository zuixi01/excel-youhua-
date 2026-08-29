# 开发文档落实与验收状态

更新日期：2026-08-29

本页只陈述当前工作区可以复核的事实。完整的逐项状态见 `docs/requirements-evidence-matrix.md`；“代码存在”不等于已经取得生产验收证据。

## 当前实现

已完成的本地实现包括：严格且版本化的规则模型、JSON/CSV/受管 HTTP 标准数据接入、不可变标准快照、工作簿安全预检、表头映射、单/联合主键匹配、类型化比较、字段及跨字段校验、受控修复、JSON/HTML/Excel 报告、FastAPI/Web、PostgreSQL/Redis/RQ/MinIO 适配器、租户隔离、审计与保留策略。Web 的任务、规则与下载请求统一支持 Bearer 鉴权，令牌只保存在当前浏览器会话，文件通过鉴权请求下载。生产 Compose 为受管 HTTP 连接注册表和秘密目录提供只读挂载；连接配置严格拒绝未知字段，密钥不进入规则或镜像。

.NET Renderer 当前覆盖标色、批注、插列、类型化写值、追加记录、报告/隐藏元数据工作表、样式去重、Table/AutoFilter/DataValidation/DefinedName/合并区域和简单 A1 公式引用维护。内部超链接、局部命名区域、筛选列索引和排序引用均有插列维护；图表/绘图、透视表、外部链接以及复杂/共享/数组公式在插列时明确返回 `UNSUPPORTED_FEATURE`，不会静默修改。Renderer 还提供结构化 `--version` 自检，API 就绪探针会实际执行并校验其身份与语义版本。

固定 Golden 语料目前为 19 个工作簿：15 个核心表头/记录场景，外加联合主键与完整类型、结构化多工作表、合并表头人工审核，以及 10,000×20 大文件报告模式。大文件输入和 10,000 条标准 JSONL 快照均锁定 SHA，并强制工作簿与标准记录走磁盘后备序列及 Polars 分区连接。Golden 测试校验输入 SHA、完整差异计数/语义投影以及关键渲染结构；Renderer 合约另覆盖公式引用、Table、筛选、验证、命名区域、内部超链接、宏部件字节保持和不安全插列拒绝。

## 当前验证结果

- 当前完整本地生产 Renderer 回归：`159 passed, 8 skipped`（共收集 167 项）；条件跳过的性能、真实基础设施和 LibreOffice 已由专用远程作业承接。
- 当前远端提交 `e66788de2a81b936610d7dd7157cf6badfa8c6fe` 的 [主 CI](https://github.com/zuixi01/excel-youhua-/actions/runs/33253936071) 全部绿色：Python 3.12/3.14、Web、Windows Renderer、LibreOffice、真实 PostgreSQL/Redis/RQ/MinIO 和依赖安全共七项均通过；JUnit、Coverage、LibreOffice、Renderer、基础设施及依赖治理制品均已保存。
- 同一提交的 [性能基线 v4 最终复核运行](https://github.com/zuixi01/excel-youhua-/actions/runs/33254109004) 九项全部绿色：五个工作簿负载场景、50,000 条分页标准源、500,000 条直接比较、500,000 条受管 HTTP 服务全链路，以及 500,000 条上传 JSON 的流式解析、校验与快照均通过相对回归门禁并保存指标及 JUnit 制品。
- 条件跳过项分别由上述性能、真实基础设施和 LibreOffice 专用远程作业承接；远程通过不改变“本机未执行”的事实。
- 本地 Windows Renderer locked restore/build/publish 通过；新增 Renderer 合约测试通过。
- 仓库记录的 v4 参考性能包括：10k×50 为 14.416 秒/50.48 MiB；100k×100 为 41.054 秒/254.23 MiB；100k×200、50% 差异为 53.641 秒/287.38 MiB；5 工作表 50k×50 为 11.045 秒/18.2 MiB；50,000 条分页标准源为 0.161 秒。500,000 条上传 JSON 的流式解析、校验和快照为 15.058 秒/132.49 MiB；500,000 条直接比较为 78.03 秒/268.13 MiB；受管 HTTP 服务全链路为 155.871 秒/303.38 MiB，并生成约 549 MiB 的 JSON、JSONL、HTML 工件。详见 `docs/performance-baselines.json`。
- 性能工作流会下载上一成功主分支同版本、同场景指标；工作簿和最大规模场景以固定纯 Python 校准归一化进程 CPU，再与峰值 RSS 一起执行相对回归门禁，墙钟时间仍受硬限制。分页短场景保留进程 CPU 与绝对余量。正式发布缺少兼容参考或发生实质回归都会失败。
- Python、Node、NuGet 锁文件以及安全/许可证/SBOM 工作流已配置；本地已对锁定的 Python、Node、NuGet 依赖执行漏洞检查，新增 ijson 依赖为 BSD-3-Clause/ISC 且已写入 NOTICE。远程 CI 已保存 `dependency-governance` 制品；GitHub Actions 已升级到官方最新发布版并固定不可变提交 SHA，运行不再产生 Node 20 运行时弃用告警。

## 尚未取得的完成证据

- Git remote、主 CI、性能、基础设施、LibreOffice 和安全制品证据已经取得；尚未创建版本发布标签，因此还没有由 release 门禁构建并推送的不可变镜像 digest、镜像 SBOM 和 release manifest。
- 当前主机未安装 Microsoft Excel 或 LibreOffice；LibreOffice 已由远程专用作业出具绿色 JUnit 证据，Microsoft Excel 桌面打开验收仍不能伪造。
- 当前 Docker `desktop-linux` Engine 可响应，但只读前置检查显示可用内存约 2.9 GiB（低于总内存 25% 的安全门槛）、C 盘可用比例约 19%（低于 20%），且已有 11 个其他项目容器运行。按 `AGENTS.md` 已熔断，不启动额外基础设施或构建镜像；最新源码对应镜像的最终重建、容器冒烟和 High/Critical 扫描仍无本机证据，历史镜像结果不能替代。
- `.xlsm` 现有测试使用固定 VBA 部件验证字节保持，不等同于真实业务宏、数字签名、ActiveX/VML 控件在 Excel 客户端中的验收。
- 大文件 Inspector、规范标准记录和大差异序列均可溢写到临时文件；上传 JSON/CSV 以 1 MiB 分块落盘并逐记录解析，达到可配置阈值后迁移到磁盘后备序列。标准数据 Pandera 校验按块执行、跨块检查唯一性，快照按源顺序流式写入且不再全量排序复制；受管 HTTP 分页记录同样会按阈值溢写，偏移索引使用紧凑的 64 位数组。大标准主键/记录载荷使用 SQLite 磁盘映射和 Polars 分区 Parquet 连接，数据库差异索引按千条批量写入，完整 JSON、JSONL 与 HTML 工件可在只报告模式生成。500,000 条上传 JSON 的解析、校验和快照，500,000 条标准记录与 400,001 条差异的直接比较，以及 100 页受管 HTTP 获取、校验、快照、服务比较和完整工件生成均已由远端负载证明。表单内联 `standard_json` 仍由框架作为字符串接收，但受默认 64 MiB 上传限制约束；真实外部网络和生产对象存储延迟不属于确定性 MockTransport 性能基线。
- precision/recall 已覆盖人工标注的表头、记录、字段差异和修复小型基准，但仍需业务方提供具有代表性的人工标注基准集。
- 尚未连接部署服务器，因此没有部署前资源检查、实际部署 digest、健康检查和回滚演练记录。
- 开发文档第 23 节要求的业务标准模板、2–3 份脱敏异常文件、真实接口契约、主键/字段/敏感数据规则和容量目标仍需业务方提供，才能完成业务验收而非通用样例验收。

因此，当前状态是“通用核心闭环已实现并有较强本地自动化证据”，不是“生产验收全部完成”。
