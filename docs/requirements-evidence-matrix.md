# 开发文档逐项验收证据矩阵

基线：`excel-standard-audit-development.md` v1.0.0，第 21 节 DoD 1–20 与第 22 节场景 A–F。状态含义：

- **本地已证明**：当前工作区存在可重复的实现和自动化断言。
- **部分证明**：核心实现存在，但缺少文档要求的某类环境或发布证据。
- **待外部验收**：必须依赖远程 CI、真实客户端、镜像仓库、服务器或业务材料，当前环境不能替代。

## Definition of Done

| # | 要求 | 状态 | 当前证据 | 仍需证据 |
|---:|---|---|---|---|
| 1 | API 上传 Excel/标准数据或受管接口获取 | 本地已证明 | `src/excel_auditor/api.py`、`standard_sources.py`；生产只读连接/秘密挂载；`tests/test_api.py`、`test_standard_sources.py` | 真实业务接口契约验收 |
| 2 | 标准数据固化为带 SHA-256 的审计快照 | 本地已证明 | `snapshots.py` 回读校验、对象存储路径；`test_storage.py`、基础设施集成测试 | 生产对象存储保留策略验收 |
| 3 | 按规则识别工作表和表头行 | 本地已证明 | `workbook.py`、`engine.py`；别名/自动定位/缺失表头测试 | 业务模板样例 |
| 4 | 缺失、多余、别名、重复、歧义表头 | 本地已证明 | 15 个核心 Golden、模糊表头人工审核端到端测试 | 业务异常样例 |
| 5 | 缺失列确定插入并标绿，多余列保留标红 | 本地已证明 | 首/中/末 Golden 的完整表头和填充断言；Renderer before/after 合约 | Excel 客户端视觉抽验 |
| 6 | 单主键和联合主键准确匹配 | 本地已证明 | `engine.py`、属性测试；`typed_compound.xlsx` 固定 Golden | 人工标注业务基准集 |
| 7 | 多余/缺失记录、空/重复主键 | 本地已证明 | 核心 Golden、`test_end_to_end.py` | 业务异常样例 |
| 8 | 类型化比较字符串、数值、日期时间、布尔、枚举、集合 | 本地已证明 | `normalization.py`、单元/属性测试；高级类型 Golden | 人工标注业务基准集 |
| 9 | 字段及跨字段数据质量规则 | 本地已证明 | `pandera_adapter.py`、注册校验器；`test_validation.py` | 业务规则清单 |
| 10 | 默认不覆盖非空数据、不删除列/记录 | 本地已证明 | 默认动作模型、修复授权检查、端到端测试 | 业务方修复策略签署 |
| 11 | 输出标色 Excel、JSON、HTML | 本地已证明 | `service.py`、`reporting.py`、Renderer；端到端与 Golden 渲染断言 | 客户端下载抽验 |
| 12 | Excel 与 LibreOffice 可打开，结构回归通过 | 部分证明 | Open XML Validator、回读、结构化 Golden；提交 `08afd68` 的 [主 CI/LibreOffice 兼容作业](https://github.com/zuixi01/excel-youhua-/actions/runs/33251264106) 及 JUnit 制品已通过 | Microsoft Excel 桌面验收记录 |
| 13 | 差异追踪至任务、规则、快照、工作表、单元格、业务主键 | 本地已证明 | `Difference`/`AuditReport`、数据库索引、报告投影测试 | 生产审计抽样 |
| 14 | 修复含规则 ID、原值、标准值和审计 | 本地已证明 | `service.py` 修复审计、差异模型、数据库测试 | 生产审计抽样 |
| 15 | Golden、集成、安全、性能测试全部通过 | 部分证明 | 19 个固定 Golden；本地生产 Renderer 完整回归通过；提交 `0534ff2` 的 [主 CI](https://github.com/zuixi01/excel-youhua-/actions/runs/33252789525) 七项与 [性能基线 v4](https://github.com/zuixi01/excel-youhua-/actions/runs/33252958226) 八项全部绿色并保存制品；500k 标准/400,001 差异直接比较及 100 页受管 HTTP 服务全链路均通过；小型人工标注 precision/recall | 业务标注基准 |
| 16 | 许可证、NOTICE、锁定和 SBOM 完整 | 部分证明 | 三类锁文件、`THIRD_PARTY_NOTICES.md`；提交 `08afd68` 的依赖安全作业已通过并保存 `dependency-governance`（SBOM、许可证、漏洞扫描）制品 | 对正式发布 SHA 保存并归档同类制品 |
| 17 | CI 构建版本化镜像，服务器只拉取启动 | 待外部验收 | Git remote 已配置；`ci.yml`/`release.yml` 先测试扫描后推送；生产 Compose 强制 tag | 成功的标签发布运行、GHCR 中的 SHA/digest 镜像 |
| 18 | 精确回滚版本与命令 | 待外部验收 | `docs/deployment.md`、release manifest 生成逻辑 | 发布前后真实 digest 和一次回滚演练记录 |
| 19 | 日志/报告不泄密或未脱敏敏感数据 | 本地已证明 | 掩码、manifest 脱敏、安全错误；Web Bearer 令牌仅存会话且下载走鉴权请求；`test_privacy.py`、`test_startup_security.py` | 生产日志抽检 |
| 20 | 不支持/不安全场景明确失败或人工审核 | 本地已证明（已知边界） | 严格 manifest、未知操作拒绝；复杂公式/图表/透视/外链插列拒绝；已支持结构只报告而不误拦截；真正不支持的包结构即使配置 `allow/report` 也不可绕过人工审核；模糊/合并表头测试 | 业务真实复杂工作簿扩充 Golden，持续维护边界清单 |

## 关键场景 A–F

| 场景 | 状态 | 证据 |
|---|---|---|
| A 表头缺失和多余 | 本地已证明 | `missing_and_extra_header.xlsx`；完整差异与渲染表头/颜色断言 |
| B 记录级差异 | 本地已证明 | `extra_and_missing_record.xlsx`；默认不追加与显式 `append_and_mark_green` 端到端测试 |
| C 数值容差 | 本地已证明 | `test_end_to_end.py` 的 `10000.005` 对 `10000.00`、属性边界测试 |
| D 主键重复 | 本地已证明 | `duplicate_primary_key.xlsx`；两行紫色且不参与一对一匹配 |
| E 模糊表头 | 本地已证明 | `test_fuzzy_header_suggestion_requires_manual_review` |
| F 结果可复现 | 本地已证明 | 双运行报告/manifest 语义投影测试；随机任务 ID/时间明确排除，不宣称二进制逐字节相同 |

## 当前阻止“生产完成”的外部输入

1. 创建正式版本标签，让 release 门禁在已有 CI 和性能参考基线上构建、扫描并推送版本化镜像。
2. 保存发布镜像 digest、镜像 SBOM、漏洞报告和 release manifest，并验证 GHCR 拉取。
3. 在安装 Microsoft Excel 的验收机打开固定 Golden 及真实 `.xlsm` 样例，核验格式、关系、宏、签名和控件。
4. 提供开发文档第 23 节列出的业务模板、脱敏异常文件、标准接口、主键/字段/敏感规则与容量目标。
5. 提供目标服务器后，按 `AGENTS.md` 完成只读资源检查、指定 digest 部署、健康检查和回滚演练；服务器不承担构建。
