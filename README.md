# Excel Standard Auditor

配置驱动的 Excel 标准数据核验与受控修复平台。实现采用 Python/FastAPI 比对核心、PostgreSQL/Redis/RQ/MinIO 基础设施、Vue 3 前端，以及 .NET 8 Open XML SDK 高保真渲染器。

## 已实现能力

- 严格、版本化的 YAML/JSON 规则；所有嵌套配置拒绝未知字段，发布版本不可变。
- XLSX/XLSM 安全预检、表头自动定位、工作表/表头别名、模糊候选人工审核。
- 单主键/联合主键匹配，空键、重复键、缺失/多余记录隔离。
- 字符串、整数、Decimal、日期时间、布尔、枚举、集合、JSON 的类型化比较，以及字段和跨字段质量规则。
- 上传 JSON/CSV、请求内 JSON、受管 HTTP 分页数据源；固化、回读并校验 SHA-256 标准快照。
- 标色、批注、插列、填值、追加记录和内置报告；维护受支持的公式、Table、筛选、验证、合并区域、定义名称和宏部件，对复杂结构明确失败或转人工。
- JSON/HTML/Excel 报告、不可预测差异 ID、可复现语义投影、逐项修复结果与审计记录。
- Bearer 鉴权、租户级规则/草稿/任务/下载隔离、并发配额、软删除和延迟清理。
- PostgreSQL、Redis/RQ、MinIO/S3 适配器、Alembic 迁移、健康检查和指标。
- 19 个固定 Golden 工作簿（含 10,000×20 大文件与固定标准快照）、属性/安全/可复现/性能测试，以及真实基础设施和 LibreOffice CI 作业。
- 锁定的 Python、Node 和 NuGet 依赖；许可证清单、SBOM、漏洞扫描和“扫描通过后才推送”的 SHA 镜像发布流程。

## 本地验证

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python -m pip install uv==0.12.5
.\.venv\Scripts\uv sync --frozen --extra test --extra security

dotnet restore services\excel_renderer\ExcelRenderer.csproj --locked-mode
dotnet publish services\excel_renderer\ExcelRenderer.csproj -c Release -r win-x64 --self-contained true --no-restore -o .renderer
$env:EXCEL_RENDERER_COMMAND = (Resolve-Path .renderer\ExcelRenderer.exe)
.\.venv\Scripts\python -m pytest -q

pnpm --dir apps/web install --frozen-lockfile
pnpm --dir apps/web build
```

示例规则在 `configs/examples/employee-roster.yaml`。开发环境可使用 `deploy/docker-compose.yaml`；生产发布只能使用 `deploy/docker-compose.prod.yaml` 和不可变的 Git SHA/digest 镜像，服务器不执行构建。

实现状态见 `docs/implementation-status.md`，逐项验收证据见 `docs/requirements-evidence-matrix.md`，构建与回滚规则见 `docs/deployment.md`。
