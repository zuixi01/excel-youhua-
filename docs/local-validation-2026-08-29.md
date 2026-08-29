# 本地验证记录（2026-08-29）

工作区：`E:\Excel优化`

> 历史说明：本文件保存 2026-08-29 较早阶段的本地验证快照，其中“尚无 Git remote/远程证据”等陈述已被后续工作取代。当前权威状态与最新 CI/性能链接见 `docs/implementation-status.md` 和 `docs/requirements-evidence-matrix.md`，不得用本快照否定后续证据。
平台：Windows x64，CPython 3.12.13，.NET SDK 8.0.424，pnpm 9.15.9

## 已执行并通过

| 门禁 | 结果 |
|---|---|
| Python/端到端/Windows Renderer 完整回归 | `141 passed, 5 skipped`，44.89 秒 |
| Python 覆盖率 | 81%；JUnit/coverage XML 生成成功 |
| 固定 Golden | 19 个二进制输入；新增 10,000×20 大文件及固定标准快照；高级 Golden 可确定性再生成并保持 SHA |
| Windows Renderer | locked restore、Release build、自包含 publish 成功 |
| Linux musl Renderer | `linux-musl-x64` 自包含 publish 成功 |
| Renderer 合约 | 8 通过；LibreOffice 条件测试 1 跳过；包含结构化版本自检及 API 就绪探针执行校验 |
| Python 锁与漏洞 | `uv lock --check`；对 `uv.lock` 导出的完整固定版本集合执行 `pip-audit --strict --no-deps`，无已知漏洞 |
| Python 许可证 | 86 个环境包完成许可证枚举，未发现 GPL/AGPL 依赖 |
| Web | pnpm 9.15.9 frozen install、Vue/TypeScript production build 成功；pnpm 无 High 漏洞 |
| NuGet | locked restore；direct/transitive vulnerability 列表为空 |
| 配置 | GitHub Actions/Compose 标准 YAML 解析通过；生产 Compose `config --quiet` 通过 |
| 性能回归门禁 | 相对阈值判定器 4 个单元测试通过；release 强制要求上一成功 main 同场景制品；工作簿场景使用 v4 CPU 校准归一化 |
| 性能记录 | 10k×50：12.297 秒/52.13 MiB；100k×100：37.783 秒/252.79 MiB；100k×200、50% 差异：51.325 秒/422.52 MiB；5 工作表 50k×50：11.232 秒/30.96 MiB；分页源 50,000 条：0.119 秒 |

完整测试命令使用 `services/excel_renderer/bin/Release/net8.0/win-x64/publish/ExcelRenderer.exe` 作为 `EXCEL_RENDERER_COMMAND`。条件跳过项为四个性能基线、两个真实基础设施测试和 LibreOffice 兼容性测试；性能用例已另行显式执行并记录，并由专用远程作业对当前提交复验。

## 本轮修复后的关键安全契约

- 插列同步维护内部超链接 location、局部 DefinedName、AutoFilter FilterColumn 与排序引用。
- 插列遇到 drawing/chart、pivot、external link、复杂/共享/数组公式时返回 `UNSUPPORTED_FEATURE`，输出文件不落地。
- 工作簿检查将可维护的 Table、筛选、冻结、验证、命名区域和普通公式保留为结构警告；真正不支持的包结构、复杂公式、受保护表和合并表头不可被 `allow/report` 静默绕过。
- Web 统一为 API 请求附加会话级 Bearer 令牌，Excel/JSON/JSONL/HTML 下载不再使用无法携带鉴权头的裸链接。
- 生产 Compose 将受管 HTTP 连接注册表和秘密目录只读挂载给 API/Worker；连接配置拒绝未知字段和路径型秘密引用，秘密文件加载有目录边界、大小和换行检查。
- 标准记录达到阈值后在规范化阶段直接溢写；Pandera 分块校验并跨块核验唯一性，快照写入不再创建全量排序副本。
- 重复 Excel 主键即使开启 `append_and_mark_green` 也禁止自动追加同键标准记录。
- release manifest 直接读取 registry manifest digest，不再对命令输出二次计算伪 digest。
- CI 中 matrix Python 表达式改为合法 YAML，所有工作流可被标准解析器加载。

## 未执行及原因

1. Docker Desktop 的 `com.docker.service` 显示 `Stopped`，但 `desktop-linux` Engine 当前可响应。只读前置检查显示可用内存约 2.9 GiB（低于总内存 25% 的安全门槛）、C 盘可用比例约 19%（低于 20%），并有 11 个其他项目容器运行。按照安全规则已熔断，未启动本项目基础设施、构建镜像、停止其他容器或清理缓存；最新源码镜像的重建、容器冒烟和 High/Critical 扫描尚无最终证据。
2. 当前主机没有 Microsoft Excel 或 LibreOffice，不能执行客户端打开验收。
3. 仓库没有 Git remote，不能生成 GitHub Actions 远程运行、GHCR digest、SBOM/release artifact 证据。
4. 没有目标服务器，未执行部署前资源检查、健康检查或回滚演练。

以上未执行项保持为发布阻断项，详见 `docs/requirements-evidence-matrix.md`。
