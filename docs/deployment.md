# 构建与部署

## 不可变构建

所有测试和镜像构建在本地或 CI 完成。发布工作流使用 Git SHA 生成：

- `excel-auditor-api:<git-sha>`（包含已编译的 .NET 自包含渲染器）；
- `excel-auditor-web:<git-sha>`。

生产配置 `deploy/docker-compose.prod.yaml` 强制要求显式 `IMAGE_TAG`，不接受 `latest`。
发布工作流先复用完整 CI，再本地构建、扫描、生成 SBOM 和执行容器冒烟；所有门禁通过后才登录仓库并推送镜像。

生产环境启动时执行 fail-fast 安全校验。至少要求：PostgreSQL URL 使用 `sslmode=require`、`verify-ca` 或 `verify-full`；Redis 使用 `rediss://`；对象存储启用 `AES256` 或 `aws:kms` 服务端加密，使用自定义 S3 endpoint 时必须为 HTTPS；鉴权必须开启且令牌长度至少 32 字符。多租户 API 可用 `EXCEL_AUDITOR_API_TOKENS_JSON` 把长随机令牌映射到独立的 `tenant_id` 和 `user_id`。秘密只通过部署平台注入，不写入仓库、Compose 文件、日志或报告。

受管 HTTP 标准源使用 `STANDARD_CONNECTIONS_FILE` 挂载只读连接注册表；注册表只保存 `auth_secret_ref`，不保存密钥。`STANDARD_SECRETS_DIR` 指向受保护的宿主机目录，每个引用对应一个同名、只读的小文件。上传型标准源可继续使用仓库提供的空注册表和空秘密目录。不得把实际连接密钥写入 `env.example`、连接 JSON 或镜像。

## 服务器发布前只读检查

执行任何拉取或启动前，必须检查并记录：

```bash
uptime
free -h
swapon --show
df -h
df -ih
docker system df
docker compose -f deploy/docker-compose.prod.yaml ps
```

同时检查最近系统日志是否有 OOM、磁盘 I/O、文件系统或内核错误。资源达到项目 `AGENTS.md` 阈值时停止发布，不自动清理、重启或远程构建。

## 发布与验证

记录当前镜像 digest 后，仅在服务器执行：

```bash
docker compose -f deploy/docker-compose.prod.yaml pull
docker compose -f deploy/docker-compose.prod.yaml up -d --remove-orphans
docker compose -f deploy/docker-compose.prod.yaml ps
```

随后验证 `/health/live`、`/health/ready`、Worker 心跳和一份无敏感数据的小型 smoke comparison。日志只读取有限行数，禁止输出环境变量和秘密信息。

发布记录必须包含 Git SHA、API/Web 镜像完整 digest、SBOM/漏洞扫描制品、部署前旧 digest、健康检查结果和操作者/时间。没有这些记录时不得把部署标记为完成。

## 回滚

将 `IMAGE_TAG` 恢复为发布前记录的 Git SHA 或 digest，再执行：

```bash
docker compose -f deploy/docker-compose.prod.yaml pull
docker compose -f deploy/docker-compose.prod.yaml up -d --remove-orphans
```

不得在生产服务器运行 `docker build`、语言构建命令或未经确认的 prune/重启/清理操作。

若健康检查失败，停止继续发布并使用记录的旧 digest 回滚；不得临时改用 `latest`、现场修改源码或在服务器重建镜像。
