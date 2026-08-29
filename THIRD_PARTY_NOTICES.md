# Third-Party Notices

This project uses the following direct dependencies. Version locks are maintained in `pyproject.toml`, `apps/web/package.json`, `apps/web/pnpm-lock.yaml`, and the .NET project file.

| Component | Version | License | Source | Purpose |
|---|---:|---|---|---|
| Alembic | 1.19.1 | MIT | https://github.com/sqlalchemy/alembic | Database schema migrations |
| FastAPI | 0.141.1 | MIT | https://github.com/fastapi/fastapi | HTTP API |
| Pydantic | 2.13.4 | MIT | https://github.com/pydantic/pydantic | Rule and report models |
| OpenPyXL | 3.1.5 | MIT | https://foss.heptapod.net/openpyxl/openpyxl | Development-only renderer and test fixtures |
| RapidFuzz | 3.14.5 | MIT | https://github.com/rapidfuzz/RapidFuzz | Suggestion-only fuzzy header candidates |
| DataComPy | 1.0.4 | Apache-2.0 | https://github.com/capitalone/datacompy | Optional dataset-join cross-check adapter |
| Pandera | 0.32.1 | MIT | https://github.com/unionai-oss/pandera | Standard-data schema validation adapter |
| Polars | 1.43.2 | MIT | https://github.com/pola-rs/polars | Partitioned Parquet key joins for large comparisons |
| PyYAML | 6.0.3 | MIT | https://github.com/yaml/pyyaml | YAML rule loading |
| Uvicorn | 0.52.4 | BSD-3-Clause | https://github.com/encode/uvicorn | ASGI server |
| python-multipart | 0.0.32 | Apache-2.0 | https://github.com/Kludex/python-multipart | Multipart upload parsing |
| HTTPX | 0.28.1 | BSD-3-Clause | https://github.com/encode/httpx | API tests and managed HTTP adapter foundation |
| ijson | 3.5.1 | BSD-3-Clause | https://github.com/ICRAR/ijson | Bounded-memory JSON standard-data parsing |
| SQLAlchemy | 2.0.52 | MIT | https://github.com/sqlalchemy/sqlalchemy | PostgreSQL persistence layer |
| Psycopg | 3.3.4 | LGPL-3.0 with exceptions | https://github.com/psycopg/psycopg | PostgreSQL driver |
| Redis | 8.1.0 | MIT | https://github.com/redis/redis-py | Redis client |
| RQ | 2.11.0 | BSD-2-Clause | https://github.com/rq/rq | Background task queue |
| Boto3 | 1.43.82 | Apache-2.0 | https://github.com/boto/boto3 | S3/MinIO artifact storage |
| DocumentFormat.OpenXml | 3.5.1 | MIT | https://github.com/dotnet/Open-XML-SDK | Production Excel renderer |
| Vue | 3.5.20 | MIT | https://github.com/vuejs/core | Web user interface |
| Vite | 8.2.2 | MIT | https://github.com/vitejs/vite | Web build tooling |
| TypeScript | 5.9.2 | Apache-2.0 | https://github.com/microsoft/TypeScript | Web type checking |
| Hypothesis | 6.165.10 | MPL-2.0 | https://github.com/HypothesisWorks/hypothesis | Property-based tests |
| psutil | 7.2.2 | BSD-3-Clause | https://github.com/giampaolo/psutil | Performance-test memory measurement |

Transitive dependency licenses are captured by the CI-generated SBOM and license inventory. No AGPL source is copied. The renderer is an internal Open XML SDK implementation behind the `ExcelRenderer` adapter; the proposed `xlsx-review` fork was not copied because its integration was not needed to preserve the replaceable adapter boundary.
