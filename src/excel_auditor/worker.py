from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from .models import RuleSet
from .persistence import DatabaseRepository
from .service import AuditService
from .standard_sources import ConnectionRegistry, ManagedHttpSource
from .storage import S3ArtifactStore


def run() -> None:
    from redis import Redis
    from rq import Queue, Worker

    from .security_config import validate_production_environment

    validate_production_environment()
    redis_url = os.environ.get("REDIS_URL", "redis://127.0.0.1:6379/0")
    connection = Redis.from_url(redis_url)
    queue = Queue("excel-auditor", connection=connection)
    Worker([queue], connection=connection).work()


def run_job(data_root: str, job_id: str, excel_path: str, standard_path: str | None, rule_payload: dict[str, Any], parameters: dict[str, Any]) -> None:
    connections_path = os.environ.get("EXCEL_AUDITOR_CONNECTIONS")
    managed = ManagedHttpSource(ConnectionRegistry(Path(connections_path))) if connections_path else None
    database_url = os.environ.get("DATABASE_URL")
    database = DatabaseRepository(database_url, create_schema=False) if database_url else None
    s3_bucket = os.environ.get("S3_BUCKET")
    artifact_store = S3ArtifactStore(s3_bucket, os.environ.get("S3_ENDPOINT_URL"), os.environ.get("AWS_REGION"), os.environ.get("S3_SERVER_SIDE_ENCRYPTION", "AES256")) if s3_bucket else None
    service = AuditService(Path(data_root), managed_http=managed, database=database, artifact_store=artifact_store)
    service.run(job_id, Path(excel_path), Path(standard_path) if standard_path else None, RuleSet.model_validate(rule_payload), parameters)


def run_product_job(
    data_root: str,
    job_id: str,
    excel_path: str,
    rule_payload: dict[str, Any],
    tenant_id: str,
    actor_id: str,
) -> None:
    from .product_workflow import ManagedHttpCatalogAdapter
    from .product_workflow.service import ProductWorkflowService

    service, managed, database = _build_runtime(Path(data_root))
    rules = RuleSet.model_validate(rule_payload)
    if rules.product_workflow is None or managed is None:
        raise RuntimeError("product workflow requires a managed platform catalog connection")
    adapter = ManagedHttpCatalogAdapter(managed, rules.product_workflow)
    ProductWorkflowService(service, database).run(
        job_id,
        Path(excel_path),
        rules,
        adapter,
        tenant_id=tenant_id,
        actor_id=actor_id,
    )


def run_product_revision(
    data_root: str,
    job_id: str,
    rule_payload: dict[str, Any],
    tenant_id: str,
    actor_id: str,
) -> None:
    from .product_workflow.service import ProductWorkflowService

    service, _managed, database = _build_runtime(Path(data_root))
    ProductWorkflowService(service, database).rerun_after_reviews_safe(
        job_id,
        RuleSet.model_validate(rule_payload),
        tenant_id=tenant_id,
        actor_id=actor_id,
    )


def _build_runtime(data_root: Path) -> tuple[AuditService, ManagedHttpSource | None, DatabaseRepository | None]:
    connections_path = os.environ.get("EXCEL_AUDITOR_CONNECTIONS")
    managed = ManagedHttpSource(ConnectionRegistry(Path(connections_path))) if connections_path else None
    database_url = os.environ.get("DATABASE_URL")
    database = DatabaseRepository(database_url, create_schema=False) if database_url else None
    s3_bucket = os.environ.get("S3_BUCKET")
    artifact_store = S3ArtifactStore(
        s3_bucket,
        os.environ.get("S3_ENDPOINT_URL"),
        os.environ.get("AWS_REGION"),
        os.environ.get("S3_SERVER_SIDE_ENCRYPTION", "AES256"),
    ) if s3_bucket else None
    return (
        AuditService(data_root, managed_http=managed, database=database, artifact_store=artifact_store),
        managed,
        database,
    )
