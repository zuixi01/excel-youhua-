from __future__ import annotations

import os
from pathlib import Path

from .persistence import DatabaseRepository
from .service import AuditService
from .storage import S3ArtifactStore
from .security_config import validate_production_environment


def run() -> None:
    validate_production_environment()
    database_url = os.environ.get("DATABASE_URL")
    bucket = os.environ.get("S3_BUCKET")
    database = DatabaseRepository(database_url, create_schema=False) if database_url else None
    store = S3ArtifactStore(bucket, os.environ.get("S3_ENDPOINT_URL"), os.environ.get("AWS_REGION"), os.environ.get("S3_SERVER_SIDE_ENCRYPTION", "AES256")) if bucket else None
    root = Path(os.environ.get("EXCEL_AUDITOR_DATA", "var")).resolve()
    purged = AuditService(root, database=database, artifact_store=store).purge_expired()
    print(f"purged_jobs={len(purged)}")
