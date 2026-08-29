from __future__ import annotations

from pathlib import Path
from typing import Any

from redis import Redis
from rq import Queue

from .models import RuleSet


class RedisJobQueue:
    def __init__(self, redis_url: str, queue_name: str = "excel-auditor") -> None:
        self.connection = Redis.from_url(redis_url, decode_responses=False)
        self.queue = Queue(queue_name, connection=self.connection, default_timeout=900)

    def enqueue(self, data_root: Path, job_id: str, excel_path: Path, standard_path: Path | None, rules: RuleSet, parameters: dict[str, Any]) -> None:
        self.queue.enqueue(
            "excel_auditor.worker.run_job",
            str(data_root),
            job_id,
            str(excel_path),
            str(standard_path) if standard_path else None,
            rules.model_dump(mode="json"),
            parameters,
            job_id="rq_" + job_id,
            job_timeout=rules.workbook.processing_timeout_seconds,
            result_ttl=3600,
            failure_ttl=86400,
        )

    def ping(self) -> None:
        self.connection.ping()
