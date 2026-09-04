from pathlib import Path

from excel_auditor.queueing import RedisJobQueue
from excel_auditor.rules import load_rules


def test_product_jobs_and_revisions_use_versioned_rq_payloads(monkeypatch):
    import excel_auditor.queueing as queueing

    calls = []

    class FakeQueue:
        def __init__(self, name, connection, default_timeout):
            calls.append(("init", name, connection, default_timeout))

        def enqueue(self, *args, **kwargs):
            calls.append(("enqueue", args, kwargs))

    connection = object()
    monkeypatch.setattr(queueing.Redis, "from_url", staticmethod(lambda url, decode_responses: connection))
    monkeypatch.setattr(queueing, "Queue", FakeQueue)
    rules = load_rules(Path("configs/examples/product-normalization.yaml"))
    queue = RedisJobQueue("redis://example/0")

    queue.enqueue_product(Path("runtime"), "job_1", Path("input.xlsx"), rules, "tenant-a", "operator")
    queue.enqueue_product_revision(Path("runtime"), "job_1", rules, "tenant-a", "operator", 3)

    initial = calls[1]
    assert initial[1][0] == "excel_auditor.worker.run_product_job"
    assert initial[2]["job_id"] == "rq_product_job_1"
    assert initial[2]["job_timeout"] == rules.workbook.processing_timeout_seconds
    revision = calls[2]
    assert revision[1][0] == "excel_auditor.worker.run_product_revision"
    assert revision[2]["job_id"] == "rq_product_revision_job_1_3"
