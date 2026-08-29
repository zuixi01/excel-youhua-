from pathlib import Path


def test_startup_runs_security_migrations_then_server(tmp_path, monkeypatch):
    import excel_auditor.startup as startup

    calls = []
    config = tmp_path / "alembic.ini"
    config.write_text("[alembic]", encoding="utf-8")
    monkeypatch.setenv("EXCEL_AUDITOR_ALEMBIC_CONFIG", str(config))
    monkeypatch.setenv("DATABASE_URL", "sqlite:///ignored-by-mock")
    monkeypatch.setenv("PORT", "8123")
    monkeypatch.setattr(startup, "validate_production_environment", lambda: calls.append("security"))
    monkeypatch.setattr(startup, "Config", lambda value: (calls.append(("config", value)) or "CONFIG"))
    monkeypatch.setattr(startup.command, "upgrade", lambda value, revision: calls.append(("upgrade", value, revision)))
    monkeypatch.setattr(startup.uvicorn, "run", lambda app, **kwargs: calls.append(("server", app, kwargs)))

    startup.run()

    assert calls == [
        "security",
        ("config", str(config)),
        ("upgrade", "CONFIG", "head"),
        ("server", "excel_auditor.api:app", {"host": "0.0.0.0", "port": 8123}),
    ]


def test_worker_uses_configured_redis_queue_and_security(monkeypatch):
    import redis
    import rq
    from excel_auditor import security_config, worker

    calls = []
    connection = object()
    queue = object()
    monkeypatch.setenv("REDIS_URL", "rediss://redis.example:6379/7")
    monkeypatch.setattr(security_config, "validate_production_environment", lambda: calls.append("security"))
    monkeypatch.setattr(redis.Redis, "from_url", staticmethod(lambda url: (calls.append(("redis", url)) or connection)))
    monkeypatch.setattr(rq, "Queue", lambda name, connection: (calls.append(("queue", name, connection)) or queue))

    class FakeWorker:
        def __init__(self, queues, connection):
            calls.append(("worker", queues, connection))

        def work(self):
            calls.append("work")

    monkeypatch.setattr(rq, "Worker", FakeWorker)
    worker.run()
    assert calls == [
        "security",
        ("redis", "rediss://redis.example:6379/7"),
        ("queue", "excel-auditor", connection),
        ("worker", [queue], connection),
        "work",
    ]


def test_retention_entrypoint_builds_adapters_and_reports_count(tmp_path, monkeypatch, capsys):
    import excel_auditor.retention as retention

    calls = []
    database = object()
    store = object()
    monkeypatch.setenv("DATABASE_URL", "postgresql+psycopg://db.example/auditor")
    monkeypatch.setenv("S3_BUCKET", "auditor")
    monkeypatch.setenv("S3_ENDPOINT_URL", "https://s3.example")
    monkeypatch.setenv("AWS_REGION", "us-east-1")
    monkeypatch.setenv("S3_SERVER_SIDE_ENCRYPTION", "AES256")
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "data"))
    monkeypatch.setattr(retention, "validate_production_environment", lambda: calls.append("security"))
    monkeypatch.setattr(retention, "DatabaseRepository", lambda url, create_schema: (calls.append(("database", url, create_schema)) or database))
    monkeypatch.setattr(retention, "S3ArtifactStore", lambda *args: (calls.append(("store", args)) or store))

    class FakeService:
        def __init__(self, root: Path, database, artifact_store):
            calls.append(("service", root, database, artifact_store))

        def purge_expired(self):
            calls.append("purge")
            return ["job_a", "job_b"]

    monkeypatch.setattr(retention, "AuditService", FakeService)
    retention.run()
    assert calls[0] == "security"
    assert ("database", "postgresql+psycopg://db.example/auditor", False) in calls
    assert ("store", ("auditor", "https://s3.example", "us-east-1", "AES256")) in calls
    assert ("service", (tmp_path / "data").resolve(), database, store) in calls
    assert calls[-1] == "purge"
    assert capsys.readouterr().out.strip() == "purged_jobs=2"
