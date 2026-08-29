import pytest

from excel_auditor.security_config import validate_production_environment


def _production_environment() -> dict[str, str]:
    return {
        "EXCEL_AUDITOR_ENV": "production",
        "DATABASE_URL": "postgresql+psycopg://user:secret@database:5432/auditor?sslmode=verify-full",
        "REDIS_URL": "rediss://redis:6379/0",
        "S3_BUCKET": "auditor",
        "S3_ENDPOINT_URL": "https://objects.example.test",
        "S3_SERVER_SIDE_ENCRYPTION": "AES256",
        "EXCEL_RENDERER_COMMAND": "/app/renderer/ExcelRenderer",
        "EXCEL_AUDITOR_REQUIRE_AUTH": "1",
        "EXCEL_AUDITOR_API_TOKEN": "a" * 32,
    }


def test_development_environment_does_not_require_production_services():
    validate_production_environment({})
    validate_production_environment({"EXCEL_AUDITOR_ENV": "development"})


def test_secure_production_environment_is_accepted():
    validate_production_environment(_production_environment())


def test_insecure_production_environment_fails_without_echoing_secrets():
    environment = _production_environment()
    environment.update({
        "DATABASE_URL": "postgresql+psycopg://user:database-secret@database:5432/auditor",
        "REDIS_URL": "redis://redis-secret@redis:6379/0",
        "S3_ENDPOINT_URL": "http://objects.example.test",
        "S3_SERVER_SIDE_ENCRYPTION": "",
        "EXCEL_AUDITOR_REQUIRE_AUTH": "0",
        "EXCEL_AUDITOR_API_TOKEN": "short-secret",
    })
    with pytest.raises(RuntimeError) as captured:
        validate_production_environment(environment)
    message = str(captured.value)
    assert "PRODUCTION_CONFIG_INVALID" in message
    assert "sslmode" in message and "rediss" in message and "HTTPS" in message
    assert "database-secret" not in message and "redis-secret" not in message and "short-secret" not in message


def test_production_rejects_missing_or_relative_renderer_command():
    environment = _production_environment()
    environment.pop("EXCEL_RENDERER_COMMAND")
    with pytest.raises(RuntimeError, match="EXCEL_RENDERER_COMMAND is required"):
        validate_production_environment(environment)
    environment["EXCEL_RENDERER_COMMAND"] = "renderer/ExcelRenderer"
    with pytest.raises(RuntimeError, match="EXCEL_RENDERER_COMMAND must be an absolute path"):
        validate_production_environment(environment)
