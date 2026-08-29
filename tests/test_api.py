import asyncio
from io import BytesIO
import json
import os
from pathlib import Path

from fastapi import HTTPException, UploadFile
from fastapi.testclient import TestClient
from openpyxl import Workbook
import pytest


def test_publish_create_and_download_api(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "api-data"))
    import importlib
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    rules = Path("configs/examples/employee-roster.yaml").read_text(encoding="utf-8")
    import yaml
    response = client.post("/api/v1/schemas/publish", json=yaml.safe_load(rules))
    assert response.status_code == 201, response.text
    workbook_path = tmp_path / "input.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员信息"
    sheet.append(["员工数据"])
    sheet.append(["工号", "姓名", "工资", "部门"])
    sheet.append(["E001", "张三", "100", "技术部"])
    book.save(workbook_path)
    precheck = client.post(
        "/api/v1/workbooks/precheck",
        data={"schema_id": "employee-roster", "schema_version": "1.0.0"},
        files={"excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert precheck.status_code == 200, precheck.text
    assert precheck.json()["structure"][0]["sheet_name"] == "人员信息"
    assert "auto_repair_authorizations" in precheck.json()
    response = client.post(
        "/api/v1/comparisons",
        headers={"Idempotency-Key": "comparison-001"},
        data={"schema_id": "employee-roster", "schema_version": "1.0.0"},
        files={
            "excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "standard_data": ("standard.json", json.dumps({"employees": [{"employee_id": "E001", "employee_name": "张三", "salary": "100", "department": "技术部"}]}).encode(), "application/json"),
        },
    )
    assert response.status_code == 202, response.text
    job_id = response.json()["job_id"]
    replay = client.post(
        "/api/v1/comparisons",
        headers={"Idempotency-Key": "comparison-001"},
        data={"schema_id": "employee-roster", "schema_version": "1.0.0"},
        files={
            "excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "standard_data": ("standard.json", json.dumps({"employees": [{"employee_id": "E001", "employee_name": "张三", "salary": "100", "department": "技术部"}]}).encode(), "application/json"),
        },
    )
    assert replay.status_code == 200 and replay.json()["job_id"] == job_id and replay.json()["idempotent_replay"] is True
    conflict = client.post(
        "/api/v1/comparisons",
        headers={"Idempotency-Key": "comparison-001"},
        data={"schema_id": "employee-roster", "schema_version": "1.0.0"},
        files={
            "excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "standard_data": ("standard.json", b'{"employees":[]}', "application/json"),
        },
    )
    assert conflict.status_code == 409
    assert not list((api_module.DATA_ROOT / "incoming").iterdir())
    status = client.get(f"/api/v1/comparisons/{job_id}")
    assert status.json()["status"] == "completed", status.text
    artifact = client.get(f"/api/v1/comparisons/{job_id}/artifacts/json")
    assert artifact.status_code == 200
    assert artifact.json()["summary"]["matched_records"] == 1
    api_module.service.artifact(job_id, "json").unlink()
    differences = client.get(f"/api/v1/comparisons/{job_id}/differences?page=1&page_size=1")
    assert differences.status_code == 200
    assert len(differences.json()["items"]) <= 1
    missing = client.get("/api/v1/comparisons/job_doesnotexist")
    assert missing.status_code == 404
    assert missing.headers["content-type"].startswith("application/problem+json")
    inline = client.post(
        "/api/v1/comparisons",
        headers={"Idempotency-Key": "comparison-inline"},
        data={
            "schema_id": "employee-roster",
            "schema_version": "1.0.0",
            "standard_json": json.dumps({"employees": [{"employee_id": "E001", "employee_name": "张三", "salary": "100", "department": "技术部"}]}),
        },
        files={"excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert inline.status_code == 202, inline.text
    inline_job = inline.json()["job_id"]
    assert client.get(f"/api/v1/comparisons/{inline_job}").json()["status"] == "completed"
    duplicate_inline = client.post(
        "/api/v1/comparisons",
        data={
            "schema_id": "employee-roster",
            "schema_version": "1.0.0",
            "standard_json": '{"employees":[],"employees":[]}',
        },
        files={"excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert duplicate_inline.status_code == 422
    assert "standard_json is not valid JSON" in duplicate_inline.json()["detail"]
    duplicate_parameters = client.post(
        "/api/v1/comparisons",
        data={
            "schema_id": "employee-roster",
            "schema_version": "1.0.0",
            "parameters": '{"department":"D1","department":"D2"}',
            "standard_json": '{"employees":[]}',
        },
        files={"excel_file": ("input.xlsx", workbook_path.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert duplicate_parameters.status_code == 422
    assert duplicate_parameters.json()["detail"] == "parameters must be a JSON object"
    deleted = client.delete(f"/api/v1/comparisons/{inline_job}")
    assert deleted.status_code == 202 and deleted.json()["deleted_at"]
    assert client.get(f"/api/v1/comparisons/{inline_job}").status_code == 404


def test_standard_upload_is_staged_in_bounded_chunks_and_cleans_oversize(tmp_path, monkeypatch):
    import excel_auditor.api as api_module

    monkeypatch.setattr(api_module, "DATA_ROOT", tmp_path)
    payload = b"x" * (2 * 1024 * 1024 + 17)
    upload = UploadFile(filename="standard.json", file=BytesIO(payload))
    requested_sizes: list[int] = []
    original_read = upload.read

    async def tracked_read(size: int = -1) -> bytes:
        requested_sizes.append(size)
        return await original_read(size)

    monkeypatch.setattr(upload, "read", tracked_read)
    staged, size = asyncio.run(api_module._stage_standard_upload(upload, len(payload)))
    assert size == len(payload)
    assert staged.read_bytes() == payload
    assert requested_sizes and set(requested_sizes) == {1024 * 1024}
    staged.unlink()

    oversized = UploadFile(filename="standard.json", file=BytesIO(b"12345"))
    with pytest.raises(HTTPException) as raised:
        asyncio.run(api_module._stage_standard_upload(oversized, 4))
    assert raised.value.status_code == 413
    assert not list((tmp_path / "incoming").iterdir())


def test_bearer_auth_can_be_required(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "auth-data"))
    monkeypatch.setenv("EXCEL_AUDITOR_API_TOKEN", "test-secret")
    import importlib
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    assert client.get("/health/live").status_code == 200
    assert client.get("/api/v1/schemas/missing/versions/1.0.0").status_code == 401
    authorized = client.get("/api/v1/schemas/missing/versions/1.0.0", headers={"Authorization": "Bearer test-secret"})
    assert authorized.status_code == 404


def test_precheck_uses_rule_specific_upload_limit(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "precheck-limit"))
    import importlib
    import yaml
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    config = yaml.safe_load(Path("configs/examples/employee-roster.yaml").read_text(encoding="utf-8"))
    config["schema_id"] = "precheck-limit"
    config["workbook"]["max_upload_mib"] = 1
    assert client.post("/api/v1/schemas/publish", json=config).status_code == 201
    response = client.post(
        "/api/v1/workbooks/precheck",
        data={"schema_id": "precheck-limit", "schema_version": config["schema_version"]},
        files={"excel_file": ("oversized.xlsx", b"x" * (1024 * 1024 + 1), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 413
    assert response.json()["detail"] == "FILE_LIMIT_EXCEEDED"


def test_readiness_executes_renderer_self_check(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "renderer-ready"))
    monkeypatch.delenv("DATABASE_URL", raising=False)
    monkeypatch.delenv("REDIS_URL", raising=False)
    monkeypatch.delenv("S3_BUCKET", raising=False)
    import importlib
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)

    class Renderer:
        def self_check(self):
            return "9.8.7"

    monkeypatch.setattr(api_module, "DotNetOpenXmlRenderer", Renderer)
    api_module.service.renderer = Renderer()
    response = TestClient(api_module.app).get("/health/ready")
    assert response.status_code == 200
    assert response.json()["checks"]["renderer"] == "ExcelRenderer/9.8.7"


def test_validation_problem_details_never_echo_untrusted_values(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "safe-errors"))
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKEN", raising=False)
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKENS_JSON", raising=False)
    monkeypatch.delenv("EXCEL_AUDITOR_REQUIRE_AUTH", raising=False)
    import importlib
    import yaml
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    secret = "sensitive-value-must-not-be-reflected"
    config = yaml.safe_load(Path("configs/examples/employee-roster.yaml").read_text(encoding="utf-8"))
    config["schema_version"] = secret
    response = client.post("/api/v1/schemas/publish", json=config)
    assert response.status_code == 422
    assert response.headers["content-type"].startswith("application/problem+json")
    assert secret not in response.text
    payload = response.json()
    assert payload["detail"] == "Request body or parameters did not satisfy the API schema"
    assert payload["errors"] == [{"loc": ["body", "schema_version"], "type": "value_error"}]


def test_job_resources_are_tenant_isolated(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "tenant-data"))
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKEN", raising=False)
    monkeypatch.setenv("EXCEL_AUDITOR_API_TOKENS_JSON", json.dumps({
        "token-a": {"tenant_id": "tenant-a", "user_id": "alice"},
        "token-b": {"tenant_id": "tenant-b", "user_id": "bob"},
    }))
    import importlib
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    job_id = api_module.service.create_job(tenant_id="tenant-a", user_id="alice")
    denied = client.get(f"/api/v1/comparisons/{job_id}", headers={"Authorization": "Bearer token-b"})
    assert denied.status_code == 404
    allowed = client.get(f"/api/v1/comparisons/{job_id}", headers={"Authorization": "Bearer token-a"})
    assert allowed.status_code == 200 and allowed.json()["tenant_id"] == "tenant-a"
    import yaml
    config = yaml.safe_load(Path("configs/examples/employee-roster.yaml").read_text(encoding="utf-8"))
    published = client.post("/api/v1/schemas/publish", json=config, headers={"Authorization": "Bearer token-a"})
    assert published.status_code == 201, published.text
    assert client.get("/api/v1/schemas/employee-roster/versions/1.0.0", headers={"Authorization": "Bearer token-b"}).status_code == 404
    assert client.get("/api/v1/schemas/employee-roster/versions/1.0.0", headers={"Authorization": "Bearer token-a"}).status_code == 200


def test_draft_validate_mapping_publish_and_version_list(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "draft-api"))
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKEN", raising=False)
    import importlib
    import yaml
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    config = yaml.safe_load(Path("configs/examples/employee-roster.yaml").read_text(encoding="utf-8"))
    imported = client.post("/api/v1/schemas/import", files={"file": ("rules.yaml", yaml.safe_dump(config, allow_unicode=True).encode("utf-8"), "application/yaml")})
    assert imported.status_code == 201, imported.text
    created = client.post("/api/v1/schemas", json=config)
    assert created.status_code == 201, created.text
    draft_id = created.json()["draft_id"]
    confirmed = client.post("/api/v1/mappings/confirm", json={"schema_id": "employee-roster", "draft_id": draft_id, "sheet_id": "employees", "raw_header": "员工ID号", "canonical_field": "employee_id"})
    assert confirmed.status_code == 200, confirmed.text
    validation = client.post(f"/api/v1/schemas/employee-roster/drafts/{draft_id}/validate")
    assert validation.json()["valid"] is True
    published = client.post(f"/api/v1/schemas/employee-roster/drafts/{draft_id}/publish")
    assert published.status_code == 201, published.text
    versions = client.get("/api/v1/schemas/employee-roster/versions")
    assert versions.json()["total"] == 1
    exported = client.get("/api/v1/schemas/employee-roster/versions/1.0.0/export?format=yaml")
    assert exported.status_code == 200 and yaml.safe_load(exported.text)["schema_id"] == "employee-roster"
    immutable = client.put(f"/api/v1/schemas/employee-roster/drafts/{draft_id}", json=config)
    assert immutable.status_code == 409


@pytest.mark.parametrize(
    ("filename", "content"),
    [
        ("rules.json", b'{"schema_id":"first","schema_id":"second"}'),
        ("rules.yaml", b"schema_id: first\nschema_id: second\n"),
    ],
)
def test_rule_import_rejects_duplicate_keys(tmp_path, monkeypatch, filename, content):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "duplicate-rule-api"))
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKEN", raising=False)
    import importlib
    import excel_auditor.api as api_module

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    response = client.post("/api/v1/schemas/import", files={"file": (filename, content, "application/octet-stream")})

    assert response.status_code == 422
    assert response.json()["detail"] == "RULE_CONFIG_INVALID: uploaded rule document is invalid"


def test_cancel_request_moves_job_to_cancelled_at_safe_checkpoint(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DATA", str(tmp_path / "cancel-api"))
    monkeypatch.delenv("EXCEL_AUDITOR_API_TOKEN", raising=False)
    import importlib
    import excel_auditor.api as api_module
    from excel_auditor.rules import load_rules

    api_module = importlib.reload(api_module)
    client = TestClient(api_module.app)
    job_id = api_module.service.create_job()
    requested = client.post(f"/api/v1/comparisons/{job_id}/cancel")
    assert requested.status_code == 202 and requested.json()["cancel_requested"] is True
    api_module.service.run(job_id, tmp_path / "not-read.xlsx", None, load_rules(Path("configs/examples/employee-roster.yaml")))
    assert api_module.service.status(job_id)["status"] == "cancelled"
    terminal = client.post(f"/api/v1/comparisons/{job_id}/cancel")
    assert terminal.status_code == 409
