import json

import httpx
import pytest
from openpyxl import Workbook

from excel_auditor.models import StandardSourceConfig
from excel_auditor.standard_sources import ConnectionRegistry, ManagedHttpSource
from excel_auditor.service import AuditService
from excel_auditor.models import RuleSet


def _registry(tmp_path, **updates):
    connection = {
        "id": "hr", "base_url": "https://api.example.com/", "allowed_paths": ["/employees"],
        "max_records": 10,
    }
    connection.update(updates)
    path = tmp_path / "connections.json"
    path.write_text(json.dumps({"connections": [connection]}), encoding="utf-8")
    return ConnectionRegistry(path)


def test_managed_http_paginates_and_only_uses_registered_origin(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        page = int(request.url.params["page"])
        records = [{"id": "1"}, {"id": "2"}] if page == 1 else [{"id": "3"}]
        return httpx.Response(200, headers={"content-type": "application/json"}, json={"data": {"items": records, "total": 3}})

    source = ManagedHttpSource(_registry(tmp_path), httpx.MockTransport(handler), resolver=lambda _host: ["93.184.216.34"])
    config = StandardSourceConfig.model_validate({
        "type": "managed_http", "connection_id": "hr", "path": "/employees", "data_json_path": "$.data.items",
        "pagination": {"size": 2, "total_json_path": "$.data.total"},
    })
    assert [item["id"] for item in source.fetch(config)] == ["1", "2", "3"]


def test_managed_http_reads_secret_from_read_only_secret_directory(tmp_path, monkeypatch):
    secret_directory = tmp_path / "secrets"
    secret_directory.mkdir()
    (secret_directory / "hr-token").write_text("file-secret", encoding="utf-8")
    monkeypatch.setenv("EXCEL_AUDITOR_SECRET_DIR", str(secret_directory))
    monkeypatch.delenv("EXCEL_AUDITOR_SECRET_HR_TOKEN", raising=False)

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["Authorization"] == "Bearer file-secret"
        return httpx.Response(200, headers={"content-type": "application/json"}, json={"data": []})

    source = ManagedHttpSource(
        _registry(tmp_path, auth_secret_ref="hr-token"),
        httpx.MockTransport(handler),
        resolver=lambda _host: ["93.184.216.34"],
    )
    source.fetch(StandardSourceConfig(type="managed_http", connection_id="hr", path="/employees"))


def test_managed_connection_rejects_unknown_or_unsafe_configuration(tmp_path):
    with pytest.raises(ValueError, match="extra_forbidden"):
        _registry(tmp_path, unknown_option=True)
    with pytest.raises(ValueError, match="string_pattern_mismatch"):
        _registry(tmp_path, auth_secret_ref="../escape")


def test_managed_http_only_maps_declared_task_parameters(tmp_path):
    observed = {}
    def handler(request: httpx.Request) -> httpx.Response:
        observed.update(request.url.params)
        return httpx.Response(200, headers={"content-type": "application/json"}, json={"data": []})
    source = ManagedHttpSource(_registry(tmp_path), httpx.MockTransport(handler), resolver=lambda _host: ["93.184.216.34"])
    config = StandardSourceConfig.model_validate({"type": "managed_http", "connection_id": "hr", "path": "/employees", "static_parameters": {"active": "1"}, "parameter_mapping": {"department": "department_id"}})
    source.fetch(config, {"department_id": "D1", "untrusted": "must-not-pass"})
    assert observed == {"active": "1", "department": "D1"}


def test_managed_http_blocks_private_address_and_unlisted_path(tmp_path):
    source = ManagedHttpSource(_registry(tmp_path), httpx.MockTransport(lambda _request: httpx.Response(200)), resolver=lambda _host: ["169.254.169.254"])
    allowed = StandardSourceConfig(type="managed_http", connection_id="hr", path="/employees")
    with pytest.raises(ValueError, match="blocked"):
        source.fetch(allowed)
    source = ManagedHttpSource(_registry(tmp_path), resolver=lambda _host: ["93.184.216.34"])
    denied = StandardSourceConfig(type="managed_http", connection_id="hr", path="/admin")
    with pytest.raises(ValueError, match="not allowed"):
        source.fetch(denied)


def test_managed_http_stream_aborts_at_response_limit(tmp_path):
    source = ManagedHttpSource(
        _registry(tmp_path, max_response_bytes=1024),
        httpx.MockTransport(lambda _request: httpx.Response(200, headers={"content-type": "application/json"}, content=b" " * 1025)),
        resolver=lambda _host: ["93.184.216.34"],
    )
    config = StandardSourceConfig(type="managed_http", connection_id="hr", path="/employees")
    with pytest.raises(ValueError, match="exceeds configured size"):
        source.fetch(config)


def test_managed_http_is_snapshotted_by_service(tmp_path):
    transport = httpx.MockTransport(lambda _request: httpx.Response(200, headers={"content-type": "application/json"}, json={"data": [{"id": "E001", "name": "张三"}]}))
    source = ManagedHttpSource(_registry(tmp_path), transport, resolver=lambda _host: ["93.184.216.34"])
    rules = RuleSet.model_validate({
        "schema_id": "managed", "schema_version": "1.0.0", "name": "Managed",
        "sheets": [{"id": "people", "name": "人员", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "编号", "required": True}, {"name": "name", "title": "姓名"}
        ]}],
        "standard_source": {"type": "managed_http", "connection_id": "hr", "path": "/employees", "data_json_path": "$.data"},
    })
    excel = tmp_path / "input.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员"
    sheet.append(["编号", "姓名"])
    sheet.append(["E001", "张三"])
    book.save(excel)
    service = AuditService(tmp_path / "runtime", managed_http=source)
    job_id = service.create_job()
    service.run(job_id, excel, None, rules)
    status = service.status(job_id)
    assert status["status"] == "completed"
    assert status["summary"]["matched_records"] == 1
    assert list(service.job_directory(job_id).glob("std_*.jsonl"))
