import json
from decimal import Decimal

import httpx
import pytest
from openpyxl import Workbook

from excel_auditor.models import StandardSourceConfig
from excel_auditor.snapshots import SpilledRecords
from excel_auditor.standard_sources import ConnectionRegistry, ManagedHttpSource
from excel_auditor.service import AuditService, _canonicalize_standard
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


def test_managed_http_spills_paginated_records_and_transfers_ownership(tmp_path):
    def handler(request: httpx.Request) -> httpx.Response:
        page = int(request.url.params["page"])
        records = [{"id": "1"}, {"id": "2"}] if page == 1 else [{"id": "3"}]
        return httpx.Response(200, headers={"content-type": "application/json"}, json={"data": records, "total": 3})

    source = ManagedHttpSource(
        _registry(tmp_path),
        httpx.MockTransport(handler),
        resolver=lambda _host: ["93.184.216.34"],
        spill_after_records=2,
    )
    config = StandardSourceConfig.model_validate({
        "type": "managed_http", "connection_id": "hr", "path": "/employees", "data_json_path": "$.data",
        "pagination": {"size": 2, "total_json_path": "$.total"},
    })
    records, metadata = source.fetch_with_metadata(config)
    try:
        assert isinstance(records, SpilledRecords)
        assert [item["id"] for item in records] == ["1", "2", "3"]
        assert metadata["record_storage"] == "disk_spill"
    finally:
        records.close()


def test_managed_http_preserves_high_precision_decimal_tokens(tmp_path):
    content = b'{"data":[{"id":"1","amount":1234567890.1234567890123456789}]}'
    source = ManagedHttpSource(
        _registry(tmp_path),
        httpx.MockTransport(lambda _request: httpx.Response(
            200,
            headers={"content-type": "application/json"},
            content=content,
        )),
        resolver=lambda _host: ["93.184.216.34"],
    )
    config = StandardSourceConfig(
        type="managed_http",
        connection_id="hr",
        path="/employees",
        data_json_path="$.data",
    )

    records = source.fetch(config)
    assert records[0]["amount"] == Decimal("1234567890.1234567890123456789")


def test_managed_http_closes_spill_when_record_limit_aborts(tmp_path, monkeypatch):
    def handler(request: httpx.Request) -> httpx.Response:
        page = int(request.url.params["page"])
        records = [{"id": "1"}, {"id": "2"}] if page == 1 else [{"id": "3"}]
        return httpx.Response(200, headers={"content-type": "application/json"}, json={"data": records, "total": 3})

    closed = []
    original_close = SpilledRecords.close

    def recording_close(self):
        closed.append(True)
        original_close(self)

    monkeypatch.setattr(SpilledRecords, "close", recording_close)
    source = ManagedHttpSource(
        _registry(tmp_path, max_records=2),
        httpx.MockTransport(handler),
        resolver=lambda _host: ["93.184.216.34"],
        spill_after_records=1,
    )
    config = StandardSourceConfig.model_validate({
        "type": "managed_http", "connection_id": "hr", "path": "/employees", "data_json_path": "$.data",
        "pagination": {"size": 2, "total_json_path": "$.total"},
    })
    with pytest.raises(ValueError, match="record limit exceeded"):
        source.fetch(config)
    assert closed == [True]


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
    with pytest.raises(ValueError, match="dot segments"):
        _registry(tmp_path, allowed_paths=["/employees/../admin"])

    duplicate = tmp_path / "duplicate-connections.json"
    duplicate.write_text(
        '{"connections":[{"id":"hr","base_url":"https://first.example/","base_url":"https://second.example/","allowed_paths":["/employees"]}]}',
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="duplicate key"):
        ConnectionRegistry(duplicate)


@pytest.mark.parametrize(
    "path",
    [
        "/employees/../admin",
        "/employees/%2e%2e/admin",
        "/employees/%252e%252e/admin",
        "/employees\\..\\admin",
        "//example.com/employees",
        "/employees?admin=true",
        "/employees//admin",
    ],
)
def test_managed_source_rejects_non_normalized_or_allowlist_bypass_paths(path):
    with pytest.raises(ValueError, match="path|dot segments|slashes"):
        StandardSourceConfig(type="managed_http", connection_id="hr", path=path)


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


@pytest.mark.parametrize(
    ("content", "message"),
    [
        (b'{"data":[{"id":"E1","id":"E2"}]}', "duplicate object key"),
        (b'{"data":[', "response JSON is malformed"),
        (b'{"data":[{"id":"E1","amount":NaN}]}', "non-finite number"),
    ],
)
def test_managed_http_rejects_ambiguous_or_malformed_json(content, message, tmp_path):
    source = ManagedHttpSource(
        _registry(tmp_path),
        httpx.MockTransport(lambda _request: httpx.Response(200, headers={"content-type": "application/json"}, content=content)),
        resolver=lambda _host: ["93.184.216.34"],
    )
    config = StandardSourceConfig(type="managed_http", connection_id="hr", path="/employees")

    with pytest.raises(ValueError, match=message):
        source.fetch(config)


def test_managed_http_is_snapshotted_by_service(tmp_path):
    transport = httpx.MockTransport(lambda _request: httpx.Response(200, headers={"content-type": "application/json"}, json={"data": [{"id": "E001", "name": "张三"}, {"id": "E002", "name": "李四"}]}))
    source = ManagedHttpSource(_registry(tmp_path), transport, resolver=lambda _host: ["93.184.216.34"], spill_after_records=1)
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
    sheet.append(["E002", "李四"])
    book.save(excel)
    service = AuditService(tmp_path / "runtime", managed_http=source)
    job_id = service.create_job()
    service.run(job_id, excel, None, rules)
    status = service.status(job_id)
    assert status["status"] == "completed"
    assert status["summary"]["matched_records"] == 2
    assert list(service.job_directory(job_id).glob("std_*.jsonl"))


def test_managed_http_canonicalization_rejects_conflicting_aliases():
    rules = RuleSet.model_validate({
        "schema_id": "managed", "schema_version": "1.0.0", "name": "Managed",
        "sheets": [{"id": "people", "name": "人员", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "编号", "aliases": ["工号"], "required": True}
        ]}],
    })

    with pytest.raises(ValueError, match=r"STANDARD_DATA_INVALID: people\.id has conflicting field representations at record 1"):
        _canonicalize_standard({"people": [{"id": "E001", "工号": "E002"}]}, rules)


def test_managed_http_canonicalization_rejects_duplicate_sheet_mapping():
    rules = RuleSet.model_validate({
        "schema_id": "managed", "schema_version": "1.0.0", "name": "Managed",
        "sheets": [{"id": "people", "name": "人员", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "编号", "required": True}
        ]}],
    })

    with pytest.raises(ValueError, match="STANDARD_DATA_INVALID: duplicate sheet mapping: people"):
        _canonicalize_standard({"people": [{"id": "E001"}], "人员": [{"id": "E002"}]}, rules)
