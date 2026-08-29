import json

from openpyxl import Workbook
from sqlalchemy import func, select
from sqlalchemy.orm import Session

from excel_auditor.models import RuleSet
from excel_auditor.persistence import DatabaseRepository, DifferenceIndexRow
from excel_auditor.record_store import DiskBackedRecordMap
from excel_auditor.rendering import ExcelRenderer
from excel_auditor.service import AuditService
from excel_auditor.spill import SpillableSequence


def test_spillable_sequence_preserves_order_and_removes_temporary_payload():
    values = SpillableSequence[int](spill_after_items=2)
    values.extend([10, 20, 30, 40])
    assert values.spilled is True
    assert len(values) == 4
    assert list(values) == [10, 20, 30, 40]
    assert values[-1] == 40
    values.close()


def test_disk_backed_record_map_supports_mapping_contract_and_cleanup():
    records = DiskBackedRecordMap()
    path = records.path
    first = (("string", "E1"),)
    second = (("string", "E2"),)
    records[first] = {"id": "E1", "name": "Alice"}
    records[second] = {"id": "E2", "name": "Bob"}
    assert list(records) == [first, second]
    assert records[first]["name"] == "Alice"
    assert records.pop(first)["id"] == "E1"
    assert first not in records and len(records) == 1
    records.close()
    assert not path.exists()


def test_large_difference_stream_uses_report_mode_and_keeps_complete_jsonl(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DIFFERENCE_SPILL_THRESHOLD", "2")
    monkeypatch.setenv("EXCEL_AUDITOR_POLARS_JOIN_THRESHOLD", "1")
    rules = RuleSet.model_validate({
        "schema_id": "difference-spill",
        "schema_version": "1.0.0",
        "name": "Difference spill",
        "sheets": [{
            "id": "data",
            "name": "Data",
            "primary_key": ["id"],
            "columns": [{"name": "id", "title": "ID", "required": True}],
        }],
    })
    excel = tmp_path / "input.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID"])
    sheet.append(["E0"])
    book.save(excel)
    standard = tmp_path / "standard.json"
    standard.write_text(json.dumps({"data": [{"id": f"E{index}"} for index in range(1_206)]}), encoding="utf-8")

    class MustNotRender(ExcelRenderer):
        def render(self, source, destination, workbook, rules, comparison, report_payload):
            raise AssertionError("spilled difference streams must use report-only mode")

    database = DatabaseRepository(f"sqlite:///{tmp_path / 'audit.db'}")
    service = AuditService(tmp_path / "runtime", renderer=MustNotRender(), database=database)
    job_id = service.create_job()
    service.run(job_id, excel, standard, rules)

    status = service.status(job_id)
    assert status["status"] == "completed" and status["mode"] == "report_only"
    assert "comparison_storage:disk_differences" in status["warnings"]
    assert "comparison_storage:memory_standard_records" in status["warnings"]
    manifest = json.loads(service.artifact(job_id, "manifest").read_text(encoding="utf-8"))
    assert manifest["rendering"]["reason"] == "large_difference_report_only"
    report = json.loads(service.artifact(job_id, "json").read_text(encoding="utf-8"))
    assert len(report["differences"]) == 1_205
    assert all(item["job_id"] == job_id and item["type"] == "MISSING_RECORD" for item in report["differences"])
    jsonl = [json.loads(line) for line in service.artifact(job_id, "differences_jsonl").read_text(encoding="utf-8").splitlines()]
    assert jsonl == report["differences"]
    with Session(database.engine) as session:
        assert session.scalar(select(func.count()).select_from(DifferenceIndexRow)) == 1_205
