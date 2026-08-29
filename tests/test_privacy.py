import json
from pathlib import Path

from openpyxl import Workbook

from excel_auditor.engine import compare_workbook
from excel_auditor.models import RuleSet
from excel_auditor.rendering import ExcelRenderer
from excel_auditor.service import AuditService
from excel_auditor.workbook import inspect_workbook


def test_sensitive_values_are_masked_in_differences(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "pii", "schema_version": "1.0.0", "name": "PII",
        "sheets": [{
            "id": "people", "name": "人员", "primary_key": ["id"],
            "columns": [
                {"name": "id", "title": "身份证号", "required": True, "type": "id_code", "sensitive": True},
                {"name": "phone", "title": "手机号", "type": "phone", "sensitive": True},
            ],
        }],
    })
    path = tmp_path / "pii.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员"
    sheet.append(["身份证号", "手机号"])
    sheet.append(["123456789012345678", "13800138000"])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"people": [{"id": "123456789012345678", "phone": "13900139000"}]}, rules)
    mismatch = next(item for item in result.differences if item.canonical_field == "phone")
    assert mismatch.excel_raw_value == "13***00"
    assert mismatch.standard_raw_value == "13***00"
    assert mismatch.business_key == {"id": "12***78"}


def test_failure_status_and_diagnostic_never_store_exception_message(tmp_path):
    secret = "sensitive-renderer-detail-must-not-leak"

    class FailingRenderer(ExcelRenderer):
        def render(self, source, destination, workbook, rules, comparison, report_payload):
            raise RuntimeError(f"RENDER_FAILED: {secret}")

    rules = RuleSet.model_validate({
        "schema_id": "safe-failure", "schema_version": "1.0.0", "name": "Safe failure",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "columns": [{"name": "id", "title": "ID", "required": True}]}],
    })
    workbook_path = tmp_path / "input.xlsx"
    book = Workbook()
    book.active.title = "Data"
    book.active.append(["ID"])
    book.active.append(["E1"])
    book.save(workbook_path)
    standard_path = tmp_path / "standard.json"
    standard_path.write_text(json.dumps({"data": [{"id": "E1"}]}), encoding="utf-8")
    service = AuditService(tmp_path / "runtime", renderer=FailingRenderer())
    job_id = service.create_job()
    service.run(job_id, workbook_path, standard_path, rules)
    status = service.status(job_id)
    diagnostic = (service.job_directory(job_id) / "diagnostic.log").read_text(encoding="utf-8")
    assert status["status"] == "failed" and status["error_code"] == "RENDER_FAILED"
    assert status["error_message_safe"] == "Workbook rendering failed."
    assert secret not in json.dumps(status) and secret not in diagnostic
