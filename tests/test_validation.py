import pytest
from openpyxl import Workbook

from excel_auditor.engine import compare_workbook
from excel_auditor.models import RuleSet
from excel_auditor.pandera_adapter import StandardDataValidator
from excel_auditor.snapshots import SpilledRecords
from excel_auditor.workbook import inspect_workbook


def test_non_default_sheet_actions_are_honored(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "actions", "schema_version": "1.0.0", "name": "Actions",
        "sheets": [{
            "id": "data", "name": "Data", "primary_key": ["id"],
            "actions": {"extra_header": "report_only", "extra_record": "report_only", "mismatched_value": "report_only", "invalid_value": "report_only", "duplicate_key": "report_only"},
            "columns": [
                {"name": "id", "title": "ID", "required": True},
                {"name": "amount", "title": "Amount", "type": "integer"},
            ],
        }],
    })
    path = tmp_path / "actions.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Amount", "Extra"])
    sheet.append(["E1", "bad", "x"])
    sheet.append(["E2", "2", "y"])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": "E1", "amount": 1}]}, rules)
    configurable = {
        "EXTRA_HEADER", "EXTRA_RECORD", "VALUE_MISMATCH", "INVALID_VALUE", "VALIDATION_ERROR", "DUPLICATE_PRIMARY_KEY"
    }
    assert all(item.render_action == "report_only" for item in result.differences if item.type.value in configurable)


def test_header_auto_detection_and_order_mismatch(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "header-detect", "schema_version": "1.0.0", "name": "Header detect",
        "sheets": [{
            "id": "data", "name": "Data", "header": {"row": 1, "auto_detect": True}, "primary_key": ["id"],
            "columns": [
                {"name": "id", "title": "ID", "required": True},
                {"name": "name", "title": "Name"},
            ],
        }],
    })
    path = tmp_path / "header-detect.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["Report title"])
    sheet.append(["Name", "ID"])
    sheet.append(["Alice", "E1"])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": "E1", "name": "Alice"}]}, rules)
    assert not result.manual_review_reasons
    assert any(item.type.value == "HEADER_ORDER_MISMATCH" for item in result.differences)
    assert result.summary.matched_records == 1


def test_ambiguous_auto_detected_header_requires_review(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "header-ambiguous", "schema_version": "1.0.0", "name": "Header ambiguous",
        "sheets": [{"id": "data", "name": "Data", "header": {"auto_detect": True}, "primary_key": ["id"], "columns": [{"name": "id", "title": "ID", "required": True}]}],
    })
    path = tmp_path / "header-ambiguous.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID"])
    sheet.append(["ID"])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": []}, rules)
    assert result.manual_review_reasons
    assert [item.type.value for item in result.differences] == ["HEADER_NOT_FOUND"]


def test_unmatched_records_still_receive_field_and_cross_field_validation(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "quality", "schema_version": "1.0.0", "name": "Quality",
        "sheets": [{
            "id": "staff", "name": "员工", "primary_key": ["id"],
            "columns": [
                {"name": "id", "title": "编号", "required": True},
                {"name": "email", "title": "邮箱", "validation": {"regex": "^[^@]+@[^@]+$", "unique": True}},
                {"name": "status", "title": "状态"},
                {"name": "end_date", "title": "离职日期", "type": "date"},
            ],
            "cross_field_rules": [{"rule_id": "end-required", "validator": "conditional_required", "params": {"when_field": "status", "equals": "离职", "required_field": "end_date"}}],
        }],
    })
    path = tmp_path / "quality.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "员工"
    sheet.append(["编号", "邮箱", "状态", "离职日期"])
    sheet.append(["E001", "bad", "离职", None])
    sheet.append(["E002", "bad", "在职", None])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"staff": []}, rules)
    rule_ids = [item.rule_id for item in result.differences]
    assert "email.validation" in rule_ids
    assert rule_ids.count("email.unique") == 2
    assert "end-required" in rule_ids


def test_fuzzy_string_is_suggestion_only_even_when_overwrite_is_enabled(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "fuzzy", "schema_version": "1.0.0", "name": "Fuzzy",
        "sheets": [{"id": "data", "name": "数据", "primary_key": ["id"], "actions": {"overwrite_mismatch": True}, "columns": [
            {"name": "id", "title": "编号", "required": True},
            {"name": "address", "title": "地址", "type": "fuzzy_string"},
        ]}],
    })
    path = tmp_path / "fuzzy.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "数据"
    sheet.append(["编号", "地址"])
    sheet.append(["E1", "上海市浦东新区"])
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": "E1", "address": "上海浦东新区"}]}, rules)
    difference = next(item for item in result.differences if item.canonical_field == "address")
    assert difference.severity == "warning" and difference.repair_status == "not_requested"
    assert not result.repairs


def test_omitted_optional_standard_field_never_clears_excel_but_explicit_null_can(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "omitted-standard", "schema_version": "1.0.0", "name": "Omitted standard",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"],
            "actions": {"overwrite_mismatch": True},
            "columns": [
                {"name": "id", "title": "ID", "required": True},
                {"name": "note", "title": "Note"},
            ],
        }],
    })
    path = tmp_path / "omitted-standard.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Note"])
    sheet.append(["E1", "keep me"])
    book.save(path)

    omitted = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": "E1"}]}, rules)
    assert omitted.summary.matched_records == 1
    assert not [item for item in omitted.differences if item.canonical_field == "note"]
    assert not omitted.repairs

    explicit_null = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": "E1", "note": None}]}, rules)
    note_difference = next(item for item in explicit_null.differences if item.canonical_field == "note")
    assert note_difference.type.value == "VALUE_MISMATCH"
    assert note_difference.repair_status == "planned"
    assert [(repair.type, repair.canonical_field, repair.value) for repair in explicit_null.repairs] == [
        ("set_cell", "note", None)
    ]


def test_empty_primary_key_policy_can_match_by_row_number_without_silent_default(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "empty-key", "schema_version": "1.0.0", "name": "Empty key",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "empty_primary_key_action": "use_row_number", "columns": [
            {"name": "id", "title": "ID", "required": True}, {"name": "value", "title": "Value"},
        ]}],
    })
    path = tmp_path / "empty-key.xlsx"
    book = Workbook(); sheet = book.active; sheet.title = "Data"; sheet.append(["ID", "Value"]); sheet.append([None, "A"]); sheet.append([None, "B"]); book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": None, "value": "A"}, {"id": None, "value": "B"}]}, rules)
    assert result.summary.matched_records == 2
    assert all(item.type.value != "EMPTY_PRIMARY_KEY" for item in result.differences)


def test_empty_primary_key_policy_can_explicitly_skip_rows(tmp_path):
    rules = RuleSet.model_validate({
        "schema_id": "skip-key", "schema_version": "1.0.0", "name": "Skip key",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "empty_primary_key_action": "skip_row", "columns": [{"name": "id", "title": "ID", "required": True}]}],
    })
    path = tmp_path / "skip-key.xlsx"
    book = Workbook(); sheet = book.active; sheet.title = "Data"; sheet.append(["ID"]); sheet.append([None]); book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), {"data": [{"id": None}]}, rules)
    assert result.summary.differences == 0 and result.summary.matched_records == 0


def test_standard_validation_is_chunked_and_checks_uniqueness_across_chunks():
    rules = RuleSet.model_validate({
        "schema_id": "chunked-standard",
        "schema_version": "1.0.0",
        "name": "Chunked standard",
        "sheets": [{
            "id": "data",
            "name": "Data",
            "primary_key": ["id"],
            "columns": [
                {"name": "id", "title": "ID", "required": True},
                {"name": "code", "title": "Code", "validation": {"unique": True}},
            ],
        }],
    })
    rows = SpilledRecords()
    for record in ({"id": "E1", "code": "A"}, {"id": "E2", "code": "B"}, {"id": "E3", "code": "A"}):
        rows.append(record)
    try:
        with pytest.raises(ValueError, match="is not unique at record 3"):
            StandardDataValidator(chunk_size=2).validate({"data": rows}, rules)
    finally:
        rows.close()
