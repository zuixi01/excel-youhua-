import json
from pathlib import Path

from excel_auditor.engine import map_headers
from excel_auditor.models import RuleSet
from excel_auditor.engine import compare_workbook
from excel_auditor.workbook import SheetSnapshot, inspect_workbook
from openpyxl import Workbook


BENCHMARK = json.loads(Path("tests/benchmarks/precision_recall.json").read_text(encoding="utf-8"))
CORE = json.loads(Path("tests/golden_files/core_scenarios.json").read_text(encoding="utf-8"))


def test_confirmed_header_mapping_precision_and_recall_meet_release_thresholds():
    rules = RuleSet.model_validate(CORE["rule"])
    predicted: set[tuple[int, int, str]] = set()
    expected: set[tuple[int, int, str]] = set()
    for case_index, case in enumerate(BENCHMARK["header_cases"]):
        snapshot = SheetSnapshot(rules.sheets[0].name, 1, len(case["headers"]), [(1, case["headers"])], set(), [])
        mappings, _columns = map_headers(rules.sheets[0], snapshot)
        predicted.update((case_index, item.physical_column, item.canonical_field) for item in mappings if item.status == "matched" and item.canonical_field)
        expected.update((case_index, int(column), field) for column, field in case["expected"].items())
    true_positive = len(predicted & expected)
    precision = true_positive / len(predicted) if predicted else 1.0
    recall = true_positive / len(expected) if expected else 1.0
    assert precision >= BENCHMARK["minimum_precision"], {"precision": precision, "false_positives": sorted(predicted - expected)}
    assert recall >= BENCHMARK["minimum_recall"], {"recall": recall, "false_negatives": sorted(expected - predicted)}


def _assert_precision_recall(predicted, expected, label):
    predicted, expected = set(predicted), set(expected)
    true_positive = len(predicted & expected)
    precision = true_positive / len(predicted) if predicted else 1.0
    recall = true_positive / len(expected) if expected else 1.0
    assert precision >= BENCHMARK["minimum_precision"], {"label": label, "precision": precision, "false_positives": sorted(predicted - expected)}
    assert recall >= BENCHMARK["minimum_recall"], {"label": label, "recall": recall, "false_negatives": sorted(expected - predicted)}


def test_record_field_and_repair_precision_recall_meet_release_thresholds(tmp_path):
    case = BENCHMARK["record_field_repair_case"]
    rules = RuleSet.model_validate(case["rule"])
    path = tmp_path / "annotated.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(case["excel"]["headers"])
    for row in case["excel"]["rows"]:
        sheet.append(row)
    book.save(path)
    result = compare_workbook(inspect_workbook(path, rules), case["standard"], rules)
    differences = {item.difference_id: item for item in result.differences}
    record_types = {"EXTRA_RECORD", "MISSING_RECORD"}
    predicted_records = [
        (item.type.value, item.business_key["id"])
        for item in result.differences
        if item.type.value in record_types and item.business_key
    ]
    predicted_fields = [
        (item.type.value, item.business_key["id"], item.canonical_field)
        for item in result.differences
        if item.type.value in {"VALUE_MISMATCH", "INVALID_VALUE", "VALIDATION_ERROR"} and item.business_key
    ]
    predicted_repairs = [
        (repair.type, differences[repair.difference_id].business_key["id"], repair.canonical_field)
        for repair in result.repairs
    ]
    _assert_precision_recall(predicted_records, map(tuple, case["expected_record_labels"]), "record matching")
    _assert_precision_recall(predicted_fields, map(tuple, case["expected_field_labels"]), "field differences")
    _assert_precision_recall(predicted_repairs, map(tuple, case["expected_repair_labels"]), "automatic repairs")
