import json
from pathlib import Path

from excel_auditor.engine import map_headers
from excel_auditor.models import RuleSet
from excel_auditor.engine import compare_workbook
from excel_auditor.workbook import SheetSnapshot, inspect_workbook
from openpyxl import Workbook
from excel_auditor.product_workflow import CatalogFieldDefinition, CatalogFieldSource, map_product_headers


BENCHMARK = json.loads(Path("tests/benchmarks/precision_recall.json").read_text(encoding="utf-8"))
CORE = json.loads(Path("tests/golden_files/core_scenarios.json").read_text(encoding="utf-8"))
PRODUCT_BENCHMARK = json.loads(Path("tests/benchmarks/product_mapping_precision_recall.json").read_text(encoding="utf-8"))


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


def test_distinct_headers_mapping_to_one_canonical_field_are_duplicates():
    rules = RuleSet.model_validate(CORE["rule"])
    sheet = rules.sheets[0]
    snapshot = SheetSnapshot(sheet.name, 1, 4, [(1, ["编号", "id", "姓名", "手机号"])], set(), [])

    mappings, columns = map_headers(sheet, snapshot)

    id_mappings = [item for item in mappings if item.canonical_field == "id"]
    assert [(item.physical_column, item.status) for item in id_mappings] == [(1, "duplicate"), (2, "duplicate")]
    assert "id" not in columns
    assert columns == {"name": 3, "phone": 4}


def test_product_mapping_precision_recall_and_manual_review_gate_meet_release_thresholds():
    fields = [
        CatalogFieldDefinition(
            field_id=item["field_id"],
            title=item["title"],
            aliases=item["aliases"],
            source=CatalogFieldSource.FIXED,
        )
        for item in PRODUCT_BENCHMARK["fields"]
    ]
    predicted_accepted = set()
    expected_accepted = set()
    predicted_review = set()
    expected_review = set()
    for case_index, case in enumerate(PRODUCT_BENCHMARK["cases"]):
        mappings = map_product_headers(case["headers"], fields, fuzzy_threshold=50)
        predicted_accepted.update(
            (case_index, mapping.physical_column, mapping.field_id)
            for mapping in mappings
            if mapping.status == "accepted"
        )
        expected_accepted.update(
            (case_index, int(column), field_id)
            for column, field_id in case["accepted"].items()
        )
        predicted_review.update(
            (case_index, mapping.physical_column, mapping.field_id or mapping.candidates[0].field_id)
            for mapping in mappings
            if mapping.status == "manual_review" and (mapping.field_id or mapping.candidates)
        )
        expected_review.update(
            (case_index, int(column), field_id)
            for column, field_id in case["review"].items()
        )
        assert all(
            mapping.status != "accepted"
            for mapping in mappings
            if mapping.match_type == "fuzzy_suggestion"
        )
    _assert_precision_recall(predicted_accepted, expected_accepted, "product automatic field mapping")
    _assert_precision_recall(predicted_review, expected_review, "product manual-review detection")


def _assert_precision_recall(predicted, expected, label):
    predicted, expected = set(predicted), set(expected)
    true_positive = len(predicted & expected)
    precision = true_positive / len(predicted) if predicted else 1.0
    recall = true_positive / len(expected) if expected else 1.0
    assert precision >= BENCHMARK["minimum_precision"], {"label": label, "precision": precision, "false_positives": sorted(predicted - expected)}
    assert recall >= BENCHMARK["minimum_recall"], {"label": label, "recall": recall, "false_negatives": sorted(expected - predicted)}


def test_record_field_and_repair_precision_recall_meet_release_thresholds(tmp_path):
    cases = BENCHMARK["record_field_repair_cases"]
    for case_index, case in enumerate(cases):
        rules = RuleSet.model_validate(case["rule"])
        path = tmp_path / f"annotated-{case_index}.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = rules.sheets[0].name
        sheet.append(case["excel"]["headers"])
        for row in case["excel"]["rows"]:
            sheet.append(row)
        book.save(path)
        workbook = inspect_workbook(path, rules)
        result = compare_workbook(workbook, case["standard"], rules)
        try:
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
            name = case.get("name", f"case-{case_index}")
            _assert_precision_recall(predicted_records, map(tuple, case["expected_record_labels"]), f"{name}: record matching")
            _assert_precision_recall(predicted_fields, map(tuple, case["expected_field_labels"]), f"{name}: field differences")
            _assert_precision_recall(predicted_repairs, map(tuple, case["expected_repair_labels"]), f"{name}: automatic repairs")
        finally:
            result.close()
            workbook.close()
