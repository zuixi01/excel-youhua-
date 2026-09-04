import json
from pathlib import Path

from excel_auditor.engine import map_headers
from excel_auditor.models import RuleSet
from excel_auditor.engine import compare_workbook
from excel_auditor.workbook import SheetSnapshot, inspect_workbook
from openpyxl import Workbook
from excel_auditor.product_workflow import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CategoryDefinition,
    map_product_headers,
    resolve_categories,
)


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


def test_product_category_resolution_matches_adversarial_release_annotations():
    categories = [CategoryDefinition.model_validate(item) for item in PRODUCT_BENCHMARK["categories"]]
    cases = PRODUCT_BENCHMARK["category_cases"]
    results = resolve_categories(
        [case["row"] for case in cases],
        categories,
        fuzzy_threshold=60,
        candidate_score_margin=5,
    )

    assert len(results) == len(cases)
    for result, expected in zip(results, cases, strict=True):
        assert result.status == expected["status"]
        assert result.match_type == expected["match_type"]
        assert result.category_id == expected.get("category_id")
        assert [candidate.field_id for candidate in result.candidates] == expected["candidates"]


def test_product_category_resolution_has_perfect_accuracy_on_500_annotated_adversarial_cases():
    category_names = [
        "Smartphone Devices", "Running Footwear", "Laptop Computers", "Digital Cameras", "Wireless Headphones",
        "Kitchen Appliances", "Office Furniture", "Outdoor Lighting", "Fitness Equipment", "Camping Supplies",
        "Baby Clothing", "Pet Nutrition", "Garden Machinery", "Automotive Tools", "Musical Instruments",
        "Board Games", "Skincare Products", "Haircare Products", "Travel Luggage", "Wrist Watches",
        "Home Textiles", "Dining Utensils", "Coffee Equipment", "Tea Accessories", "Cleaning Supplies",
        "Bathroom Fixtures", "Bedroom Furniture", "Livingroom Furniture", "Storage Organizers", "Power Tools",
        "Hand Tools", "Protective Equipment", "Cycling Accessories", "Swimming Equipment", "Winter Sports",
        "Team Sports", "Educational Toys", "Building Blocks", "Art Supplies", "Writing Instruments",
        "Computer Components", "Network Equipment", "Mobile Accessories", "Audio Equipment", "Video Equipment",
        "Smart Lighting", "Security Cameras", "Climate Control", "Water Filtration", "Renewable Energy",
    ]
    categories = [
        CategoryDefinition(
            category_id=f"catalog-{index:03d}",
            name=name,
            aliases=[f"Merchant {name}"],
        )
        for index, name in enumerate(category_names)
    ]
    rows = []
    expected = []
    for index, category in enumerate(categories):
        next_category = categories[(index + 1) % len(categories)]
        first_word, suffix = category.name.split(" ", 1)
        middle = max(1, len(first_word) // 2)
        deleted = first_word[:middle] + first_word[middle + 1:] + " " + suffix
        swapped = (
            first_word[:middle]
            + first_word[middle + 1]
            + first_word[middle]
            + first_word[middle + 2:]
            + " "
            + suffix
        )
        annotated = [
            ({"platform_category_id": category.category_id, "merchant_category": category.name}, "resolved", "id", category.category_id),
            ({"merchant_category": category.name}, "resolved", "exact", category.category_id),
            ({"merchant_category": category.aliases[0]}, "resolved", "exact", category.category_id),
            ({"merchant_category": category.name.lower()}, "resolved", "exact", category.category_id),
            ({"merchant_category": f"  {first_word}   {suffix}  "}, "resolved", "exact", category.category_id),
            ({"merchant_category": deleted}, "manual_review", "fuzzy_suggestion", category.category_id),
            ({"merchant_category": swapped}, "manual_review", "fuzzy_suggestion", category.category_id),
            ({"platform_category_id": "retired-id", "merchant_category": category.name}, "manual_review", "invalid_id", category.category_id),
            ({"platform_category_id": category.category_id, "merchant_category": next_category.name}, "manual_review", "id_name_conflict", category.category_id),
            ({}, "unresolved", "missing", None),
        ]
        for row, status, match_type, target in annotated:
            rows.append(row)
            expected.append((status, match_type, target, next_category.category_id if match_type == "id_name_conflict" else None))

    results = resolve_categories(rows, categories, fuzzy_threshold=70, candidate_score_margin=5)
    assert len(results) == len(expected) == 500
    true_positive = false_positive = false_negative = 0
    for result, (status, match_type, target, conflict_target) in zip(results, expected, strict=True):
        assert result.status == status
        assert result.match_type == match_type
        candidate_ids = [candidate.field_id for candidate in result.candidates]
        predicted = result.category_id or (candidate_ids[0] if candidate_ids else None)
        if target is None:
            false_positive += int(predicted is not None)
            continue
        if result.category_id == target or target in candidate_ids:
            true_positive += 1
        else:
            false_negative += 1
        if predicted is not None and predicted != target and match_type != "id_name_conflict":
            false_positive += 1
        if conflict_target is not None:
            assert set(candidate_ids) == {target, conflict_target}

    precision = true_positive / (true_positive + false_positive)
    recall = true_positive / (true_positive + false_negative)
    assert precision == 1.0
    assert recall == 1.0


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
