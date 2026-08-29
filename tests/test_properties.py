from decimal import Decimal
import json
from pathlib import Path

from hypothesis import given, settings, strategies as st

from excel_auditor.engine import compare_workbook
from excel_auditor.models import ColumnRule, RuleSet, normalize_header
from excel_auditor.normalization import apply_normalizers, parse_value, values_equal
from excel_auditor.workbook import SheetSnapshot, WorkbookSnapshot


@given(st.text())
def test_safe_text_normalization_is_idempotent(value):
    pipeline = ["trim", "unicode_nfkc", "collapse_spaces", "uppercase"]
    once = apply_normalizers(value, pipeline)
    assert apply_normalizers(once, pipeline) == once


@given(st.decimals(allow_nan=False, allow_infinity=False, places=4, min_value=Decimal("-1000000"), max_value=Decimal("1000000")), st.decimals(allow_nan=False, allow_infinity=False, places=4, min_value=Decimal("0"), max_value=Decimal("100")))
def test_decimal_absolute_tolerance_matches_definition(value, delta):
    tolerance = Decimal("0.01")
    rule = ColumnRule.model_validate({"name": "amount", "title": "金额", "type": "decimal", "compare": {"mode": "numeric", "absolute_tolerance": str(tolerance)}})
    left, right = parse_value(str(value), rule), parse_value(str(value + delta), rule)
    assert values_equal(left, right, rule) is (delta <= tolerance)


@given(st.text(), st.text())
def test_header_normalization_is_idempotent(left, right):
    assert normalize_header(normalize_header(left)) == normalize_header(left)
    if normalize_header(left) != normalize_header(right):
        assert ("string", normalize_header(left)) != ("string", normalize_header(right))


@given(st.decimals(allow_nan=False, allow_infinity=False, places=4), st.decimals(allow_nan=False, allow_infinity=False, places=4))
def test_numeric_equality_is_symmetric(left_value, right_value):
    rule = ColumnRule.model_validate({"name": "amount", "title": "Amount", "type": "decimal", "compare": {"mode": "numeric", "absolute_tolerance": "0.01", "relative_tolerance": "0.001"}})
    left, right = parse_value(left_value, rule), parse_value(right_value, rule)
    assert values_equal(left, right, rule) == values_equal(right, left, rule)


@given(st.lists(st.text(min_size=1).filter(lambda value: "," not in value), max_size=20))
def test_set_parsing_is_order_and_duplicate_independent(values):
    rule = ColumnRule.model_validate({"name": "tags", "title": "Tags", "type": "set", "compare": {"mode": "set"}})
    forward = parse_value(",".join(values), rule)
    reverse = parse_value(",".join([*reversed(values), *values]), rule)
    assert values_equal(forward, reverse, rule)


@given(st.dictionaries(st.text(min_size=1, max_size=12), st.integers(), max_size=12))
def test_json_comparison_is_key_order_independent(payload):
    rule = ColumnRule.model_validate({"name": "payload", "title": "Payload", "type": "json"})
    left = parse_value(json.dumps(payload, sort_keys=True), rule)
    right = parse_value(json.dumps(dict(reversed(list(payload.items())))), rule)
    assert values_equal(left, right, rule)


@given(st.booleans())
def test_boolean_aliases_round_trip(value):
    rule = ColumnRule.model_validate({"name": "active", "title": "Active", "type": "boolean"})
    aliases = ("true", "1", "yes", "y") if value else ("false", "0", "no", "n")
    assert all(parse_value(alias, rule).normalized is value for alias in aliases)


@given(st.text(max_size=80))
def test_value_aliases_are_deterministic(value):
    rule = ColumnRule.model_validate({"name": "code", "title": "Code", "value_aliases": {"N/A": ""}, "regex_replacements": [{"pattern": r"\s+", "replacement": " "}]})
    assert parse_value(value, rule) == parse_value(value, rule)


@given(
    st.lists(st.integers(min_value=1, max_value=10_000), min_size=1, max_size=12, unique=True),
    st.data(),
)
@settings(max_examples=25, deadline=None)
def test_standard_record_order_does_not_change_semantic_differences_or_repairs(ids, data):
    rules = RuleSet.model_validate({
        "schema_id": "order-property", "schema_version": "1.0.0", "name": "Order property",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"],
            "actions": {"overwrite_mismatch": True},
            "columns": [
                {"name": "id", "title": "ID", "required": True, "type": "integer"},
                {"name": "amount", "title": "Amount", "type": "decimal", "compare": {"mode": "numeric"}},
            ],
        }],
    })
    excel_rows = [(1, ["ID", "Amount"])] + [
        (row_number, [identifier, identifier * 10 + (1 if identifier == ids[0] else 0)])
        for row_number, identifier in enumerate(ids, start=2)
    ]
    snapshot = WorkbookSnapshot(
        Path("memory.xlsx"), "fixed",
        {"Data": SheetSnapshot("Data", len(excel_rows), 2, excel_rows)},
    )
    standard = [{"id": identifier, "amount": identifier * 10} for identifier in ids]
    order = data.draw(st.permutations(tuple(range(len(standard)))))
    shuffled = [standard[index] for index in order]

    first = compare_workbook(snapshot, {"data": standard}, rules)
    second = compare_workbook(snapshot, {"data": shuffled}, rules)
    try:
        def semantics(result):
            differences = {
                item.difference_id: (item.type.value, json.dumps(item.business_key, sort_keys=True), item.canonical_field)
                for item in result.differences
            }
            repairs = sorted(
                (repair.type, differences[repair.difference_id][1], repair.canonical_field, str(repair.value))
                for repair in result.repairs
            )
            return sorted(differences.values()), repairs

        assert semantics(first) == semantics(second)
    finally:
        first.close()
        second.close()
