from decimal import Decimal
import json
from pathlib import Path

from hypothesis import given, settings, strategies as st

from excel_auditor.engine import compare_workbook
from excel_auditor.models import ColumnRule, RuleSet, normalize_header
from excel_auditor.product_workflow import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CatalogSchemaSnapshot,
    build_dynamic_schema,
    map_product_headers,
)
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


@given(st.lists(st.from_regex(r"[A-Za-z0-9]{1,8}", fullmatch=True).map(lambda value: f"商家_{value}"), max_size=20, unique=True))
def test_dynamic_product_schema_never_loses_or_reorders_merchant_extra_headers(extra_headers):
    fixed = [ColumnRule(name="product_id", title="商品ID", required=True)]
    platform = [CatalogFieldDefinition(
        field_id="brand",
        title="品牌",
        source=CatalogFieldSource.PLATFORM_ATTRIBUTE,
        category_id="phone",
        attribute_id="brand",
    )]
    targets = [CatalogFieldDefinition(
        field_id="product_id",
        title="商品ID",
        source=CatalogFieldSource.FIXED,
        required=True,
    ), *platform]
    headers = ["商品ID", "品牌", *extra_headers]
    mappings = map_product_headers(headers, targets, fuzzy_threshold=100)
    plan = build_dynamic_schema(
        category_id="phone",
        category_name="手机",
        fixed_columns=fixed,
        platform_fields=platform,
        headers=headers,
        mappings=mappings,
    )
    merchant_fields = [item for item in plan.fields if item.field.source == CatalogFieldSource.MERCHANT_EXTRA]
    assert [item.source_header for item in merchant_fields] == extra_headers
    assert [item.field.title for item in merchant_fields] == extra_headers
    assert len({item.field.field_id for item in plan.fields}) == len(plan.fields)


@given(st.permutations(("brand", "material", "color")))
def test_catalog_snapshot_hash_and_field_order_are_independent_of_api_record_order(order):
    fields_by_id = {
        field_id: CatalogFieldDefinition(
            field_id=field_id,
            title=field_id.title(),
            source=CatalogFieldSource.PLATFORM_ATTRIBUTE,
            category_id="phone",
            attribute_id=field_id,
            display_order={"brand": 1, "material": 2, "color": 3}[field_id],
        )
        for field_id in ("brand", "material", "color")
    }
    snapshot = CatalogSchemaSnapshot.create(
        snapshot_id="catalog_property",
        connection_id="platform",
        category_id="phone",
        fields=[fields_by_id[field_id] for field_id in order],
    )
    canonical = CatalogSchemaSnapshot.create(
        snapshot_id="catalog_property",
        connection_id="platform",
        category_id="phone",
        fields=[fields_by_id[field_id] for field_id in ("brand", "material", "color")],
    )
    assert snapshot.content_sha256 == canonical.content_sha256
    assert [field.field_id for field in snapshot.fields] == ["brand", "material", "color"]
