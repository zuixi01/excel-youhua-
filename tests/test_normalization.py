import json
from decimal import Decimal, localcontext
from datetime import date

from excel_auditor.models import ColumnRule
from excel_auditor.normalization import excel_datetime_write_safe, parse_value, values_equal


def test_decimal_tolerance_boundary():
    rule = ColumnRule.model_validate({
        "name": "salary", "title": "工资", "type": "decimal",
        "compare": {"mode": "numeric", "absolute_tolerance": "0.01"},
    })
    assert values_equal(parse_value("10000.005", rule), parse_value("10000.00", rule), rule)
    assert not values_equal(parse_value("10000.02", rule), parse_value("10000.00", rule), rule)


def test_percentage_and_decimal_places_are_explicit():
    percent = ColumnRule.model_validate({"name": "rate", "title": "比例", "type": "decimal", "normalize": ["trim", "percent_to_decimal"], "compare": {"mode": "numeric", "decimal_places": 3}})
    assert parse_value("12.5%", percent).normalized == Decimal("0.125")
    assert values_equal(parse_value("12.54%", percent), parse_value("0.125", percent), percent)


def test_non_finite_numbers_are_invalid_and_large_decimals_quantize_safely():
    decimal_rule = ColumnRule.model_validate({
        "name": "amount", "title": "Amount", "type": "decimal",
        "compare": {"mode": "numeric", "decimal_places": 2},
    })
    integer_rule = ColumnRule.model_validate({"name": "count", "title": "Count", "type": "integer"})

    for raw in ("NaN", "Infinity", "-Infinity"):
        assert not parse_value(raw, decimal_rule).valid
        assert not parse_value(raw, integer_rule).valid

    left = parse_value("123456789012345678901234567890.124", decimal_rule)
    right = parse_value("123456789012345678901234567890.12", decimal_rule)
    assert left.valid and right.valid
    assert values_equal(left, right, decimal_rule)


def test_relative_tolerance_does_not_round_up_for_high_precision_values():
    rule = ColumnRule.model_validate({
        "name": "amount", "title": "Amount", "type": "decimal",
        "compare": {"mode": "numeric", "relative_tolerance": "0.000000000000000000000000000001"},
    })
    base = Decimal("1234567890123456789012345678901234567890")
    just_outside = Decimal("1234567890.12345678901234567895")
    just_inside = Decimal("1234567890.12345678901234567885")
    with localcontext() as context:
        context.prec = 120
        outside_value = base + just_outside
        inside_value = base + just_inside

    assert not values_equal(parse_value(str(base), rule), parse_value(str(outside_value), rule), rule)
    assert values_equal(parse_value(str(base), rule), parse_value(str(inside_value), rule), rule)


def test_string_primary_key_preserves_leading_zero():
    rule = ColumnRule(name="id", title="编号", type="string", normalize=["trim"])
    assert parse_value(" 001 ", rule).normalized == "001"


def test_chinese_date_format_is_parsed_deterministically():
    rule = ColumnRule(name="date", title="日期", type="date", parse_formats=["yyyy年M月d日"])
    assert parse_value("2026年8月28日", rule).normalized == date(2026, 8, 28)


def test_datetime_timezone_and_precision_are_applied():
    rule = ColumnRule.model_validate({
        "name": "event_at", "title": "时间", "type": "datetime",
        "compare": {"timezone": "Asia/Shanghai", "precision": "minute"},
    })
    left = parse_value("2026-08-28T10:30:59+08:00", rule)
    right = parse_value("2026-08-28T02:30:01Z", rule)
    assert values_equal(left, right, rule)
    naive = ColumnRule.model_validate({"name": "event_at", "title": "时间", "type": "datetime"})
    assert not parse_value("2026-08-28T10:30:00", naive).valid


def test_datetime_rejects_ambiguous_and_nonexistent_dst_local_times_without_offset():
    rule = ColumnRule.model_validate({
        "name": "event_at", "title": "Time", "type": "datetime",
        "compare": {"mode": "datetime", "timezone": "America/New_York"},
    })

    nonexistent = parse_value("2024-03-10T02:30:00", rule)
    ambiguous = parse_value("2024-11-03T01:30:00", rule)
    assert not nonexistent.valid and "nonexistent local datetime" in (nonexistent.error or "")
    assert not ambiguous.valid and "ambiguous local datetime" in (ambiguous.error or "")

    summer_occurrence = parse_value("2024-11-03T01:30:00-04:00", rule)
    winter_occurrence = parse_value("2024-11-03T01:30:00-05:00", rule)
    normal = parse_value("2024-11-03T03:30:00", rule)
    assert summer_occurrence.valid and winter_occurrence.valid and normal.valid
    assert not values_equal(summer_occurrence, winter_occurrence, rule)
    assert values_equal(summer_occurrence, parse_value("2024-11-03T05:30:00Z", rule), rule)
    assert not excel_datetime_write_safe(summer_occurrence.normalized, "America/New_York")
    assert not excel_datetime_write_safe(winter_occurrence.normalized, "America/New_York")
    assert excel_datetime_write_safe(normal.normalized, "America/New_York")

    day_rule = rule.model_copy(update={"compare": rule.compare.model_copy(update={"precision": "day"})})
    assert values_equal(
        parse_value("2024-11-03T01:30:00-04:00", day_rule),
        parse_value("2024-11-03T01:30:00-05:00", day_rule),
        day_rule,
    )


def test_boolean_aliases_are_explicit_and_disjoint():
    rule = ColumnRule.model_validate({"name": "enabled", "title": "启用", "type": "boolean", "boolean_true_values": ["启用"], "boolean_false_values": ["停用"]})
    assert parse_value("启用", rule).normalized is True
    assert parse_value("停用", rule).normalized is False
    assert not parse_value("是", rule).valid


def test_enum_ignore_case_applies_to_canonical_values_and_aliases():
    rule = ColumnRule.model_validate({
        "name": "status", "title": "Status", "type": "enum",
        "enum_values": ["Active", "Disabled"],
        "enum_aliases": {"ENABLED": "Active"},
        "compare": {"mode": "ignore_case"},
    })

    assert parse_value("active", rule).normalized == "Active"
    assert parse_value("enabled", rule).normalized == "Active"
    assert values_equal(parse_value("ACTIVE", rule), parse_value("Active", rule), rule)


def test_json_rejects_non_finite_numbers():
    rule = ColumnRule.model_validate({"name": "payload", "title": "Payload", "type": "json"})

    assert not parse_value('{"amount": NaN}', rule).valid
    assert not parse_value({"amount": float("inf")}, rule).valid


def test_json_preserves_high_precision_numeric_tokens_as_numbers():
    rule = ColumnRule.model_validate({"name": "payload", "title": "Payload", "type": "json"})
    exact = Decimal("1234567890.1234567890123456789")

    from_text = parse_value('{"amount":1234567890.1234567890123456789}', rule)
    from_object = parse_value({"amount": exact}, rule)

    assert from_text.valid and from_object.valid
    assert from_text.normalized == from_object.normalized
    assert '"amount":"' not in from_text.normalized
    assert json.loads(from_text.normalized, parse_float=Decimal)["amount"] == exact
