from decimal import Decimal
from datetime import date

from excel_auditor.models import ColumnRule
from excel_auditor.normalization import parse_value, values_equal


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


def test_boolean_aliases_are_explicit_and_disjoint():
    rule = ColumnRule.model_validate({"name": "enabled", "title": "启用", "type": "boolean", "boolean_true_values": ["启用"], "boolean_false_values": ["停用"]})
    assert parse_value("启用", rule).normalized is True
    assert parse_value("停用", rule).normalized is False
    assert not parse_value("是", rule).valid
