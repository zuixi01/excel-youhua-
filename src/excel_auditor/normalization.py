from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation, localcontext
from typing import Any
from zoneinfo import ZoneInfo

from openpyxl.utils.datetime import from_excel

from .models import ColumnRule, FieldType
from .strict_serialization import dump_json_exact


@dataclass(frozen=True)
class ParsedValue:
    raw: Any
    normalized: Any
    valid: bool
    error: str | None = None


def is_formula_text(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def parse_row_number(value: Any) -> int:
    """Parse an exact positive Excel row number without lossy coercion."""
    if isinstance(value, bool):
        raise ValueError("row number must be an integer")
    if isinstance(value, int):
        parsed = value
    elif isinstance(value, str) and re.fullmatch(r"[0-9]+", value):
        parsed = int(value)
    else:
        raise ValueError("row number must be an integer")
    if not 1 <= parsed <= 1_048_576:
        raise ValueError("row number is outside the Excel row range")
    return parsed


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def apply_normalizers(value: Any, names: list[str]) -> Any:
    current = value
    for name in names:
        if name == "trim":
            current = _text(current).strip()
        elif name == "unicode_nfkc":
            current = unicodedata.normalize("NFKC", _text(current))
        elif name == "collapse_spaces":
            current = " ".join(_text(current).split())
        elif name == "uppercase":
            current = _text(current).upper()
        elif name == "lowercase" or name == "casefold":
            current = _text(current).casefold()
        elif name == "remove_group_separator":
            current = _text(current).replace(",", "").replace("，", "")
        elif name == "remove_currency_symbol":
            current = re.sub(r"[¥￥$€£]", "", _text(current)).strip()
        elif name == "percent_to_decimal":
            text = _text(current).strip()
            current = Decimal(text[:-1].strip()) / Decimal("100") if text.endswith(("%", "％")) else text
        else:
            raise ValueError(f"unknown normalizer: {name}")
    return current


def parse_value(value: Any, rule: ColumnRule) -> ParsedValue:
    try:
        if rule.compare.formula_mode == "formula" and is_formula_text(value):
            return ParsedValue(value, value, True)
        normalized = apply_normalizers(value, rule.normalize)
        if rule.regex_replacements and normalized is not None:
            text = str(normalized)
            for replacement in rule.regex_replacements:
                flags = re.IGNORECASE if replacement.ignore_case else 0
                text = re.sub(replacement.pattern, replacement.replacement, text, flags=flags)
            normalized = text
        if rule.value_aliases and normalized is not None:
            normalized = rule.value_aliases.get(str(normalized), normalized)
        if normalized is None or normalized == "":
            return ParsedValue(value, None, True)
        if rule.type in {FieldType.STRING, FieldType.PHONE, FieldType.ID_CODE, FieldType.POSTAL_CODE, FieldType.FUZZY_STRING}:
            return ParsedValue(value, str(normalized), True)
        if rule.type == FieldType.INTEGER:
            decimal = Decimal(str(normalized))
            if not decimal.is_finite():
                raise ValueError("not a finite integer")
            if decimal != decimal.to_integral_value():
                raise ValueError("not an integer")
            return ParsedValue(value, int(decimal), True)
        if rule.type == FieldType.DECIMAL:
            decimal = Decimal(str(normalized))
            if not decimal.is_finite():
                raise ValueError("not a finite decimal")
            return ParsedValue(value, decimal, True)
        if rule.type == FieldType.DATE:
            return ParsedValue(value, _parse_date(normalized, rule.parse_formats), True)
        if rule.type == FieldType.DATETIME:
            return ParsedValue(value, _parse_datetime(normalized, rule.parse_formats, rule.compare.timezone, rule.compare.allow_naive_datetime), True)
        if rule.type == FieldType.BOOLEAN:
            if isinstance(normalized, bool):
                return ParsedValue(value, normalized, True)
            lowered = str(normalized).strip().casefold()
            if lowered in {str(item).strip().casefold() for item in rule.boolean_true_values}:
                return ParsedValue(value, True, True)
            if lowered in {str(item).strip().casefold() for item in rule.boolean_false_values}:
                return ParsedValue(value, False, True)
            raise ValueError("not a recognized boolean")
        if rule.type == FieldType.ENUM:
            text = str(normalized)
            if rule.compare.mode == "ignore_case":
                aliases = {str(alias).casefold(): target for alias, target in rule.enum_aliases.items()}
                mapped = aliases.get(text.casefold(), text)
                canonical = {str(value).casefold(): value for value in rule.enum_values}
                mapped = canonical.get(str(mapped).casefold(), mapped)
            else:
                mapped = rule.enum_aliases.get(text, text)
            if mapped not in rule.enum_values:
                raise ValueError(f"not in enum: {mapped}")
            return ParsedValue(value, mapped, True)
        if rule.type == FieldType.SET:
            if isinstance(normalized, (list, tuple, set)):
                items = normalized
            else:
                items = str(normalized).split(rule.separator)
            return ParsedValue(value, tuple(sorted({str(item).strip() for item in items if str(item).strip()})), True)
        if rule.type == FieldType.JSON:
            obj = normalized if isinstance(normalized, (dict, list)) else json.loads(
                str(normalized),
                parse_float=Decimal,
                parse_constant=lambda token: (_ for _ in ()).throw(ValueError(f"non-finite JSON number: {token}")),
            )
            stable = dump_json_exact(obj, ensure_ascii=False, sort_keys=True)
            return ParsedValue(value, stable, True)
        raise ValueError(f"unsupported type: {rule.type}")
    except (ValueError, TypeError, InvalidOperation, json.JSONDecodeError) as exc:
        return ParsedValue(value, None, False, str(exc))


def parse_excel_value(value: Any, rule: ColumnRule, epoch: datetime) -> ParsedValue:
    """Parse an Excel cell value, interpreting numeric date serials by workbook epoch."""
    if rule.type not in {FieldType.DATE, FieldType.DATETIME} or isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return parse_value(value, rule)
    try:
        converted = from_excel(float(value), epoch=epoch)
    except (TypeError, ValueError, OverflowError) as exc:
        return ParsedValue(value, None, False, f"invalid Excel date serial: {exc}")
    parsed = parse_value(converted, rule)
    return ParsedValue(value, parsed.normalized, parsed.valid, parsed.error)


def _parse_date(value: Any, formats: list[str]) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    python_formats = [_to_python_format(item) for item in formats]
    for fmt in [*python_formats, "%Y-%m-%d", "%Y/%m/%d"]:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return datetime.fromisoformat(text).date()


def _parse_datetime(value: Any, formats: list[str], timezone_name: str | None, allow_naive: bool) -> datetime:
    if isinstance(value, datetime):
        parsed = value
    else:
        text = str(value).strip()
        parsed = None
        for fmt in [_to_python_format(item) for item in formats]:
            try:
                parsed = datetime.strptime(text, fmt)
                break
            except ValueError:
                pass
        if parsed is None:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    target = ZoneInfo(timezone_name) if timezone_name else None
    if parsed.tzinfo is None:
        if target is not None:
            parsed = parsed.replace(tzinfo=target)
        elif not allow_naive:
            raise ValueError("naive datetime requires compare.timezone or allow_naive_datetime=true")
        else:
            return parsed
    return parsed.astimezone(target or timezone.utc)


def _to_python_format(value: str) -> str:
    tokens = {"yyyy": "%Y", "MM": "%m", "dd": "%d", "HH": "%H", "mm": "%M", "ss": "%S", "M": "%m", "d": "%d"}
    pattern = re.compile("|".join(sorted(tokens, key=len, reverse=True)))
    return pattern.sub(lambda match: tokens[match.group(0)], value)


def values_equal(left: ParsedValue, right: ParsedValue, rule: ColumnRule) -> bool:
    if not left.valid or not right.valid:
        return False
    if left.normalized is None or right.normalized is None:
        return left.normalized is right.normalized
    if rule.type == FieldType.DECIMAL or rule.compare.mode == "numeric":
        a, b = Decimal(left.normalized), Decimal(right.normalized)
        if rule.compare.decimal_places is not None:
            places = rule.compare.decimal_places
            quantum = Decimal(1).scaleb(-places)
            required_precision = max(
                28,
                len(a.as_tuple().digits),
                len(b.as_tuple().digits),
                a.adjusted() + places + 1,
                b.adjusted() + places + 1,
            )
            with localcontext() as context:
                context.prec = required_precision
                a, b = a.quantize(quantum), b.quantize(quantum)
        delta = abs(a - b)
        absolute = rule.compare.absolute_tolerance
        relative = rule.compare.relative_tolerance * max(abs(a), abs(b))
        return delta <= absolute or delta <= relative
    if rule.compare.mode == "ignore_case":
        return str(left.normalized).casefold() == str(right.normalized).casefold()
    if rule.type == FieldType.DATETIME:
        return _truncate_datetime(left.normalized, rule.compare.precision) == _truncate_datetime(right.normalized, rule.compare.precision)
    return left.normalized == right.normalized


def _truncate_datetime(value: datetime, precision: str) -> datetime | date:
    if precision == "day":
        return value.date()
    if precision == "hour":
        return value.replace(minute=0, second=0, microsecond=0)
    if precision == "minute":
        return value.replace(second=0, microsecond=0)
    return value.replace(microsecond=0)
