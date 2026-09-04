from __future__ import annotations

import json
import math
from decimal import Decimal
from typing import Any, Callable

import yaml
from yaml.nodes import MappingNode


def load_json_strict(
    document: str | bytes,
    *,
    context: str = "JSON document",
    preserve_decimal: bool = False,
) -> Any:
    def unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"{context} contains duplicate key: {key}")
            result[key] = value
        return result

    def finite_number(value: str) -> Any:
        raise ValueError(f"{context} contains non-finite number: {value}")

    try:
        return json.loads(
            document,
            object_pairs_hook=unique_object,
            parse_constant=finite_number,
            parse_float=Decimal if preserve_decimal else float,
        )
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise ValueError(f"{context} is malformed") from exc


def dump_json_exact(
    value: Any,
    *,
    ensure_ascii: bool = False,
    sort_keys: bool = False,
    default: Callable[[Any], Any] | None = None,
) -> str:
    """Encode JSON without converting Decimal values to binary floats or strings."""
    return _encode_json_exact(value, ensure_ascii=ensure_ascii, sort_keys=sort_keys, default=default)


def _encode_json_exact(
    value: Any,
    *,
    ensure_ascii: bool,
    sort_keys: bool,
    default: Callable[[Any], Any] | None,
) -> str:
    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, str):
        return json.dumps(value, ensure_ascii=ensure_ascii)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, Decimal):
        return _decimal_json_token(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError("non-finite numbers are not valid JSON")
        return _decimal_json_token(Decimal(str(value)))
    if isinstance(value, (list, tuple)):
        return "[" + ",".join(
            _encode_json_exact(item, ensure_ascii=ensure_ascii, sort_keys=sort_keys, default=default)
            for item in value
        ) + "]"
    if isinstance(value, dict):
        if any(not isinstance(key, str) for key in value):
            raise TypeError("JSON object keys must be strings")
        keys = sorted(value) if sort_keys else value.keys()
        return "{" + ",".join(
            json.dumps(key, ensure_ascii=ensure_ascii) + ":" + _encode_json_exact(
                value[key],
                ensure_ascii=ensure_ascii,
                sort_keys=sort_keys,
                default=default,
            )
            for key in keys
        ) + "}"
    if default is None:
        raise TypeError(f"not JSON serializable: {type(value).__name__}")
    converted = default(value)
    if converted is value:
        raise TypeError(f"JSON default returned the original {type(value).__name__} value")
    return _encode_json_exact(converted, ensure_ascii=ensure_ascii, sort_keys=sort_keys, default=default)


def _decimal_json_token(value: Decimal) -> str:
    if not value.is_finite():
        raise ValueError("non-finite numbers are not valid JSON")
    if value.is_zero():
        return "0"
    sign, raw_digits, exponent = value.as_tuple()
    digits = list(raw_digits)
    while len(digits) > 1 and digits[-1] == 0:
        digits.pop()
        exponent += 1
    coefficient = str(digits[0])
    if len(digits) > 1:
        coefficient += "." + "".join(str(digit) for digit in digits[1:])
    adjusted_exponent = exponent + len(digits) - 1
    exponent_text = f"+{adjusted_exponent}" if adjusted_exponent >= 0 else str(adjusted_exponent)
    return ("-" if sign else "") + coefficient + "E" + exponent_text


class _UniqueKeyLoader(yaml.SafeLoader):
    pass


def _construct_unique_mapping(loader: _UniqueKeyLoader, node: MappingNode, deep: bool = False) -> dict[Any, Any]:
    loader.flatten_mapping(node)
    mapping: dict[Any, Any] = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        try:
            duplicate = key in mapping
        except TypeError as exc:
            raise yaml.constructor.ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                "found an unhashable mapping key",
                key_node.start_mark,
            ) from exc
        if duplicate:
            raise yaml.constructor.ConstructorError(
                "while constructing a mapping",
                node.start_mark,
                f"found duplicate key {key!r}",
                key_node.start_mark,
            )
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_UniqueKeyLoader.add_constructor(
    yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
    _construct_unique_mapping,
)


def load_yaml_strict(document: str, *, context: str = "YAML document") -> Any:
    try:
        return yaml.load(document, Loader=_UniqueKeyLoader)
    except yaml.YAMLError as exc:
        raise ValueError(f"{context} is malformed or contains duplicate keys") from exc
