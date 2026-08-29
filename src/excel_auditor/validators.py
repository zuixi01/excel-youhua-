"""Controlled cross-field validator registry.

Rules may reference only validators registered by trusted application code.
Configuration cannot import modules or execute expressions.
"""

from collections.abc import Callable
from typing import Any


Validator = Callable[[dict[str, Any], dict[str, Any]], tuple[str, str] | None]
_VALIDATORS: dict[str, Validator] = {}


def register_validator(name: str, validator: Validator) -> None:
    if not name or not name.replace("_", "").isalnum():
        raise ValueError("validator name must contain only letters, digits, and underscores")
    if name in _VALIDATORS and _VALIDATORS[name] is not validator:
        raise ValueError(f"validator is already registered: {name}")
    _VALIDATORS[name] = validator


def validator_names() -> frozenset[str]:
    return frozenset(_VALIDATORS)


def run_validator(name: str, parsed: dict[str, Any], params: dict[str, Any]) -> tuple[str, str] | None:
    try:
        validator = _VALIDATORS[name]
    except KeyError as exc:
        raise ValueError(f"unregistered cross-field validator: {name}") from exc
    return validator(parsed, params)


def _conditional_required(parsed: dict[str, Any], params: dict[str, Any]) -> tuple[str, str] | None:
    when = params.get("when_field")
    required = params.get("required_field")
    expected = params.get("equals")
    if when in parsed and required in parsed and parsed[when].normalized == expected and parsed[required].normalized is None:
        return str(required), f"when {when} equals {expected}, {required} is required"
    return None


def _date_order(parsed: dict[str, Any], params: dict[str, Any]) -> tuple[str, str] | None:
    start = params.get("start_field")
    end = params.get("end_field")
    if start in parsed and end in parsed:
        left, right = parsed[start], parsed[end]
        if left.valid and right.valid and left.normalized is not None and right.normalized is not None and right.normalized < left.normalized:
            return str(end), f"{end} must not be earlier than {start}"
    return None


register_validator("conditional_required", _conditional_required)
register_validator("date_order", _date_order)
