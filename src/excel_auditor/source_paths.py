from __future__ import annotations

import re
from urllib.parse import unquote, urlsplit


_PARAMETER_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")


def validate_managed_http_path(value: str, *, field_name: str = "path") -> str:
    """Require a normalized origin-relative path that cannot escape an allowlist prefix."""
    parsed = urlsplit(value)
    if (
        not value.startswith("/")
        or value.startswith("//")
        or parsed.scheme
        or parsed.netloc
        or parsed.query
        or parsed.fragment
        or parsed.path != value
    ):
        raise ValueError(f"{field_name} must be a normalized absolute path without origin, query, or fragment")

    current = value
    for _depth in range(4):
        if re.search(r"%(?![0-9A-Fa-f]{2})", current):
            raise ValueError(f"{field_name} contains invalid percent encoding")
        try:
            decoded = unquote(current, errors="strict")
        except UnicodeDecodeError as exc:
            raise ValueError(f"{field_name} contains invalid UTF-8 percent encoding") from exc
        _validate_decoded_path(decoded, field_name)
        if decoded == current:
            return value
        current = decoded
    if "%" in current:
        raise ValueError(f"{field_name} contains excessive percent encoding")
    return value


def _validate_decoded_path(value: str, field_name: str) -> None:
    if (
        not value.startswith("/")
        or value.startswith("//")
        or "\\" in value
        or "?" in value
        or "#" in value
        or any(ord(character) < 32 or ord(character) == 127 for character in value)
    ):
        raise ValueError(f"{field_name} is not a safe normalized path")
    segments = value.split("/")
    if any(segment in {".", ".."} for segment in segments):
        raise ValueError(f"{field_name} cannot contain dot segments")
    if any(segment == "" for segment in segments[1:-1]):
        raise ValueError(f"{field_name} cannot contain repeated slashes")


def validate_simple_json_path(value: str | None, *, field_name: str) -> str | None:
    if value is None:
        return None
    if value == "$":
        return value
    if not value.startswith("$.") or any(not segment for segment in value[2:].split(".")):
        raise ValueError(f"{field_name} must be '$' or a simple non-empty object path beginning with '$.'")
    return value


def validate_parameter_name(value: str, *, field_name: str) -> str:
    if _PARAMETER_NAME.fullmatch(value) is None:
        raise ValueError(f"{field_name} must be a bounded HTTP parameter name")
    return value
