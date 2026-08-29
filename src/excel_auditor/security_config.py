from __future__ import annotations

import json
import os
from collections.abc import Mapping
from urllib.parse import parse_qs, urlsplit


def validate_production_environment(environment: Mapping[str, str] | None = None) -> None:
    env = os.environ if environment is None else environment
    if env.get("EXCEL_AUDITOR_ENV", "development").strip().lower() not in {"production", "prod"}:
        return

    errors: list[str] = []
    database_url = env.get("DATABASE_URL", "")
    if not database_url:
        errors.append("DATABASE_URL is required")
    else:
        parsed = urlsplit(database_url)
        ssl_mode = parse_qs(parsed.query).get("sslmode", [""])[-1].lower()
        if not parsed.scheme.startswith("postgresql"):
            errors.append("DATABASE_URL must use PostgreSQL")
        elif ssl_mode not in {"require", "verify-ca", "verify-full"}:
            errors.append("DATABASE_URL must require TLS with sslmode=require, verify-ca, or verify-full")

    redis_url = env.get("REDIS_URL", "")
    if not redis_url:
        errors.append("REDIS_URL is required")
    elif urlsplit(redis_url).scheme.lower() != "rediss":
        errors.append("REDIS_URL must use rediss TLS")

    if not env.get("S3_BUCKET"):
        errors.append("S3_BUCKET is required")
    endpoint = env.get("S3_ENDPOINT_URL", "")
    if endpoint and urlsplit(endpoint).scheme.lower() != "https":
        errors.append("S3_ENDPOINT_URL must use HTTPS")
    if env.get("S3_SERVER_SIDE_ENCRYPTION", "") not in {"AES256", "aws:kms"}:
        errors.append("S3_SERVER_SIDE_ENCRYPTION must be AES256 or aws:kms")

    renderer_command = env.get("EXCEL_RENDERER_COMMAND", "")
    if not renderer_command:
        errors.append("EXCEL_RENDERER_COMMAND is required")
    elif not os.path.isabs(renderer_command):
        errors.append("EXCEL_RENDERER_COMMAND must be an absolute path")

    if env.get("EXCEL_AUDITOR_REQUIRE_AUTH") != "1":
        errors.append("EXCEL_AUDITOR_REQUIRE_AUTH must be enabled")
    configured_tokens: list[str] = []
    if token := env.get("EXCEL_AUDITOR_API_TOKEN"):
        configured_tokens.append(token)
    if payload := env.get("EXCEL_AUDITOR_API_TOKENS_JSON"):
        try:
            parsed_tokens = json.loads(payload)
            if not isinstance(parsed_tokens, dict):
                raise ValueError
            configured_tokens.extend(str(token) for token in parsed_tokens)
        except (json.JSONDecodeError, ValueError, TypeError):
            errors.append("EXCEL_AUDITOR_API_TOKENS_JSON must be a JSON object")
    if not configured_tokens:
        errors.append("at least one API token is required")
    elif any(len(token) < 32 for token in configured_tokens):
        errors.append("API tokens must contain at least 32 characters")

    if errors:
        raise RuntimeError("PRODUCTION_CONFIG_INVALID: " + "; ".join(errors))
