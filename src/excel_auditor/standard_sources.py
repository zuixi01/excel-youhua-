from __future__ import annotations

import ipaddress
import os
import socket
import time
from pathlib import Path
from typing import Any, Callable, Sequence
from urllib.parse import urljoin, urlparse

import httpx
from pydantic import BaseModel, ConfigDict, Field, field_validator

from .models import StandardSourceConfig
from .observability import metrics
from .snapshots import SpilledRecords
from .source_paths import validate_managed_http_path
from .strict_serialization import load_json_strict


class ManagedConnection(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: str = Field(pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")
    base_url: str
    allowed_paths: list[str] = Field(min_length=1)
    auth_secret_ref: str | None = Field(default=None, pattern=r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,127}$")
    auth_header: str = "Authorization"
    auth_prefix: str = "Bearer "
    timeout_seconds: float = Field(default=15, ge=1, le=60)
    max_response_bytes: int = Field(default=16 * 1024 * 1024, ge=1024, le=64 * 1024 * 1024)
    max_records: int = Field(default=500_000, ge=1, le=1_000_000)
    allow_private_network: bool = False
    allow_insecure_http: bool = False

    @field_validator("base_url")
    @classmethod
    def validate_base_url(cls, value: str) -> str:
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.username or parsed.password:
            raise ValueError("base_url must be an HTTP(S) origin without credentials")
        if parsed.query or parsed.fragment:
            raise ValueError("base_url cannot contain query or fragment")
        return value.rstrip("/") + "/"

    @field_validator("auth_header")
    @classmethod
    def validate_auth_header(cls, value: str) -> str:
        if not value or not all(character.isalnum() or character in "!#$%&'*+-.^_`|~" for character in value):
            raise ValueError("auth_header must be an HTTP token")
        return value

    @field_validator("auth_prefix")
    @classmethod
    def validate_auth_prefix(cls, value: str) -> str:
        if "\r" in value or "\n" in value:
            raise ValueError("auth_prefix cannot contain a line break")
        return value

    @field_validator("allowed_paths")
    @classmethod
    def validate_allowed_paths(cls, values: list[str]) -> list[str]:
        return [validate_managed_http_path(value, field_name="allowed_paths entry") for value in values]


class ConnectionRegistry:
    def __init__(self, path: Path) -> None:
        payload = load_json_strict(path.read_text(encoding="utf-8"), context="connection registry JSON")
        connections = payload.get("connections", payload) if isinstance(payload, dict) else payload
        if not isinstance(connections, list):
            raise ValueError("connection registry must contain a connections array")
        parsed = [ManagedConnection.model_validate(item) for item in connections]
        self.connections = {item.id: item for item in parsed}
        if len(self.connections) != len(parsed):
            raise ValueError("connection ids must be unique")

    def get(self, connection_id: str) -> ManagedConnection:
        if connection_id not in self.connections:
            raise KeyError(f"managed connection not found: {connection_id}")
        return self.connections[connection_id]


class _RecordAccumulator:
    """Own paginated response records and spill their payloads past a limit."""

    def __init__(self, spill_after_records: int) -> None:
        if spill_after_records < 1:
            raise ValueError("spill_after_records must be positive")
        self.spill_after_records = spill_after_records
        self.records: list[dict[str, Any]] | SpilledRecords = []
        self.detached = False

    def __enter__(self) -> "_RecordAccumulator":
        return self

    def __exit__(self, _exc_type: Any, _exc: Any, _traceback: Any) -> None:
        if not self.detached:
            self.close()

    def __len__(self) -> int:
        return len(self.records)

    def extend(self, page_records: list[dict[str, Any]]) -> None:
        if isinstance(self.records, list) and len(self.records) + len(page_records) > self.spill_after_records:
            spilled = SpilledRecords()
            try:
                for record in self.records:
                    spilled.append(record)
            except Exception:
                spilled.close()
                raise
            self.records = spilled
        if isinstance(self.records, list):
            self.records.extend(page_records)
        else:
            for record in page_records:
                self.records.append(record)

    def detach(self) -> Sequence[dict[str, Any]]:
        self.detached = True
        return self.records

    def close(self) -> None:
        close = getattr(self.records, "close", None)
        if close is not None:
            close()


class ManagedHttpSource:
    def __init__(self, registry: ConnectionRegistry, transport: httpx.BaseTransport | None = None, resolver: Callable[[str], list[str]] | None = None, spill_after_records: int = 50_000) -> None:
        self.registry = registry
        self.transport = transport
        self.resolver = resolver or _resolve_addresses
        if spill_after_records < 1:
            raise ValueError("spill_after_records must be positive")
        self.spill_after_records = spill_after_records

    def fetch(self, config: StandardSourceConfig, parameters: dict[str, Any] | None = None) -> Sequence[dict[str, Any]] | dict[str, list[dict[str, Any]]]:
        records, _metadata = self.fetch_with_metadata(config, parameters)
        return records

    def fetch_with_metadata(self, config: StandardSourceConfig, parameters: dict[str, Any] | None = None) -> tuple[Sequence[dict[str, Any]] | dict[str, list[dict[str, Any]]], dict[str, Any]]:
        if config.type != "managed_http" or not config.connection_id or not config.path:
            raise ValueError("managed HTTP source configuration is incomplete")
        connection = self.registry.get(config.connection_id)
        resolved_addresses = _validate_target(connection, config.path, self.resolver)
        headers = {"Accept": "application/json"}
        if connection.auth_secret_ref:
            secret = _load_secret(connection.auth_secret_ref)
            headers[connection.auth_header] = connection.auth_prefix + secret
        url = urljoin(connection.base_url, config.path.lstrip("/"))
        request_url = url
        sni_hostname: str | None = None
        if self.transport is None:
            # Connect to the already-authorized address so a second DNS lookup
            # cannot rebind the hostname to loopback, link-local or metadata IPs.
            parsed_url = urlparse(url)
            address = resolved_addresses[0]
            host = f"[{address}]" if ":" in address else address
            if parsed_url.port:
                host += f":{parsed_url.port}"
            request_url = parsed_url._replace(netloc=host).geturl()
            original_host = parsed_url.hostname or ""
            default_port = 443 if parsed_url.scheme == "https" else 80
            headers["Host"] = original_host if parsed_url.port in {None, default_port} else f"{original_host}:{parsed_url.port}"
            sni_hostname = original_host if parsed_url.scheme == "https" else None
        workbook_records: dict[str, list[dict[str, Any]]] | None = None
        total_response_bytes = 0
        page_metadata: list[dict[str, int]] = []
        page = 1
        with _RecordAccumulator(self.spill_after_records) as records, httpx.Client(timeout=connection.timeout_seconds, headers=headers, follow_redirects=False, transport=self.transport) as client:
            while True:
                request_parameters = dict(config.static_parameters)
                task_parameters = parameters or {}
                for request_name, task_name in config.parameter_mapping.items():
                    if task_name not in task_parameters:
                        raise ValueError(f"STANDARD_SOURCE_FAILED: required task parameter is missing: {task_name}")
                    request_parameters[request_name] = task_parameters[task_name]
                if config.pagination:
                    request_parameters[config.pagination.page_param] = page
                    request_parameters[config.pagination.size_param] = config.pagination.size
                requested_at = time.perf_counter()
                response, retries = _request(client, config.method, request_url, request_parameters, connection.max_response_bytes - total_response_bytes, sni_hostname)
                elapsed_ms = int((time.perf_counter() - requested_at) * 1000)
                metrics.increment("standard_http_requests_total", status=response.status_code, connection_id=config.connection_id)
                metrics.observe("standard_http_duration_seconds", elapsed_ms / 1000, connection_id=config.connection_id)
                metrics.increment("standard_http_retries_total", amount=retries, connection_id=config.connection_id)
                content_type = response.headers.get("content-type", "").split(";", 1)[0].strip().lower()
                if content_type not in {"application/json", "application/problem+json"}:
                    raise ValueError("STANDARD_SOURCE_FAILED: response is not JSON")
                total_response_bytes += len(response.content)
                metrics.increment("standard_http_response_bytes_total", amount=len(response.content), connection_id=config.connection_id)
                if len(response.content) > connection.max_response_bytes or total_response_bytes > connection.max_response_bytes:
                    raise ValueError("STANDARD_SOURCE_FAILED: response exceeds configured size")
                payload = _strict_json_loads(response.content)
                page_records = _json_path(payload, config.data_json_path)
                if isinstance(page_records, dict) and not config.pagination and all(isinstance(items, list) and all(isinstance(item, dict) for item in items) for items in page_records.values()):
                    workbook_records = {str(sheet): items for sheet, items in page_records.items()}
                    record_count = sum(len(items) for items in workbook_records.values())
                    page_metadata.append({"page": page, "status_code": response.status_code, "record_count": record_count, "response_bytes": len(response.content), "elapsed_ms": elapsed_ms, "retries": retries})
                    if record_count > connection.max_records:
                        raise ValueError("STANDARD_SOURCE_FAILED: record limit exceeded")
                    break
                if not isinstance(page_records, list) or not all(isinstance(item, dict) for item in page_records):
                    raise ValueError("STANDARD_SOURCE_FAILED: JSONPath did not resolve to an array of objects")
                records.extend(page_records)
                page_metadata.append({"page": page, "status_code": response.status_code, "record_count": len(page_records), "response_bytes": len(response.content), "elapsed_ms": elapsed_ms, "retries": retries})
                if len(records) > connection.max_records:
                    raise ValueError("STANDARD_SOURCE_FAILED: record limit exceeded")
                if not config.pagination:
                    break
                total = _json_path(payload, config.pagination.total_json_path) if config.pagination.total_json_path else None
                if (isinstance(total, int) and len(records) >= total) or len(page_records) < config.pagination.size:
                    break
                page += 1
                if page > config.pagination.max_pages:
                    raise ValueError("STANDARD_SOURCE_FAILED: pagination limit exceeded")
            return workbook_records if workbook_records is not None else records.detach(), {
                "connection_id": config.connection_id,
                "request_path": config.path,
                "method": config.method,
                "pages": page_metadata,
                "response_bytes": total_response_bytes,
                "record_storage": "disk_spill" if isinstance(records.records, SpilledRecords) else "memory",
            }


def _load_secret(reference: str) -> str:
    environment_name = "EXCEL_AUDITOR_SECRET_" + reference.upper().replace("-", "_").replace(".", "_")
    secret = os.environ.get(environment_name)
    if secret is None:
        directory = os.environ.get("EXCEL_AUDITOR_SECRET_DIR")
        if directory:
            root = Path(directory).resolve(strict=True)
            candidate = (root / reference).resolve(strict=True)
            try:
                candidate.relative_to(root)
            except ValueError as exc:
                raise ValueError("secret reference escapes the configured secret directory") from exc
            if not candidate.is_file() or candidate.stat().st_size > 64 * 1024:
                raise ValueError(f"secret reference is not a small regular file: {reference}")
            secret = candidate.read_text(encoding="utf-8").strip()
    if not secret:
        raise ValueError(f"secret reference is not configured: {reference}")
    if "\r" in secret or "\n" in secret:
        raise ValueError(f"secret reference contains an unsafe line break: {reference}")
    return secret


def _strict_json_loads(content: bytes) -> Any:
    try:
        return load_json_strict(content, context="STANDARD_SOURCE_FAILED: response JSON")
    except ValueError as exc:
        if "duplicate key" in str(exc):
            raise ValueError("STANDARD_SOURCE_FAILED: response JSON has a duplicate object key") from exc
        if "non-finite number" in str(exc):
            raise ValueError(str(exc)) from exc
        raise ValueError("STANDARD_SOURCE_FAILED: response JSON is malformed") from exc


def _request(client: httpx.Client, method: str, url: str, parameters: dict[str, Any], max_bytes: int, sni_hostname: str | None = None) -> tuple[httpx.Response, int]:
    last_error: Exception | None = None
    attempts = 3 if method == "GET" else 1
    for attempt in range(attempts):
        try:
            request = client.build_request(method, url, params=parameters if method == "GET" else None, json=parameters if method == "POST" else None)
            if sni_hostname:
                request.extensions["sni_hostname"] = sni_hostname
            streamed = client.send(request, stream=True)
            try:
                streamed.raise_for_status()
                content = bytearray()
                for chunk in streamed.iter_bytes():
                    content.extend(chunk)
                    if len(content) > max_bytes:
                        raise ValueError("STANDARD_SOURCE_FAILED: response exceeds configured size")
                response = httpx.Response(streamed.status_code, headers=streamed.headers, content=bytes(content), request=request)
                return response, attempt
            finally:
                streamed.close()
        except (httpx.TimeoutException, httpx.NetworkError, httpx.HTTPStatusError) as exc:
            last_error = exc
            retryable_status = not isinstance(exc, httpx.HTTPStatusError) or exc.response.status_code == 429 or exc.response.status_code >= 500
            if attempt + 1 < attempts and retryable_status:
                time.sleep(0.05 * (2**attempt))
            else:
                break
    raise ValueError("STANDARD_SOURCE_FAILED: request failed") from last_error


def _validate_target(connection: ManagedConnection, path: str, resolver: Callable[[str], list[str]]) -> list[str]:
    validate_managed_http_path(path, field_name="standard source path")
    if not any(path == allowed or path.startswith(allowed.rstrip("/") + "/") for allowed in connection.allowed_paths):
        raise ValueError("STANDARD_SOURCE_FAILED: path is not allowed")
    parsed = urlparse(connection.base_url)
    if parsed.scheme == "http" and not connection.allow_insecure_http:
        raise ValueError("STANDARD_SOURCE_FAILED: HTTPS is required")
    addresses = resolver(parsed.hostname or "")
    if not addresses:
        raise ValueError("STANDARD_SOURCE_FAILED: host did not resolve")
    if not connection.allow_private_network:
        for address in addresses:
            ip = ipaddress.ip_address(address)
            if not ip.is_global:
                raise ValueError("STANDARD_SOURCE_FAILED: private, loopback, link-local, and metadata targets are blocked")
    return addresses


def _resolve_addresses(hostname: str) -> list[str]:
    return sorted({item[4][0] for item in socket.getaddrinfo(hostname, None, type=socket.SOCK_STREAM)})


def _json_path(payload: Any, path: str | None) -> Any:
    if path is None:
        return None
    if path == "$":
        return payload
    if not path.startswith("$."):
        raise ValueError("only simple object JSONPath expressions are supported")
    current = payload
    for segment in path[2:].split("."):
        if not isinstance(current, dict) or segment not in current:
            raise ValueError(f"JSONPath segment not found: {segment}")
        current = current[segment]
    return current
