from __future__ import annotations

import os
import json
import hashlib
import uuid
import tempfile
import yaml
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse, RedirectResponse
from fastapi.exceptions import RequestValidationError
from pydantic import BaseModel, ValidationError

from .models import DifferenceType, RuleSet
from .engine import compare_workbook
from .workbook import inspect_workbook
from .rules import DraftRegistry, RuleRegistry
from .service import AuditService
from .standard_sources import ConnectionRegistry, ManagedHttpSource
from .persistence import DatabaseRepository
from .queueing import RedisJobQueue
from .storage import S3ArtifactStore
from .rendering import DotNetOpenXmlRenderer
from .observability import configure_logging, metrics


DATA_ROOT = Path(os.environ.get("EXCEL_AUDITOR_DATA", "var")).resolve()
configure_logging()
connections_path = os.environ.get("EXCEL_AUDITOR_CONNECTIONS")
managed_http = ManagedHttpSource(ConnectionRegistry(Path(connections_path))) if connections_path else None
database_url = os.environ.get("DATABASE_URL")
database = DatabaseRepository(database_url, create_schema=os.environ.get("EXCEL_AUDITOR_AUTO_CREATE_SCHEMA") == "1") if database_url else None
s3_bucket = os.environ.get("S3_BUCKET")
artifact_store = S3ArtifactStore(s3_bucket, os.environ.get("S3_ENDPOINT_URL"), os.environ.get("AWS_REGION"), os.environ.get("S3_SERVER_SIDE_ENCRYPTION", "AES256")) if s3_bucket else None
service = AuditService(DATA_ROOT, managed_http=managed_http, database=database, artifact_store=artifact_store)
redis_url = os.environ.get("REDIS_URL")
task_queue = RedisJobQueue(redis_url) if redis_url else None
app = FastAPI(title="Excel Standard Auditor", version="0.1.0")
api_token = os.environ.get("EXCEL_AUDITOR_API_TOKEN")
require_auth = os.environ.get("EXCEL_AUDITOR_REQUIRE_AUTH") == "1"
token_registry: dict[str, dict[str, str]] = {}
if os.environ.get("EXCEL_AUDITOR_API_TOKENS_JSON"):
    raw_registry = json.loads(os.environ["EXCEL_AUDITOR_API_TOKENS_JSON"])
    if not isinstance(raw_registry, dict):
        raise RuntimeError("EXCEL_AUDITOR_API_TOKENS_JSON must be an object")
    for token, identity in raw_registry.items():
        if isinstance(identity, str):
            token_registry[str(token)] = {"tenant_id": identity, "user_id": "service"}
        elif isinstance(identity, dict) and identity.get("tenant_id") and identity.get("user_id"):
            token_registry[str(token)] = {"tenant_id": str(identity["tenant_id"]), "user_id": str(identity["user_id"])}
        else:
            raise RuntimeError("each API token identity requires tenant_id and user_id")
if api_token:
    token_registry[api_token] = {"tenant_id": "default", "user_id": "service"}
if require_auth and not token_registry:
    raise RuntimeError("production authentication is required but no API token is configured")


@app.middleware("http")
async def bearer_auth(request: Request, call_next):
    trace_id = request.headers.get("x-request-id") or uuid.uuid4().hex
    request.state.trace_id = trace_id[:128]
    request.state.tenant_id = "local"
    request.state.user_id = "local"
    if token_registry and not request.url.path.startswith("/health/"):
        authorization = request.headers.get("authorization", "")
        token = authorization.removeprefix("Bearer ") if authorization.startswith("Bearer ") else ""
        identity = token_registry.get(token)
        if identity is None:
            return JSONResponse(status_code=401, media_type="application/problem+json", content={"type": "about:blank", "title": "Unauthorized", "status": 401, "detail": "valid bearer token required"})
        request.state.tenant_id = identity["tenant_id"]
        request.state.user_id = identity["user_id"]
    response = await call_next(request)
    response.headers["X-Request-ID"] = request.state.trace_id
    return response


@app.get("/health/live")
def live() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/health/ready")
def ready():
    checks: dict[str, str] = {}
    try:
        DATA_ROOT.mkdir(parents=True, exist_ok=True)
        checks["filesystem"] = "ok"
        if database:
            database.ping()
            checks["database"] = "ok"
        if task_queue:
            task_queue.ping()
            checks["redis"] = "ok"
        if artifact_store:
            artifact_store.ping()
            checks["object_storage"] = "ok"
        if isinstance(service.renderer, DotNetOpenXmlRenderer):
            checks["renderer"] = f"ExcelRenderer/{service.renderer.self_check()}"
        else:
            checks["renderer"] = service.renderer.__class__.__name__
    except Exception:
        metrics.increment("readiness_failures_total")
        return JSONResponse(status_code=503, content={"status": "not_ready", "checks": checks})
    return {"status": "ready", "checks": checks}


@app.get("/metrics", response_class=PlainTextResponse)
def prometheus_metrics() -> str:
    return metrics.prometheus()


@app.post("/api/v1/workbooks/precheck")
async def precheck_workbook(
    request: Request,
    excel_file: UploadFile = File(...),
    schema_id: str = Form(...),
    schema_version: str = Form(...),
) -> dict:
    try:
        rules = _get_rule(request, schema_id, schema_version)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc)) from exc
    excel_limit = rules.workbook.max_upload_mib * 1024 * 1024
    content = await excel_file.read(excel_limit + 1)
    if len(content) > excel_limit:
        raise HTTPException(413, "FILE_LIMIT_EXCEEDED")
    suffix = Path(excel_file.filename or "input.xlsx").suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        raise HTTPException(415, "FILE_UNSUPPORTED_FORMAT")
    precheck_root = DATA_ROOT / "prechecks"
    precheck_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="precheck_", dir=precheck_root) as temporary:
        path = Path(temporary) / f"input{suffix}"
        path.write_bytes(content)
        workbook = inspect_workbook(path, rules)
        try:
            comparison = compare_workbook(workbook, {sheet.id: [] for sheet in rules.sheets}, rules)
            header_types = {
                DifferenceType.MISSING_SHEET, DifferenceType.EXTRA_SHEET, DifferenceType.HEADER_NOT_FOUND,
                DifferenceType.MISSING_HEADER, DifferenceType.EXTRA_HEADER, DifferenceType.DUPLICATE_HEADER,
                DifferenceType.AMBIGUOUS_HEADER, DifferenceType.HEADER_ORDER_MISMATCH,
            }
            result = {
                "input_sha256": workbook.sha256,
                "structure": [{"sheet_name": sheet.name, "rows": sheet.max_row, "columns": sheet.max_column, "hidden_rows": len(sheet.hidden_rows), "features": sheet.risky_features} for sheet in workbook.sheets.values()],
                "header_mappings": [item.model_dump(mode="json") for item in comparison.mappings],
                "header_differences": [item.model_dump(mode="json") for item in comparison.differences if item.type in header_types],
                "warnings": workbook.warnings,
                "manual_review_reasons": [*workbook.manual_review_reasons, *(comparison.manual_review_reasons or [])],
                "auto_repair_authorizations": {sheet.id: sheet.actions.model_dump(mode="json") for sheet in rules.sheets},
            }
        finally:
            workbook.close()
    return result


@app.post("/api/v1/schemas/publish", status_code=201)
def publish_schema(request: Request, rules: RuleSet) -> dict[str, str]:
    try:
        _validate_connection_reference(rules)
        _rule_registry(request).publish(rules)
        if database:
            database.publish_rule(rules, request.state.user_id, request.state.tenant_id)
    except FileExistsError as exc:
        raise HTTPException(409, str(exc)) from exc
    return {"schema_id": rules.schema_id, "schema_version": rules.schema_version, "config_sha256": rules.content_sha256, "status": "published"}


@app.post("/api/v1/schemas/import", status_code=201)
async def import_schema(request: Request, file: UploadFile = File(...)) -> dict:
    content = await file.read(1024 * 1024 + 1)
    if len(content) > 1024 * 1024:
        raise HTTPException(413, "RULE_CONFIG_INVALID: import exceeds 1 MiB")
    suffix = Path(file.filename or "rules.json").suffix.lower()
    try:
        if suffix in {".yaml", ".yml"}:
            payload = yaml.safe_load(content.decode("utf-8-sig"))
        elif suffix == ".json":
            payload = json.loads(content.decode("utf-8-sig"))
        else:
            raise HTTPException(415, "RULE_CONFIG_INVALID: only JSON and YAML are supported")
        rules = RuleSet.model_validate(payload)
    except (UnicodeDecodeError, json.JSONDecodeError, yaml.YAMLError, ValidationError) as exc:
        raise HTTPException(422, "RULE_CONFIG_INVALID: uploaded rule document is invalid") from exc
    return _draft_registry(request).create(rules.schema_id, rules.model_dump(mode="json"))


class DraftCreate(BaseModel):
    base_version: str | None = None
    config: dict | None = None


class MappingConfirmation(BaseModel):
    schema_id: str
    draft_id: str
    raw_header: str
    canonical_field: str
    sheet_id: str | None = None


@app.post("/api/v1/schemas", status_code=201)
def create_schema(request: Request, rules: RuleSet) -> dict:
    return _draft_registry(request).create(rules.schema_id, rules.model_dump(mode="json"))


@app.post("/api/v1/schemas/{schema_id}/drafts", status_code=201)
def create_draft(schema_id: str, payload: DraftCreate, request: Request) -> dict:
    if payload.config is not None:
        config = payload.config
    elif payload.base_version:
        try:
            base = _get_rule(request, schema_id, payload.base_version)
        except FileNotFoundError as exc:
            raise HTTPException(404, str(exc)) from exc
        config = base.model_dump(mode="json")
    else:
        raise HTTPException(422, "config or base_version is required")
    if config.get("schema_id") != schema_id:
        raise HTTPException(422, "draft schema_id does not match path")
    return _draft_registry(request).create(schema_id, config, payload.base_version)


@app.put("/api/v1/schemas/{schema_id}/drafts/{draft_id}")
def update_draft(schema_id: str, draft_id: str, config: dict, request: Request) -> dict:
    try:
        return _draft_registry(request).update(schema_id, draft_id, config)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(409 if "immutable" in str(exc) else 422, str(exc)) from exc


@app.post("/api/v1/schemas/{schema_id}/drafts/{draft_id}/validate")
def validate_draft(schema_id: str, draft_id: str, request: Request) -> dict:
    try:
        rules = _draft_registry(request).validate(schema_id, draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except ValidationError as exc:
        return {"valid": False, "errors": exc.errors(include_url=False)}
    return {"valid": True, "schema_id": rules.schema_id, "schema_version": rules.schema_version, "config_sha256": rules.content_sha256, "errors": []}


@app.post("/api/v1/schemas/{schema_id}/drafts/{draft_id}/publish", status_code=201)
def publish_draft(schema_id: str, draft_id: str, request: Request) -> dict:
    try:
        tenant_drafts = _draft_registry(request)
        rules = tenant_drafts.validate(schema_id, draft_id)
        _validate_connection_reference(rules)
        _rule_registry(request).publish(rules)
        if database:
            database.publish_rule(rules, request.state.user_id, request.state.tenant_id)
        tenant_drafts.mark_published(schema_id, draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except FileExistsError as exc:
        raise HTTPException(409, str(exc)) from exc
    return {"schema_id": schema_id, "schema_version": rules.schema_version, "config_sha256": rules.content_sha256, "status": "published"}


@app.get("/api/v1/schemas/{schema_id}/versions")
def list_schema_versions(schema_id: str, request: Request) -> dict:
    items = database.list_rule_versions(schema_id, request.state.tenant_id) if database else _rule_registry(request).versions(schema_id)
    return {"items": items, "total": len(items)}


@app.post("/api/v1/mappings/confirm")
def confirm_mapping(payload: MappingConfirmation, request: Request) -> dict:
    try:
        record = _draft_registry(request).confirm_mapping(payload.schema_id, payload.draft_id, payload.raw_header, payload.canonical_field, payload.sheet_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except (ValueError, ValidationError) as exc:
        raise HTTPException(422, str(exc)) from exc
    if database:
        database.audit("mapping.confirmed", "schema_draft", payload.draft_id, request.state.user_id, {"raw_header": payload.raw_header, "canonical_field": payload.canonical_field, "sheet_id": payload.sheet_id}, request.state.tenant_id)
    return {"draft_id": payload.draft_id, "status": record["status"], "raw_header": payload.raw_header, "canonical_field": payload.canonical_field}


@app.get("/api/v1/schemas/{schema_id}/versions/{version}")
def get_schema(schema_id: str, version: str, request: Request) -> RuleSet:
    try:
        return _get_rule(request, schema_id, version)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc)) from exc


@app.get("/api/v1/schemas/{schema_id}/versions/{version}/export")
def export_schema(schema_id: str, version: str, request: Request, format: str = Query(default="json", pattern="^(json|yaml)$")):
    try:
        rules = _get_rule(request, schema_id, version)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc)) from exc
    payload = rules.model_dump(mode="json")
    if format == "yaml":
        return PlainTextResponse(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), media_type="application/yaml", headers={"Content-Disposition": f'attachment; filename="{schema_id}-{version}.yaml"'})
    return JSONResponse(payload, headers={"Content-Disposition": f'attachment; filename="{schema_id}-{version}.json"'})


@app.post("/api/v1/comparisons", status_code=202)
async def create_comparison(
    request: Request,
    background_tasks: BackgroundTasks,
    excel_file: UploadFile = File(...),
    standard_data: UploadFile | None = File(None),
    standard_json: str | None = Form(None),
    schema_id: str = Form(...),
    schema_version: str = Form(...),
    parameters: str = Form("{}"),
    idempotency_key: str | None = Header(default=None, alias="Idempotency-Key"),
):
    try:
        rules = _get_rule(request, schema_id, schema_version)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc)) from exc
    excel_limit = rules.workbook.max_upload_mib * 1024 * 1024
    excel_content = await excel_file.read(excel_limit + 1)
    if len(excel_content) > excel_limit:
        raise HTTPException(413, "FILE_LIMIT_EXCEEDED")
    try:
        parsed_parameters = json.loads(parameters)
        if not isinstance(parsed_parameters, dict):
            raise ValueError
    except (json.JSONDecodeError, ValueError) as exc:
        raise HTTPException(422, "parameters must be a JSON object") from exc
    excel_suffix = Path(excel_file.filename or "input.xlsx").suffix.lower()
    if excel_suffix not in {".xlsx", ".xlsm"}:
        raise HTTPException(415, "FILE_UNSUPPORTED_FORMAT")
    standard_content: bytes | None = None
    standard_suffix: str | None = None
    if standard_data is not None and standard_json is not None:
        raise HTTPException(422, "standard_data and standard_json are mutually exclusive")
    if standard_data is not None:
        standard_limit = rules.workbook.max_standard_upload_mib * 1024 * 1024
        standard_content = await standard_data.read(standard_limit + 1)
        if len(standard_content) > standard_limit:
            raise HTTPException(413, "STANDARD_DATA_INVALID: payload too large")
        standard_suffix = Path(standard_data.filename or "standard.json").suffix.lower()
        if standard_suffix not in {".json", ".csv"}:
            raise HTTPException(415, "STANDARD_DATA_INVALID: only JSON and CSV uploads are supported")
    elif standard_json is not None:
        encoded = standard_json.encode("utf-8")
        if len(encoded) > rules.workbook.max_standard_upload_mib * 1024 * 1024:
            raise HTTPException(413, "STANDARD_DATA_INVALID: payload too large")
        try:
            inline_payload = json.loads(standard_json)
        except json.JSONDecodeError as exc:
            raise HTTPException(422, "STANDARD_DATA_INVALID: standard_json is not valid JSON") from exc
        if not isinstance(inline_payload, (dict, list)):
            raise HTTPException(422, "STANDARD_DATA_INVALID: standard_json root must be an object or array")
        standard_content = json.dumps(inline_payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
        standard_suffix = ".json"
    elif rules.standard_source.type != "managed_http":
        raise HTTPException(422, "standard_data is required for upload-based rules")
    fingerprint = hashlib.sha256()
    fingerprint.update(rules.content_sha256.encode("ascii"))
    fingerprint.update(json.dumps(parsed_parameters, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8"))
    fingerprint.update(excel_suffix.encode("ascii"))
    fingerprint.update(excel_content)
    if standard_content is not None:
        fingerprint.update((standard_suffix or "").encode("ascii"))
        fingerprint.update(standard_content)
    try:
        job_id, replayed = service.create_or_get_job(idempotency_key, fingerprint.hexdigest(), request.state.tenant_id, request.state.user_id, request.state.trace_id)
    except FileExistsError as exc:
        raise HTTPException(409, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(429 if "TENANT_QUOTA_EXCEEDED" in str(exc) else 422, str(exc)) from exc
    if replayed:
        current = service.status(job_id)
        return JSONResponse(status_code=200, content={"job_id": job_id, "status": current["status"], "schema_id": schema_id, "schema_version": schema_version, "created_at": current["created_at"], "idempotent_replay": True})
    directory = service.job_directory(job_id)
    service.record_input_metadata(job_id, excel_file.filename or f"upload{excel_suffix}", len(excel_content))
    excel_path = directory / f"upload{excel_suffix}"
    standard_path = None
    if standard_content is not None and standard_suffix is not None:
        standard_path = directory / f"standard{standard_suffix}"
        standard_path.write_bytes(standard_content)
    excel_path.write_bytes(excel_content)
    if task_queue:
        try:
            task_queue.enqueue(DATA_ROOT, job_id, excel_path, standard_path, rules, parsed_parameters)
        except Exception as exc:
            raise HTTPException(503, "task queue is unavailable") from exc
    else:
        background_tasks.add_task(service.run, job_id, excel_path, standard_path, rules, parsed_parameters)
    return {"job_id": job_id, "status": "queued", "schema_id": schema_id, "schema_version": schema_version, "created_at": service.status(job_id)["created_at"]}


@app.post("/api/v1/comparisons/{job_id}/cancel", status_code=202)
def cancel_comparison(job_id: str, request: Request) -> dict:
    try:
        _authorize_job(request, job_id)
        status = service.request_cancel(job_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(409, str(exc)) from exc
    return {"job_id": job_id, "status": status["status"], "cancel_requested": True}


@app.delete("/api/v1/comparisons/{job_id}", status_code=202)
def delete_comparison(job_id: str, request: Request) -> dict:
    _authorize_job(request, job_id)
    try:
        status = service.soft_delete(job_id, request.state.user_id)
    except ValueError as exc:
        raise HTTPException(409, str(exc)) from exc
    return {"job_id": job_id, "deleted_at": status["deleted_at"], "purge_after": status["purge_after"]}


@app.get("/api/v1/comparisons/{job_id}")
def comparison_status(job_id: str, request: Request) -> dict:
    try:
        _authorize_job(request, job_id)
        return service.status(job_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc


@app.get("/api/v1/comparisons/{job_id}/differences")
def comparison_differences(
    job_id: str,
    request: Request,
    type: str | None = None,
    sheet_id: str | None = None,
    canonical_field: str | None = None,
    severity: str | None = None,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=50, ge=1, le=200),
) -> dict:
    try:
        _authorize_job(request, job_id)
        differences_path = service.artifact(job_id, "differences_jsonl")
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    start = (page - 1) * page_size
    stop = start + page_size
    total = 0
    items: list[dict] = []
    with differences_path.open("r", encoding="utf-8") as handle:
        for line in handle:
            if not line.strip():
                continue
            item = json.loads(line)
            if type and item["type"] != type:
                continue
            if sheet_id and item["sheet_id"] != sheet_id:
                continue
            if canonical_field and item.get("canonical_field") != canonical_field:
                continue
            if severity and item.get("severity") != severity:
                continue
            if start <= total < stop:
                items.append(item)
            total += 1
    return {"items": items, "page": page, "page_size": page_size, "total": total}


@app.get("/api/v1/comparisons/{job_id}/artifacts/{artifact}")
def get_artifact(job_id: str, artifact: str, request: Request):
    _authorize_job(request, job_id)
    if artifact_store:
        try:
            status = service.status(job_id)
            file_name = status.get("artifacts", {}).get(artifact)
            key = status.get("object_keys", {}).get(file_name)
            if key:
                if database:
                    database.audit("comparison.artifact_downloaded", "comparison_job", job_id, request.state.user_id, {"artifact": artifact}, request.state.tenant_id)
                return RedirectResponse(artifact_store.download_url(key), status_code=307)
        except FileNotFoundError as exc:
            raise HTTPException(404, str(exc)) from exc
    try:
        path = service.artifact(job_id, artifact)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    excel_media = "application/vnd.ms-excel.sheet.macroEnabled.12" if path.suffix.lower() == ".xlsm" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    media = {"excel": excel_media, "json": "application/json", "differences_jsonl": "application/x-ndjson", "html": "text/html", "manifest": "application/json"}[artifact]
    if database:
        database.audit("comparison.artifact_downloaded", "comparison_job", job_id, request.state.user_id, {"artifact": artifact}, request.state.tenant_id)
    return FileResponse(path, media_type=media, filename=path.name)


@app.exception_handler(HTTPException)
async def http_error(request: Request, exc: HTTPException) -> JSONResponse:
    return JSONResponse(status_code=exc.status_code, media_type="application/problem+json", content={"type": "about:blank", "title": _problem_title(exc.status_code), "status": exc.status_code, "detail": str(exc.detail), "instance": request.url.path})


@app.exception_handler(RequestValidationError)
async def request_validation_error(request: Request, exc: RequestValidationError) -> JSONResponse:
    return JSONResponse(status_code=422, media_type="application/problem+json", content={
        "type": "about:blank",
        "title": "Request validation failed",
        "status": 422,
        "detail": "Request body or parameters did not satisfy the API schema",
        "instance": request.url.path,
        "errors": _safe_validation_issues(exc.errors()),
    })


@app.exception_handler(ValidationError)
async def validation_error(request: Request, exc: ValidationError) -> JSONResponse:
    return JSONResponse(status_code=422, media_type="application/problem+json", content={
        "type": "about:blank",
        "title": "RULE_CONFIG_INVALID",
        "status": 422,
        "detail": "Rule configuration did not satisfy the published schema",
        "instance": request.url.path,
        "errors": _safe_validation_issues(exc.errors()),
    })


def _safe_validation_issues(errors: list[dict]) -> list[dict]:
    """Return actionable field locations without echoing untrusted input/context."""
    return [
        {
            "loc": [component if isinstance(component, int) else str(component)[:128] for component in error.get("loc", ())],
            "type": str(error.get("type", "validation_error"))[:128],
        }
        for error in errors[:100]
    ]


def _problem_title(status: int) -> str:
    return {401: "Unauthorized", 403: "Forbidden", 404: "Not Found", 409: "Conflict", 413: "Payload Too Large", 415: "Unsupported Media Type", 422: "Unprocessable Content", 429: "Too Many Requests", 503: "Service Unavailable"}.get(status, "Request failed")


def _authorize_job(request: Request, job_id: str) -> dict:
    try:
        status = service.status(job_id)
    except FileNotFoundError as exc:
        raise HTTPException(404, str(exc)) from exc
    if status.get("tenant_id", "local") != request.state.tenant_id:
        # Do not reveal whether a cross-tenant job exists.
        raise HTTPException(404, "job not found")
    if status.get("deleted_at"):
        raise HTTPException(404, "job not found")
    return status


def _tenant_registry_root(request: Request) -> Path:
    # Tenant identifiers originate in the trusted token registry, but hashing also
    # keeps them out of filesystem paths and prevents traversal by construction.
    tenant_key = hashlib.sha256(request.state.tenant_id.encode("utf-8")).hexdigest()
    return DATA_ROOT / "tenants" / tenant_key


def _rule_registry(request: Request) -> RuleRegistry:
    return RuleRegistry(_tenant_registry_root(request) / "schemas")


def _draft_registry(request: Request) -> DraftRegistry:
    return DraftRegistry(_tenant_registry_root(request) / "drafts")


def _get_rule(request: Request, schema_id: str, version: str) -> RuleSet:
    if database:
        return database.get_rule(schema_id, version, request.state.tenant_id)
    return _rule_registry(request).get(schema_id, version)


def _validate_connection_reference(rules: RuleSet) -> None:
    source = rules.standard_source
    if source.type != "managed_http":
        return
    if managed_http is None:
        raise HTTPException(422, "RULE_CONFIG_INVALID: managed connection registry is unavailable")
    try:
        managed_http.registry.get(source.connection_id or "")
    except KeyError as exc:
        raise HTTPException(422, f"RULE_CONFIG_INVALID: {exc}") from exc


def run() -> None:
    import uvicorn
    from .security_config import validate_production_environment

    validate_production_environment()
    uvicorn.run("excel_auditor.api:app", host=os.environ.get("EXCEL_AUDITOR_HOST", "127.0.0.1"), port=int(os.environ.get("EXCEL_AUDITOR_PORT", "8000")))
