from __future__ import annotations

import json
from contextlib import contextmanager
import hashlib
import os
import shutil
import sqlite3
import stat
import tempfile
import time
from collections.abc import Sequence as SequenceABC
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Sequence

from .engine import compare_workbook
from .models import AuditReport, RuleSet
from .rendering import DotNetOpenXmlRenderer, ExcelRenderer, OpenPyxlDevelopmentRenderer, _repair_provenance_comment
from .reporting import write_differences_jsonl, write_html_report, write_json_report
from .snapshots import SpilledRecords, create_snapshot, load_snapshot
from .standard_sources import ManagedHttpSource
from .standard_files import canonicalize_standard_row, load_standard_file
from .persistence import DatabaseRepository
from .storage import ArtifactStore
from .ids import new_ulid
from .pandera_adapter import StandardDataValidator
from .workbook import inspect_workbook
from .observability import log_event, metrics


class AuditService:
    def __init__(self, root: Path, renderer: ExcelRenderer | None = None, managed_http: ManagedHttpSource | None = None, database: DatabaseRepository | None = None, artifact_store: ArtifactStore | None = None) -> None:
        self.root = root
        self.jobs = root / "jobs"
        self.jobs.mkdir(parents=True, exist_ok=True)
        renderer_command = os.environ.get("EXCEL_RENDERER_COMMAND")
        self.renderer = renderer or (DotNetOpenXmlRenderer(Path(renderer_command)) if renderer_command else OpenPyxlDevelopmentRenderer())
        self.managed_http = managed_http
        self.database = database
        self.artifact_store = artifact_store
        self.standard_validator = StandardDataValidator()

    def create_job(self, job_id: str | None = None, tenant_id: str = "local", user_id: str = "local", trace_id: str | None = None) -> str:
        job_id = job_id or new_ulid("job_")
        directory = self.jobs / job_id
        active_limit = int(os.environ.get("EXCEL_AUDITOR_MAX_ACTIVE_JOBS_PER_TENANT", "10"))
        with _tenant_quota_lock(self.root, tenant_id):
            active = 0
            for status_path in self.jobs.glob("job_*/status.json"):
                try:
                    existing = json.loads(status_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    continue
                if existing.get("tenant_id") == tenant_id and existing.get("status") not in {"completed", "failed", "cancelled", "manual_review"}:
                    active += 1
            if active >= active_limit:
                raise ValueError("TENANT_QUOTA_EXCEEDED: active comparison limit reached")
            directory.mkdir(parents=True, exist_ok=False)
            initial_status = {"job_id": job_id, "tenant_id": tenant_id, "user_id": user_id, "trace_id": trace_id, "status": "queued", "progress": 0, "created_at": _now(), "retention_until": (datetime.now(timezone.utc) + timedelta(days=_retention_days(tenant_id))).isoformat()}
            self._write_status(job_id, initial_status)
        if self.database:
            self.database.create_job(job_id, tenant_id, user_id)
            self.database.update_job(initial_status)
        metrics.increment("jobs_created_total")
        log_event(job_id=job_id, trace_id=trace_id, stage="queued", event="comparison.created")
        return job_id

    def create_or_get_job(self, idempotency_key: str | None, fingerprint: str, tenant_id: str = "local", user_id: str = "local", trace_id: str | None = None) -> tuple[str, bool]:
        if not idempotency_key:
            return self.create_job(tenant_id=tenant_id, user_id=user_id, trace_id=trace_id), False
        if len(idempotency_key) > 200 or not idempotency_key.strip():
            raise ValueError("Idempotency-Key must contain 1 to 200 characters")
        registry = self.root / "idempotency"
        registry.mkdir(parents=True, exist_ok=True)
        key_hash = hashlib.sha256(f"{tenant_id}\0{idempotency_key}".encode("utf-8")).hexdigest()
        path = registry / f"{key_hash}.json"
        job_id = new_ulid("job_")
        record = json.dumps({"fingerprint": fingerprint, "job_id": job_id, "created_at": _now()}, ensure_ascii=False).encode("utf-8")
        try:
            descriptor = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
        except FileExistsError:
            existing = json.loads(path.read_text(encoding="utf-8"))
            if existing.get("fingerprint") != fingerprint:
                raise FileExistsError("Idempotency-Key was already used with a different request")
            existing_job = str(existing["job_id"])
            if not (self.jobs / existing_job).is_dir():
                self.create_job(existing_job, tenant_id, user_id, trace_id)
            return existing_job, True
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(record)
            handle.flush()
            os.fsync(handle.fileno())
        self.create_job(job_id, tenant_id, user_id, trace_id)
        return job_id, False

    def request_cancel(self, job_id: str) -> dict[str, Any]:
        status = self.status(job_id)
        if status["status"] in {"completed", "failed", "cancelled", "manual_review"}:
            raise ValueError(f"job is already terminal: {status['status']}")
        marker = self.job_directory(job_id) / "cancel.requested"
        marker.write_text(_now(), encoding="utf-8")
        updated = {**status, "cancel_requested": True}
        self._write_status(job_id, updated)
        if self.database:
            self.database.audit("comparison.cancel_requested", "comparison_job", job_id, status.get("user_id"), tenant_id=status.get("tenant_id"))
        return updated

    def record_input_metadata(self, job_id: str, file_name: str, file_size: int) -> None:
        status = self.status(job_id)
        safe_name = Path(file_name).name[:255]
        self._write_status(job_id, {**status, "input_file_name": safe_name, "input_file_size": int(file_size)})

    def soft_delete(self, job_id: str, actor_id: str | None = None) -> dict[str, Any]:
        status = self.status(job_id)
        if status["status"] not in {"completed", "failed", "cancelled", "manual_review"}:
            raise ValueError("only terminal jobs can be deleted")
        delay_days = int(os.environ.get("EXCEL_AUDITOR_DELETE_DELAY_DAYS", "7"))
        deleted_at = datetime.now(timezone.utc)
        updated = {**status, "deleted_at": deleted_at.isoformat(), "purge_after": (deleted_at + timedelta(days=delay_days)).isoformat()}
        self._write_status(job_id, updated)
        if self.database:
            self.database.audit("comparison.soft_deleted", "comparison_job", job_id, actor_id, {"purge_after": updated["purge_after"]}, status.get("tenant_id"))
        return updated

    def purge_expired(self, now: datetime | None = None) -> list[str]:
        now = now or datetime.now(timezone.utc)
        purged: list[str] = []
        for directory in self.jobs.glob("job_*"):
            status_path = directory / "status.json"
            try:
                status = json.loads(status_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if status.get("status") not in {"completed", "failed", "cancelled", "manual_review"}:
                continue
            deadline_text = status.get("purge_after") if status.get("deleted_at") else status.get("retention_until")
            if not deadline_text or datetime.fromisoformat(deadline_text) > now:
                continue
            if self.artifact_store:
                for key in status.get("object_keys", {}).values():
                    self.artifact_store.delete(str(key))
            if self.database:
                self.database.audit("comparison.purged", "comparison_job", status["job_id"], status.get("user_id"), {"deleted_at": status.get("deleted_at")}, status.get("tenant_id"))
                self.database.purge_job(status["job_id"], status.get("tenant_id", "local"))
            _remove_tree(directory)
            purged.append(status["job_id"])
        return purged

    def run(self, job_id: str, excel_path: Path, standard_path: Path | None, rules: RuleSet, parameters: dict[str, Any] | None = None) -> None:
        directory = self.job_directory(job_id)
        identity = self.status(job_id)
        tenant_id = identity.get("tenant_id", "local")
        user_id = identity.get("user_id", "local")
        trace_id = identity.get("trace_id")
        started = time.perf_counter()
        comparison: Any | None = None
        try:
            metrics.observe("queue_wait_seconds", max(0.0, (datetime.now(timezone.utc) - datetime.fromisoformat(identity["created_at"])).total_seconds()))
        except (KeyError, ValueError, TypeError):
            pass
        try:
            self._check_cancelled(job_id)
            self._write_status(job_id, {**self.status(job_id), "status": "validating", "progress": 10, "schema_id": rules.schema_id, "schema_version": rules.schema_version})
            input_suffix = excel_path.suffix.lower()
            input_path = directory / f"input{input_suffix}"
            if excel_path.resolve() != input_path.resolve():
                shutil.copy2(excel_path, input_path)
            # Reject unsafe/unsupported packages before contacting a managed
            # standard source. This keeps attacker-controlled workbooks from
            # triggering outbound requests before the ZIP/OpenXML gate runs.
            stage_started = time.perf_counter()
            workbook = inspect_workbook(input_path, rules)
            metrics.observe("stage_duration_seconds", time.perf_counter() - stage_started, stage="inspection")
            _make_read_only(input_path)
            standard_stage_started = time.perf_counter()
            if standard_path is not None:
                standard = _load_standard(standard_path, rules)
                source_metadata = {"upload_sha256": _sha256(standard_path), "upload_format": standard_path.suffix.lower().lstrip(".")}
            elif rules.standard_source.type == "managed_http" and self.managed_http is not None:
                fetched, source_metadata = self.managed_http.fetch_with_metadata(rules.standard_source, parameters)
                if isinstance(fetched, dict):
                    standard = _canonicalize_standard(fetched, rules)
                elif len(rules.sheets) == 1:
                    standard = _canonicalize_standard({rules.sheets[0].id: fetched}, rules)
                else:
                    raise ValueError("STANDARD_DATA_INVALID: multi-sheet managed source must return an object keyed by sheet id or name")
            else:
                raise ValueError("STANDARD_DATA_INVALID: standard data upload or configured managed source is required")
            try:
                self.standard_validator.validate(standard, rules)
                metrics.observe("stage_duration_seconds", time.perf_counter() - standard_stage_started, stage="standard_snapshot")
                self._check_cancelled(job_id)
                snapshot = create_snapshot(standard, directory, {
                    **source_metadata,
                    "schema_id": rules.schema_id,
                    "schema_version": rules.schema_version,
                    "schema_sha256": rules.content_sha256,
                })
            finally:
                _close_record_sequences(standard)
            _make_read_only(snapshot.path)
            # All downstream comparison reads the immutable, hash-verified snapshot,
            # never the mutable fetch/upload object used to create it.
            standard_spill_threshold = _standard_spill_threshold(rules)
            standard = load_snapshot(snapshot, spill_after_records=standard_spill_threshold)
            snapshot_key = self.artifact_store.put_file(f"jobs/{job_id}/standard/{snapshot.path.name}", snapshot.path) if self.artifact_store else None
            if self.database:
                self.database.save_snapshot(snapshot, rules.standard_source.type, tenant_id, snapshot_key)
                self.database.audit("standard.snapshot_created", "standard_snapshot", snapshot.snapshot_id, user_id, {"record_count": snapshot.record_count, "content_sha256": snapshot.sha256, **snapshot.metadata}, tenant_id)
            # Production object storage is the durable encrypted-at-rest copy. The
            # plaintext staging file is no longer needed after hash verification.
            if self.artifact_store and snapshot_key:
                snapshot.path.chmod(stat.S_IREAD | stat.S_IWRITE)
                snapshot.path.unlink(missing_ok=True)
            snapshot_objects = dict(self.status(job_id).get("object_keys", {}))
            if snapshot_key:
                snapshot_objects["standard_snapshot"] = snapshot_key
            self._write_status(job_id, {**self.status(job_id), "status": "comparing", "progress": 40, "standard_snapshot_id": snapshot.snapshot_id, "object_keys": snapshot_objects})
            stage_started = time.perf_counter()
            try:
                comparison = compare_workbook(workbook, standard, rules, job_id=job_id, spill_to_disk=True)
                metrics.observe("stage_duration_seconds", time.perf_counter() - stage_started, stage="comparison")
            finally:
                for rows in standard.values():
                    close_rows = getattr(rows, "close", None)
                    if close_rows is not None:
                        close_rows()
                workbook.close()
            row_count = sum(max(0, sheet.max_row - 1) for sheet in workbook.sheets.values())
            metrics.increment("rows_processed_total", amount=row_count)
            metrics.increment("differences_total", amount=len(comparison.differences))
            mapping_review_reasons = [*(comparison.manual_review_reasons or []), *[
                f"{mapping.sheet_id}: ambiguous_header:{mapping.raw_header}"
                for mapping in comparison.mappings
                if mapping.status == "ambiguous"
            ]]
            for difference in comparison.differences:
                difference.job_id = job_id
            self._check_cancelled(job_id)
            report = AuditReport(
                job_id=job_id,
                created_at=datetime.now(timezone.utc),
                schema_id=rules.schema_id,
                schema_version=rules.schema_version,
                schema_sha256=rules.content_sha256,
                input_sha256=workbook.sha256,
                input_file_name=identity.get("input_file_name", excel_path.name),
                input_file_size=identity.get("input_file_size", input_path.stat().st_size),
                standard_snapshot_id=snapshot.snapshot_id,
                standard_sha256=snapshot.sha256,
                standard_source_metadata=snapshot.metadata,
                header_mappings=comparison.mappings,
                differences=comparison.differences,
                summary=comparison.summary,
                warnings=[
                    *workbook.warnings,
                    *mapping_review_reasons,
                    *[f"comparison_backend:{backend}" for backend in sorted(set(comparison.join_backends or []))],
                    *[f"comparison_storage:{backend}" for backend in sorted(set(comparison.storage_backends or []))],
                ],
                workbook_structure=[
                    {"sheet_name": item.name, "rows": item.max_row, "columns": item.max_column, "hidden_rows": len(item.hidden_rows), "features": item.risky_features}
                    for item in workbook.sheets.values()
                ],
                header_summary=dict(Counter(mapping.status for mapping in comparison.mappings)),
                record_summary={
                    "matched": comparison.summary.matched_records,
                    "extra": comparison.summary.extra_records,
                    "missing": comparison.summary.missing_records,
                },
                field_statistics=_field_statistics(comparison),
                data_quality_summary=dict(Counter(
                    item.type.value for item in comparison.differences
                    if item.type.value in {"EMPTY_PRIMARY_KEY", "DUPLICATE_PRIMARY_KEY", "INVALID_VALUE", "VALIDATION_ERROR"}
                )),
            )
            if self.database:
                self.database.save_differences(report)
                if report.summary.repairs_planned:
                    self.database.audit("comparison.auto_repair_planned", "comparison_job", job_id, user_id, {"repair_count": report.summary.repairs_planned, "schema_sha256": rules.content_sha256}, tenant_id)
            write_json_report(report, directory / "report.json")
            write_differences_jsonl(report.differences, directory / "differences.jsonl")
            # These reasons represent structures the current implementation cannot
            # prove safe.  They are never bypassed by the legacy allow/report
            # settings: unsupported work must fail or enter manual review (DoD 20).
            workbook_requires_review = bool(workbook.manual_review_reasons)
            if mapping_review_reasons or workbook_requires_review:
                if workbook.manual_review_reasons and rules.workbook.unsupported_feature_action == "reject":
                    raise ValueError("MANUAL_REVIEW_REQUIRED: unsupported workbook features")
                public_manifest = _non_rendering_manifest(report, "manual_review")
                (directory / "render-manifest.json").write_text(json.dumps(public_manifest, ensure_ascii=False, indent=2), encoding="utf-8")
                write_html_report(report, directory / "report.html")
                object_keys = self._upload_artifacts(job_id, directory, [input_path.name, "report.json", "differences.jsonl", "report.html", "render-manifest.json"])
                self._write_status(job_id, {
                    **self.status(job_id),
                    "status": "manual_review",
                    "progress": 100,
                    "completed_at": _now(),
                    "summary": report.summary.model_dump(mode="json"),
                    "warnings": report.warnings,
                    "artifacts": {"json": "report.json", "differences_jsonl": "differences.jsonl", "html": "report.html", "manifest": "render-manifest.json"},
                    "object_keys": object_keys,
                })
                metrics.increment("jobs_terminal_total", status="manual_review")
                return
            if workbook.report_only or comparison.report_only:
                report_only_reason = "large_file_report_only" if workbook.report_only else "large_difference_report_only"
                report.warnings.append(
                    "LARGE_DIFFERENCE_REPORT_ONLY: difference payload exceeded the in-memory threshold; full-workbook coloring was intentionally skipped"
                    if comparison.report_only and not workbook.report_only
                    else "LARGE_FILE_REPORT_ONLY: full-workbook coloring was intentionally skipped"
                )
                write_json_report(report, directory / "report.json")
                public_manifest = _non_rendering_manifest(report, report_only_reason)
                (directory / "render-manifest.json").write_text(json.dumps(public_manifest, ensure_ascii=False, indent=2), encoding="utf-8")
                write_html_report(report, directory / "report.html")
                object_keys = self._upload_artifacts(job_id, directory, [input_path.name, "report.json", "differences.jsonl", "report.html", "render-manifest.json"])
                self._write_status(job_id, {
                    **self.status(job_id),
                    "status": "completed",
                    "mode": "report_only",
                    "progress": 100,
                    "completed_at": _now(),
                    "summary": report.summary.model_dump(mode="json"),
                    "warnings": report.warnings,
                    "input_sha256": report.input_sha256,
                    "artifacts": {"json": "report.json", "differences_jsonl": "differences.jsonl", "html": "report.html", "manifest": "render-manifest.json"},
                    "object_keys": object_keys,
                })
                metrics.increment("jobs_terminal_total", status="completed", mode="report_only")
                return
            embedded_report = report.model_copy(deep=True)
            for difference in embedded_report.differences:
                if difference.repair_status == "planned":
                    difference.repair_status = "applied"
            embedded_report.summary.repairs_applied = embedded_report.summary.repairs_planned
            embedded_report.summary.repair_failures = 0
            write_json_report(embedded_report, directory / "report-render.json")
            private_manifest = _manifest(
                report,
                rules,
                comparison,
                report_source="report-render.json",
            )
            public_manifest = _redact_manifest(private_manifest)
            (directory / "render-manifest.private.json").write_text(json.dumps(private_manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
            (directory / "render-manifest.json").write_text(json.dumps(public_manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
            self._write_status(job_id, {**self.status(job_id), "status": "rendering", "progress": 75})
            self._check_cancelled(job_id)
            output_path = directory / f"result{input_suffix}"
            try:
                stage_started = time.perf_counter()
                render = self.renderer.render(input_path, output_path, workbook, rules, comparison, embedded_report.model_dump(mode="json"))
                metrics.observe("stage_duration_seconds", time.perf_counter() - stage_started, stage="rendering")
            finally:
                (directory / "render-manifest.private.json").unlink(missing_ok=True)
                (directory / "report-render.json").unlink(missing_ok=True)
            report.output_sha256 = render.sha256
            report.warnings.extend(render.warnings)
            applied_ids = {
                str(item.get("difference_id"))
                for item in render.operation_results
                if item.get("status") == "applied" and item.get("difference_id")
            }
            repair_rule_ids = {
                repair.difference_id: repair.rule_id
                for repair in comparison.repairs
            }
            for difference in report.differences:
                if difference.repair_status == "planned":
                    difference.repair_status = "applied" if difference.difference_id in applied_ids else "failed"
                    if self.database:
                        self.database.audit(
                            "comparison.auto_repair_operation",
                            "comparison_difference",
                            difference.difference_id,
                            user_id,
                            metadata={
                                "job_id": job_id,
                                "rule_id": repair_rule_ids.get(difference.difference_id, difference.rule_id),
                                "difference_rule_id": difference.rule_id,
                                "status": difference.repair_status,
                                "excel_raw_value": difference.excel_raw_value,
                                "standard_raw_value": difference.standard_raw_value,
                                "output_sha256": report.output_sha256,
                            },
                            tenant_id=tenant_id,
                        )
            report.summary.repairs_applied = sum(item.repair_status == "applied" for item in report.differences)
            report.summary.repair_failures = sum(item.repair_status == "failed" for item in report.differences)
            if self.database and report.summary.repairs_planned:
                self.database.mark_repair_results(job_id, {item.difference_id: item.repair_status for item in report.differences if item.repair_status in {"applied", "failed"}})
                self.database.audit("comparison.auto_repair_completed", "comparison_job", job_id, user_id, {"repair_count": report.summary.repairs_applied, "failure_count": report.summary.repair_failures, "output_sha256": report.output_sha256}, tenant_id)
            write_json_report(report, directory / "report.json")
            write_differences_jsonl(report.differences, directory / "differences.jsonl")
            write_html_report(report, directory / "report.html")
            object_keys = self._upload_artifacts(job_id, directory, [input_path.name, output_path.name, "report.json", "differences.jsonl", "report.html", "render-manifest.json"])
            self._write_status(job_id, {
                **self.status(job_id),
                "status": "completed",
                "progress": 100,
                "completed_at": _now(),
                "summary": report.summary.model_dump(mode="json"),
                "warnings": report.warnings,
                "input_sha256": report.input_sha256,
                "output_sha256": report.output_sha256,
                "artifacts": {"excel": output_path.name, "json": "report.json", "differences_jsonl": "differences.jsonl", "html": "report.html", "manifest": "render-manifest.json"},
                "object_keys": object_keys,
            })
            metrics.increment("jobs_terminal_total", status="completed")
            temp_bytes = sum(path.stat().st_size for path in directory.rglob("*") if path.is_file())
            metrics.set_gauge("last_job_temp_disk_bytes", temp_bytes)
            log_event(job_id=job_id, trace_id=trace_id, stage="completed", event="comparison.completed", duration_ms=int((time.perf_counter() - started) * 1000), rows_processed=row_count, differences=len(comparison.differences), temp_disk_bytes=temp_bytes)
        except JobCancelled:
            (directory / "render-manifest.private.json").unlink(missing_ok=True)
            (directory / "report-render.json").unlink(missing_ok=True)
            previous = self.status(job_id)
            self._write_status(job_id, {**previous, "status": "cancelled", "progress": previous.get("progress", 0), "completed_at": _now()})
            if self.database:
                self.database.audit("comparison.cancelled", "comparison_job", job_id, user_id, tenant_id=tenant_id)
            metrics.increment("jobs_terminal_total", status="cancelled")
            log_event(job_id=job_id, trace_id=trace_id, stage="cancelled", event="comparison.cancelled", duration_ms=int((time.perf_counter() - started) * 1000))
        except Exception as exc:
            (directory / "render-manifest.private.json").unlink(missing_ok=True)
            (directory / "report-render.json").unlink(missing_ok=True)
            previous = self.status(job_id)
            error_code = _safe_error_code(exc)
            diagnostic = directory / "diagnostic.log"
            diagnostic.write_text(json.dumps({
                "error_code": error_code,
                "exception_type": type(exc).__name__,
                "stage": previous.get("status", "unknown"),
                "trace_id": trace_id,
            }, ensure_ascii=False, sort_keys=True), encoding="utf-8")
            self._write_status(job_id, {**previous, "status": "failed", "progress": previous.get("progress", 0), "completed_at": _now(), "error_code": error_code, "error_message_safe": _safe_error_message(error_code)})
            metrics.increment("jobs_terminal_total", status="failed", error_code=error_code)
            log_event(job_id=job_id, trace_id=trace_id, stage=previous.get("status", "unknown"), event="comparison.failed", duration_ms=int((time.perf_counter() - started) * 1000), safe_error_code=error_code)
        finally:
            if comparison is not None:
                comparison.close()

    def _check_cancelled(self, job_id: str) -> None:
        if (self.job_directory(job_id) / "cancel.requested").is_file():
            raise JobCancelled()

    def status(self, job_id: str) -> dict[str, Any]:
        path = self.job_directory(job_id) / "status.json"
        return json.loads(path.read_text(encoding="utf-8"))

    def artifact(self, job_id: str, name: str) -> Path:
        if name not in {"excel", "json", "differences_jsonl", "html", "manifest"}:
            raise FileNotFoundError("unknown artifact")
        status = self.status(job_id)
        file_name = status.get("artifacts", {}).get(name)
        if not file_name or Path(file_name).name != file_name:
            raise FileNotFoundError("artifact not available")
        path = self.job_directory(job_id) / file_name
        if not path.is_file():
            raise FileNotFoundError("artifact not available")
        return path

    def job_directory(self, job_id: str) -> Path:
        if not job_id.startswith("job_") or not job_id[4:].isalnum():
            raise FileNotFoundError("invalid job id")
        path = self.jobs / job_id
        if not path.is_dir():
            raise FileNotFoundError("job not found")
        return path

    def _write_status(self, job_id: str, payload: dict[str, Any]) -> None:
        path = self.jobs / job_id / "status.json"
        temporary = path.with_suffix(".tmp")
        temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
        if self.database:
            self.database.update_job(payload)
        log_event(job_id=job_id, trace_id=payload.get("trace_id"), stage=str(payload.get("status", "unknown")), event="comparison.status_changed")

    def _upload_artifacts(self, job_id: str, directory: Path, names: list[str]) -> dict[str, str]:
        if not self.artifact_store:
            return {}
        result: dict[str, str] = dict(self.status(job_id).get("object_keys", {}))
        for name in names:
            path = directory / name
            if path.is_file():
                result[name] = self.artifact_store.put_file(f"jobs/{job_id}/artifacts/{name}", path)
        return result


def _retention_days(tenant_id: str) -> int:
    tenant_key = "".join(character if character.isalnum() else "_" for character in tenant_id).upper()
    value = os.environ.get(f"EXCEL_AUDITOR_RETENTION_DAYS_{tenant_key}", os.environ.get("EXCEL_AUDITOR_RETENTION_DAYS", "30"))
    days = int(value)
    if not 1 <= days <= 3650:
        raise ValueError("retention days must be between 1 and 3650")
    return days


def _make_read_only(path: Path) -> None:
    path.chmod(stat.S_IREAD)


def _remove_tree(directory: Path) -> None:
    for path in directory.rglob("*"):
        if path.is_file():
            try:
                path.chmod(stat.S_IREAD | stat.S_IWRITE)
            except OSError:
                pass
    shutil.rmtree(directory)


@contextmanager
def _tenant_quota_lock(root: Path, tenant_id: str):
    locks = root / ".locks"
    locks.mkdir(parents=True, exist_ok=True)
    lock = locks / ("quota-" + hashlib.sha256(tenant_id.encode("utf-8")).hexdigest()[:24] + ".lock")
    acquired = False
    for _attempt in range(200):
        try:
            lock.mkdir()
            acquired = True
            break
        except FileExistsError:
            try:
                if time.time() - lock.stat().st_mtime > 30:
                    lock.rmdir()
                    continue
            except (FileNotFoundError, OSError):
                continue
            time.sleep(0.01)
    if not acquired:
        raise ValueError("TENANT_QUOTA_EXCEEDED: quota reservation is busy")
    try:
        yield
    finally:
        try:
            lock.rmdir()
        except FileNotFoundError:
            pass


def _load_standard(path: Path, rules: RuleSet) -> dict[str, Sequence[dict[str, Any]]]:
    return load_standard_file(path, rules, spill_after_records=_standard_spill_threshold(rules))


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _field_statistics(comparison: Any) -> dict[str, dict[str, Any]]:
    if comparison.report_only:
        return _field_statistics_on_disk(comparison)
    fields = {
        (mapping.sheet_id, mapping.canonical_field)
        for mapping in comparison.mappings
        if mapping.status == "matched" and mapping.canonical_field
    }
    mismatches: dict[tuple[str, str], set[int]] = {}
    validation_errors: dict[tuple[str, str], set[int]] = {}
    for item in comparison.differences:
        if not item.canonical_field or item.excel_row is None:
            continue
        key = (item.sheet_id, item.canonical_field)
        fields.add(key)
        if item.type.value == "VALUE_MISMATCH":
            mismatches.setdefault(key, set()).add(item.excel_row)
        elif item.type.value in {"INVALID_VALUE", "VALIDATION_ERROR"}:
            validation_errors.setdefault(key, set()).add(item.excel_row)
    matched_by_sheet = comparison.matched_records_by_sheet or {}
    validated_by_sheet = comparison.validated_records_by_sheet or {}
    return {
        f"{sheet_id}.{name}": _field_statistic(
            sheet_id,
            name,
            len(mismatches.get((sheet_id, name), set())),
            len(validation_errors.get((sheet_id, name), set())),
            matched_by_sheet,
            validated_by_sheet,
        )
        for sheet_id, name in sorted(fields)
    }


def _field_statistics_on_disk(comparison: Any) -> dict[str, dict[str, Any]]:
    descriptor, raw_path = tempfile.mkstemp(prefix="excel-auditor-field-statistics-", suffix=".sqlite3")
    os.close(descriptor)
    path = Path(raw_path)
    connection = sqlite3.connect(path)
    try:
        connection.execute("PRAGMA journal_mode=OFF")
        connection.execute("PRAGMA synchronous=OFF")
        connection.execute("CREATE TABLE affected (sheet_id TEXT NOT NULL, field TEXT NOT NULL, kind TEXT NOT NULL, excel_row INTEGER NOT NULL, UNIQUE(sheet_id, field, kind, excel_row))")
        batch: list[tuple[str, str, str, int]] = []
        for item in comparison.differences:
            kind = "difference" if item.type.value == "VALUE_MISMATCH" else "validation" if item.type.value in {"INVALID_VALUE", "VALIDATION_ERROR"} else None
            if item.canonical_field and item.excel_row is not None and kind:
                batch.append((item.sheet_id, item.canonical_field, kind, item.excel_row))
                if len(batch) >= 1_000:
                    connection.executemany("INSERT OR IGNORE INTO affected VALUES (?, ?, ?, ?)", batch)
                    batch.clear()
        if batch:
            connection.executemany("INSERT OR IGNORE INTO affected VALUES (?, ?, ?, ?)", batch)
        counts = {
            (str(sheet_id), str(field), str(kind)): int(count)
            for sheet_id, field, kind, count in connection.execute("SELECT sheet_id, field, kind, COUNT(*) FROM affected GROUP BY sheet_id, field, kind")
        }
        fields = {
            (mapping.sheet_id, mapping.canonical_field)
            for mapping in comparison.mappings
            if mapping.status == "matched" and mapping.canonical_field
        } | {(sheet_id, field) for sheet_id, field, _kind in counts}
        matched_by_sheet = comparison.matched_records_by_sheet or {}
        validated_by_sheet = comparison.validated_records_by_sheet or {}
        return {
            f"{sheet_id}.{field}": _field_statistic(
                sheet_id,
                field,
                counts.get((sheet_id, field, "difference"), 0),
                counts.get((sheet_id, field, "validation"), 0),
                matched_by_sheet,
                validated_by_sheet,
            )
            for sheet_id, field in sorted(fields)
        }
    finally:
        connection.close()
        path.unlink(missing_ok=True)


def _field_statistic(
    sheet_id: str,
    field: str,
    difference_count: int,
    validation_error_count: int,
    matched_by_sheet: dict[str, int],
    validated_by_sheet: dict[str, int],
) -> dict[str, Any]:
    matched = matched_by_sheet.get(sheet_id, 0)
    validated = validated_by_sheet.get(sheet_id, 0)
    return {
        "sheet_id": sheet_id,
        "canonical_field": field,
        "compared_records": matched,
        "difference_count": difference_count,
        "difference_rate": difference_count / matched if matched else None,
        "validated_records": validated,
        "validation_error_count": validation_error_count,
        "validation_error_rate": validation_error_count / validated if validated else None,
    }


def _canonicalize_standard(payload: dict[str, Any], rules: RuleSet) -> dict[str, Sequence[dict[str, Any]]]:
    result: dict[str, Sequence[dict[str, Any]]] = {}
    spill_threshold = _standard_spill_threshold(rules)
    for key, rows in payload.items():
        if isinstance(rows, (str, bytes, bytearray)) or not isinstance(rows, SequenceABC):
            _close_record_sequences(result)
            raise ValueError(f"STANDARD_DATA_INVALID: {key} must be an array of objects")
        sheet_rule = next((sheet for sheet in rules.sheets if sheet.id == str(key) or sheet.name == str(key)), None)
        if sheet_rule is None:
            result[str(key)] = rows
            continue
        if sheet_rule.id in result:
            close_source = getattr(rows, "close", None)
            if close_source is not None:
                close_source()
            _close_record_sequences(result)
            raise ValueError(f"STANDARD_DATA_INVALID: duplicate sheet mapping: {sheet_rule.id}")
        canonical_rows: list[dict[str, Any]] | SpilledRecords = SpilledRecords() if len(rows) > spill_threshold else []
        try:
            for record_index, row in enumerate(rows, start=1):
                canonical_rows.append(canonicalize_standard_row(row, sheet_rule, record_index=record_index))
        except Exception:
            close = getattr(canonical_rows, "close", None)
            if close is not None:
                close()
            _close_record_sequences(result)
            raise
        finally:
            close_source = getattr(rows, "close", None)
            if close_source is not None:
                close_source()
        result[sheet_rule.id] = canonical_rows
    return result


def _standard_spill_threshold(rules: RuleSet) -> int:
    return min(
        100_000,
        max(
            10_000,
            rules.workbook.max_in_memory_cells // max(1, sum(len(sheet.columns) for sheet in rules.sheets)),
        ),
    )


def _close_record_sequences(standard: dict[str, Sequence[dict[str, Any]]]) -> None:
    for rows in standard.values():
        close = getattr(rows, "close", None)
        if close is not None:
            close()


def _non_rendering_manifest(report: AuditReport, reason: str) -> dict[str, Any]:
    return {
        "manifest_version": "1.0-public",
        "input_sha256": report.input_sha256,
        "metadata": {
            "schema_id": report.schema_id,
            "schema_version": report.schema_version,
            "schema_sha256": report.schema_sha256,
            "standard_sha256": report.standard_sha256,
        },
        "rendering": {"status": "skipped", "reason": reason},
        "operations": [],
    }


def _manifest(
    report: AuditReport,
    rules: RuleSet,
    comparison: Any,
    *,
    report_source: str = "report.json",
) -> dict[str, Any]:
    mark_operations: list[dict[str, Any]] = []
    insert_operations: list[dict[str, Any]] = []
    repair_operations: list[dict[str, Any]] = []
    color_by_action = {
        "mark_red": rules.colors.extra,
        "mark_row_red": rules.colors.extra,
        "mark_yellow": rules.colors.mismatch,
        "mark_orange": rules.colors.invalid,
        "mark_purple": rules.colors.ambiguous,
        "mark_row_purple": rules.colors.ambiguous,
    }
    sheets_by_id = {sheet.id: sheet for sheet in rules.sheets}
    differences_by_id = {item.difference_id: item for item in report.differences}
    mapped_by_sheet: dict[str, dict[str, int]] = {}
    for mapping in report.header_mappings:
        if mapping.status == "matched" and mapping.canonical_field:
            mapped_by_sheet.setdefault(mapping.sheet_id, {})[mapping.canonical_field] = mapping.physical_column
    final_columns_by_sheet: dict[str, dict[str, int]] = {}
    for sheet in rules.sheets:
        columns = dict(mapped_by_sheet.get(sheet.id, {}))
        resolved_header_row = (comparison.resolved_header_rows_by_sheet or {}).get(sheet.id, sheet.header.row)
        data_start_row = (comparison.data_start_rows_by_sheet or {}).get(
            sheet.id,
            sheet.data_region.start_row or resolved_header_row + 1,
        )
        formula_rows = (comparison.formula_rows_by_sheet or {}).get(sheet.id, [])
        missing = {
            item.canonical_field: item
            for item in report.differences
            if item.sheet_id == sheet.id and item.render_action == "insert_and_mark_green" and item.canonical_field
        }
        plans: list[tuple[int, int, Any, Any]] = []
        for index, column in enumerate(sheet.columns):
            item = missing.get(column.name)
            if item is None:
                continue
            previous = [columns[previous_column.name] for previous_column in sheet.columns[:index] if previous_column.name in columns]
            before_index = previous[-1] + 1 if previous else 1
            plans.append((before_index, index, column, item))
        # The renderer inserts from right to left. At an identical anchor, later
        # canonical fields must be inserted first to preserve schema order.
        for before_index, _index, column, item in sorted(plans, key=lambda plan: (plan[0], plan[1]), reverse=True):
            insert_operations.append({
                "type": "insert_column",
                "sheet": item.sheet_name,
                "before": _column_letter(before_index),
                "canonical_field": item.canonical_field,
                "header_row": resolved_header_row,
                "data_start_row": data_start_row,
                "header_value": column.title,
                "fill_color": rules.colors.inserted,
                "field_type": column.type.value,
                "number_format": _number_format(column),
                "validation": _excel_validation(column),
                "formula_template": column.formula_template,
                "formula_rows": formula_rows if column.formula_template else None,
                "comment": _repair_provenance_comment(
                    "自动插入缺失表头",
                    f"{rules.schema_id}@{rules.schema_version}:{column.name}.missing_column",
                    item,
                ),
                "difference_id": item.difference_id,
            })
            for name, position in list(columns.items()):
                if position >= before_index:
                    columns[name] = position + 1
            columns[column.name] = before_index
        final_columns_by_sheet[sheet.id] = columns
    mark_priority = {
        "mark_purple": 1,
        "mark_row_purple": 1,
        "mark_yellow": 2,
        "mark_orange": 3,
        "mark_red": 4,
        "mark_row_red": 4,
    }
    mark_targets: dict[tuple[str, str], dict[str, Any]] = {}
    for item in sorted(report.differences, key=lambda value: value.difference_id):
        if item.render_action.startswith("mark_row"):
            candidate = {"type": "mark_row", "sheet": item.sheet_name, "row": item.excel_row, "fill_color": color_by_action[item.render_action], "comment": item.message, "difference_id": item.difference_id, "_priority": mark_priority[item.render_action]}
            key = (item.sheet_name, f"row:{item.excel_row}")
        elif item.render_action in color_by_action:
            candidate = {"type": "mark_cell", "sheet": item.sheet_name, "cell": item.cell, "fill_color": color_by_action[item.render_action], "comment": item.message, "difference_id": item.difference_id, "_priority": mark_priority[item.render_action]}
            key = (item.sheet_name, f"cell:{item.cell}")
        else:
            continue
        existing = mark_targets.get(key)
        if existing is None:
            mark_targets[key] = candidate
        else:
            existing["comment"] = f"{existing['comment']} | {candidate['comment']}"[:32000]
            if candidate["_priority"] > existing["_priority"]:
                existing.update({name: value for name, value in candidate.items() if name != "comment"})
    for operation in sorted(mark_targets.values(), key=lambda item: (item["_priority"], item["sheet"], item.get("row") or 0, item.get("cell") or "")):
        operation.pop("_priority", None)
        mark_operations.append(operation)
    for repair in comparison.repairs:
        sheet_rule = sheets_by_id[repair.sheet_id]
        column_rule = next((column for column in sheet_rule.columns if column.name == repair.canonical_field), None)
        repair_comment = _repair_provenance_comment(
            "自动追加标准记录" if repair.type == "append_record" else "自动修复",
            repair.rule_id,
            differences_by_id.get(repair.difference_id),
        )
        if repair.type == "set_cell":
            repair_operations.append({"type": "set_cell", "sheet": repair.sheet_name, "cell": repair.cell, "value": repair.value, "fill_color": rules.colors.inserted, "comment": repair_comment, "difference_id": repair.difference_id})
            header_rename = repair.rule_id.endswith(".rename_confirmed_alias")
            repair_operations[-1].update({"field_type": "string" if header_rename else column_rule.type.value if column_rule else "string", "number_format": None if header_rename else _number_format(column_rule)})
        elif repair.type == "set_field":
            column = final_columns_by_sheet.get(repair.sheet_id, {}).get(repair.canonical_field or "")
            if column is None or repair.excel_row is None:
                raise ValueError(f"RENDER_FAILED: repaired field has no final column: {repair.canonical_field}")
            repair_operations.append({"type": "set_cell_after_insert", "sheet": repair.sheet_name, "cell": f"{_column_letter(column)}{repair.excel_row}", "value": repair.value, "fill_color": rules.colors.inserted, "comment": repair_comment, "difference_id": repair.difference_id})
            repair_operations[-1].update({"field_type": column_rule.type.value if column_rule else None, "number_format": _number_format(column_rule)})
        elif repair.type == "append_record":
            values = []
            for column in sheets_by_id[repair.sheet_id].columns:
                position = final_columns_by_sheet.get(repair.sheet_id, {}).get(column.name)
                if position is not None:
                    values.append({"cell": f"{_column_letter(position)}{repair.excel_row}", "value": (repair.values or {}).get(column.name), "field_type": column.type.value, "number_format": _number_format(column), "formula_template": column.formula_template})
            repair_operations.append({"type": "append_row", "sheet": repair.sheet_name, "row": repair.excel_row, "values": values, "fill_color": rules.colors.inserted, "comment": repair_comment, "difference_id": repair.difference_id})
    operations = [*mark_operations, *insert_operations, *repair_operations]
    operations.append({"type": "add_or_replace_report_sheet", "name": "核验报告", "source_json": "report.json"})
    operations[-1].update({"name": "核验报告", "source_json": report_source})
    return {
        "manifest_version": "1.0",
        "job_id": report.job_id,
        "input_sha256": report.input_sha256,
        "metadata": {
            "schema_id": report.schema_id,
            "schema_version": report.schema_version,
            "schema_sha256": report.schema_sha256,
            "standard_snapshot_id": report.standard_snapshot_id,
            "standard_sha256": report.standard_sha256,
        },
        "operations": operations,
    }


def _number_format(column: Any | None) -> str | None:
    if column is None:
        return None
    if column.format:
        return str(column.format)
    return {
        "integer": "0",
        "decimal": "0.############",
        "date": "yyyy-mm-dd",
        "datetime": "yyyy-mm-dd hh:mm:ss",
    }.get(column.type.value)


def _excel_validation(column: Any) -> dict[str, Any] | None:
    if column.type.value == "enum" and column.enum_values:
        return {"type": "list", "values": [str(value) for value in column.enum_values], "allow_blank": column.validation.nullable}
    if column.type.value in {"integer", "decimal"} and (column.validation.min is not None or column.validation.max is not None):
        return {
            "type": column.type.value,
            "min": str(column.validation.min) if column.validation.min is not None else None,
            "max": str(column.validation.max) if column.validation.max is not None else None,
            "allow_blank": column.validation.nullable,
        }
    return None


def _redact_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    redacted = json.loads(json.dumps(manifest, ensure_ascii=False, default=str))
    # Unique snapshot identifiers remain in the private renderer input
    # and hidden workbook metadata; omit them from the shareable manifest so
    # identical logical runs remain reproducible and unlinkable.
    metadata = redacted.get("metadata")
    if isinstance(metadata, dict):
        metadata.pop("standard_snapshot_id", None)
    for operation in redacted["operations"]:
        operation.pop("difference_id", None)
        if operation.get("type") in {"set_cell", "set_cell_after_insert", "append_row"} and isinstance(operation.get("comment"), str):
            rule_marker = "；规则："
            rule_id = operation["comment"].rsplit(rule_marker, 1)[-1] if rule_marker in operation["comment"] else "unknown"
            prefix = "自动追加标准记录" if operation.get("type") == "append_row" else "自动修复"
            operation["comment"] = f"{prefix}；原值和标准值已脱敏；规则：{rule_id}"
            operation["comment_values_redacted"] = True
        if "value" in operation:
            operation.pop("value")
            operation["value_redacted"] = True
        if "values" in operation:
            operation["values"] = [{"cell": item.get("cell"), "field_type": item.get("field_type"), "number_format": item.get("number_format"), "formula_template": item.get("formula_template"), "value_redacted": True} for item in operation["values"]]
    return redacted


def _column_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters or "A"


def _safe_error_code(exc: Exception) -> str:
    message = str(exc)
    known = ["FILE_UNSUPPORTED_FORMAT", "FILE_CORRUPTED", "FILE_LIMIT_EXCEEDED", "WORKBOOK_PROTECTED", "STANDARD_TOO_LARGE", "STANDARD_DATA_INVALID", "STANDARD_SOURCE_FAILED", "MANUAL_REVIEW_REQUIRED", "OUTPUT_VERIFICATION_FAILED", "UNSUPPORTED_FEATURE", "MANIFEST_OR_STRUCTURE_INVALID", "ARGUMENT_INVALID", "RENDER_FAILED"]
    return next((code for code in known if code in message), "COMPARISON_FAILED")


def _safe_error_message(error_code: str) -> str:
    messages = {
        "FILE_UNSUPPORTED_FORMAT": "The workbook format is not supported.",
        "FILE_CORRUPTED": "The workbook package is invalid or corrupted.",
        "FILE_LIMIT_EXCEEDED": "The workbook exceeds a configured safety limit.",
        "WORKBOOK_PROTECTED": "The workbook is protected and cannot be processed safely.",
        "STANDARD_TOO_LARGE": "The standard dataset exceeds a configured limit.",
        "STANDARD_DATA_INVALID": "The standard dataset failed validation.",
        "STANDARD_SOURCE_FAILED": "The managed standard source request failed.",
        "MANUAL_REVIEW_REQUIRED": "The workbook requires manual review.",
        "RENDER_FAILED": "Workbook rendering failed.",
        "OUTPUT_VERIFICATION_FAILED": "The rendered workbook failed output verification.",
        "UNSUPPORTED_FEATURE": "The workbook contains a feature the renderer cannot modify safely.",
        "MANIFEST_OR_STRUCTURE_INVALID": "The render request or workbook structure is invalid.",
        "ARGUMENT_INVALID": "The renderer invocation is invalid.",
        "COMPARISON_FAILED": "The comparison failed.",
    }
    return messages.get(error_code, messages["COMPARISON_FAILED"])


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


class JobCancelled(Exception):
    pass
