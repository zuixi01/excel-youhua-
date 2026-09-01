from __future__ import annotations

import json
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from ..ids import new_ulid
from ..models import RuleSet
from ..persistence import DatabaseRepository
from ..rendering import DotNetOpenXmlRenderer, _set_result_content_hash, _worksheet_package_entry
from ..service import AuditService
from ..workbook import inspect_workbook, sha256_file
from .catalog import CatalogAdapter
from .models import ProductReviewDecision
from .normalizer import normalize_product_workbook


class ProductWorkflowService:
    """Execute and persist a versioned dynamic-product normalization analysis."""

    def __init__(self, audit_service: AuditService, database: DatabaseRepository | None = None) -> None:
        self.audit_service = audit_service
        self.database = database

    def run(
        self,
        job_id: str,
        excel_path: Path,
        rules: RuleSet,
        catalog: CatalogAdapter,
        *,
        tenant_id: str = "local",
        actor_id: str = "local",
        confirmed_aliases: dict[str, str] | None = None,
        confirmed_aliases_by_category: dict[str, dict[str, str]] | None = None,
        category_overrides: dict[int, str] | None = None,
        forced_extra_columns: set[int] | None = None,
        parent_revision_id: str | None = None,
    ) -> None:
        directory = self.audit_service.job_directory(job_id)
        current = self.audit_service.status(job_id)
        self.audit_service._write_status(job_id, {
            **current,
            "status": "validating",
            "progress": 15,
            "completed_at": None,
            "workflow": "product_normalization",
            "schema_id": rules.schema_id,
            "schema_version": rules.schema_version,
        })
        workbook = None
        try:
            workbook = inspect_workbook(excel_path, rules)
            self.audit_service._write_status(job_id, {
                **self.audit_service.status(job_id),
                "status": "catalog_resolving",
                "progress": 40,
                "input_sha256": workbook.sha256,
            })
            result = normalize_product_workbook(
                workbook,
                rules,
                catalog,
                confirmed_aliases=confirmed_aliases,
                confirmed_aliases_by_category=confirmed_aliases_by_category,
                category_overrides=category_overrides,
                forced_extra_columns=forced_extra_columns,
            )
            payload = result.model_dump(mode="json")
            result_name = "product-result.json"
            _write_json_atomic(directory / result_name, payload)
            for snapshot in result.catalog_snapshots:
                if self.database:
                    self.database.save_product_catalog_snapshot(snapshot, tenant_id)
            if self.database:
                revision = self.database.create_product_revision(
                    job_id,
                    result,
                    actor_id=actor_id,
                    parent_revision_id=parent_revision_id,
                    tenant_id=tenant_id,
                )
                reviews = self.database.list_product_reviews(
                    job_id,
                    revision_id=revision["revision_id"],
                    tenant_id=tenant_id,
                )
            else:
                previous_state = self._state(job_id) if (directory / "product-workflow.json").is_file() else {}
                previous_revision = previous_state.get("current_revision", {})
                next_number = int(previous_revision.get("revision_number", 0)) + 1
                revision = {
                    "revision_id": new_ulid("prev_"),
                    "revision_number": next_number,
                    "parent_revision_id": parent_revision_id or previous_revision.get("revision_id"),
                    "status": "manual_review" if result.requires_manual_review else "ready",
                }
                reviews = [
                    {
                        "review_id": new_ulid("review_"),
                        "job_id": job_id,
                        "revision_id": revision["revision_id"],
                        "review_key": item.key,
                        "review_type": item.review_type,
                        "status": "pending",
                        "payload": item.model_dump(mode="json"),
                        "decision": None,
                        "created_at": current["created_at"],
                        "decided_at": None,
                        "decided_by": None,
                    }
                    for item in result.review_items
                ]
            previous_state = self._state(job_id) if (directory / "product-workflow.json").is_file() else {}
            history = list(previous_state.get("revision_history", []))
            history.append(revision)
            state = {
                "job_id": job_id,
                "current_revision": revision,
                "revision_history": history,
                "reviews": reviews,
            }
            _write_json_atomic(directory / "product-workflow.json", state)
            _write_json_atomic(directory / f"product-result-r{revision['revision_number']}.json", payload)
            artifacts = {"product_result": result_name}
            artifact_names = [excel_path.name, result_name, f"product-result-r{revision['revision_number']}.json", "product-workflow.json"]
            output_sha256 = None
            status = "manual_review"
            if not result.requires_manual_review:
                status = "rendering"
                self.audit_service._write_status(job_id, {
                    **self.audit_service.status(job_id),
                    "status": status,
                    "progress": 75,
                })
                destination = directory / f"product-result{excel_path.suffix.lower()}"
                manifest_name = "product-render-manifest.json"
                manifest_path = directory / manifest_name
                manifest = {
                    "manifest_version": "1.0",
                    "job_id": job_id,
                    "input_sha256": workbook.sha256,
                    "metadata": {
                        "schema_id": rules.schema_id,
                        "schema_version": rules.schema_version,
                        "schema_sha256": rules.content_sha256,
                    },
                    "operations": [{
                        "type": "add_or_replace_product_sheets",
                        "source_json": result_name,
                    }],
                }
                _write_json_atomic(manifest_path, manifest)
                if isinstance(self.audit_service.renderer, DotNetOpenXmlRenderer):
                    render = self.audit_service.renderer.render_manifest(
                        excel_path,
                        destination,
                        manifest_path,
                        rules.workbook.processing_timeout_seconds,
                    )
                    output_sha256 = render.sha256
                else:
                    output_sha256 = _render_product_development(
                        excel_path,
                        destination,
                        result,
                        job_id=job_id,
                        rules=rules,
                    )
                artifacts.update({"product_excel": destination.name, "product_manifest": manifest_name})
                artifact_names.extend([destination.name, manifest_name])
                status = "completed"
            object_keys = self.audit_service._upload_artifacts(job_id, directory, artifact_names)
            self.audit_service._write_status(job_id, {
                **self.audit_service.status(job_id),
                "status": status,
                "progress": 100,
                "completed_at": datetime.now(UTC).isoformat(),
                "category_count": len(result.category_sheets),
                "unresolved_row_count": len(result.unresolved_rows),
                "review_count": len(reviews),
                "issue_count": len(result.issues),
                "revision_id": revision["revision_id"],
                "revision_number": revision["revision_number"],
                "output_sha256": output_sha256,
                "artifacts": artifacts,
                "object_keys": object_keys,
            })
        except Exception as exc:
            _write_json_atomic(directory / "product-diagnostic.json", {
                "exception_type": type(exc).__name__,
                "stage": self.audit_service.status(job_id).get("status"),
            })
            self.audit_service._write_status(job_id, {
                **self.audit_service.status(job_id),
                "status": "failed",
                "completed_at": datetime.now(UTC).isoformat(),
                "error_code": "PRODUCT_NORMALIZATION_FAILED",
                "error_message_safe": "商品表格规范化失败，请检查输入、规则和平台目录配置。",
            })
        finally:
            if workbook is not None:
                workbook.close()

    def list_reviews(
        self,
        job_id: str,
        *,
        tenant_id: str = "local",
        status: str | None = None,
    ) -> list[dict[str, Any]]:
        if self.database:
            return self.database.list_product_reviews(job_id, status=status, tenant_id=tenant_id)
        state = self._state(job_id)
        reviews = state.get("reviews", [])
        return [item for item in reviews if status is None or item.get("status") == status]

    def decide_review(
        self,
        job_id: str,
        review_id: str,
        decision: ProductReviewDecision,
        *,
        tenant_id: str = "local",
        actor_id: str = "local",
    ) -> dict[str, Any]:
        state = self._state(job_id)
        review = next((item for item in state.get("reviews", []) if item.get("review_id") == review_id), None)
        if review is None:
            raise FileNotFoundError("product review item not found")
        _validate_review_decision(review, decision)
        if self.database:
            updated = self.database.decide_product_review(
                job_id,
                review_id,
                decision.model_dump(mode="json", exclude_none=True),
                actor_id=actor_id,
                tenant_id=tenant_id,
            )
        else:
            if review.get("status") != "pending":
                raise ValueError("product review item has already been decided")
            review["status"] = "resolved"
            review["decision"] = decision.model_dump(mode="json", exclude_none=True)
            review["decided_by"] = actor_id
            review["decided_at"] = datetime.now(UTC).isoformat()
            updated = review
        for index, item in enumerate(state.get("reviews", [])):
            if item.get("review_id") == review_id:
                state["reviews"][index] = updated
                break
        _write_json_atomic(self.audit_service.job_directory(job_id) / "product-workflow.json", state)
        if all(item.get("status") == "resolved" for item in state.get("reviews", [])):
            self.audit_service._write_status(job_id, {
                **self.audit_service.status(job_id),
                "status": "review_resolved",
                "pending_review_count": 0,
            })
        return updated

    def rerun_after_reviews(
        self,
        job_id: str,
        rules: RuleSet,
        catalog: CatalogAdapter,
        *,
        tenant_id: str = "local",
        actor_id: str = "local",
    ) -> None:
        state = self._state(job_id)
        reviews = state.get("reviews", [])
        if any(review.get("status") != "resolved" for review in reviews):
            raise ValueError("all review items must be resolved before creating a revision")
        decisions = [review.get("decision") or {} for review in reviews]
        if any(decision.get("action") == "reject" for decision in decisions):
            raise ValueError("rejected review items must be corrected before creating a revision")
        category_overrides: dict[int, str] = {}
        confirmed_aliases_by_category: dict[str, dict[str, str]] = {}
        forced_extra_columns: set[int] = set()
        for review, decision in zip(reviews, decisions, strict=True):
            action = decision.get("action")
            payload = review.get("payload", {})
            if action == "confirm_category":
                category_overrides[int(payload["excel_row"])] = str(decision["category_id"])
            elif action == "confirm_mapping":
                category_id = str(payload["category_id"])
                confirmed_aliases_by_category.setdefault(category_id, {})[
                    str(decision["raw_header"])
                ] = str(decision["field_id"])
            elif action == "keep_extra" and payload.get("physical_column") is not None:
                forced_extra_columns.add(int(payload["physical_column"]))
        directory = self.audit_service.job_directory(job_id)
        inputs = [path for path in directory.glob("product-input.*") if path.is_file()]
        if len(inputs) != 1:
            raise FileNotFoundError("original product workbook is unavailable")
        parent = state.get("current_revision", {}).get("revision_id")
        self.run(
            job_id,
            inputs[0],
            rules,
            catalog,
            tenant_id=tenant_id,
            actor_id=actor_id,
            confirmed_aliases_by_category=confirmed_aliases_by_category,
            category_overrides=category_overrides,
            forced_extra_columns=forced_extra_columns,
            parent_revision_id=parent,
        )

    def rerun_after_reviews_safe(
        self,
        job_id: str,
        rules: RuleSet,
        catalog: CatalogAdapter,
        *,
        tenant_id: str = "local",
        actor_id: str = "local",
    ) -> None:
        try:
            self.rerun_after_reviews(
                job_id,
                rules,
                catalog,
                tenant_id=tenant_id,
                actor_id=actor_id,
            )
        except Exception as exc:
            directory = self.audit_service.job_directory(job_id)
            _write_json_atomic(directory / "product-diagnostic.json", {
                "exception_type": type(exc).__name__,
                "stage": "revision_queued",
            })
            self.audit_service._write_status(job_id, {
                **self.audit_service.status(job_id),
                "status": "failed",
                "completed_at": datetime.now(UTC).isoformat(),
                "error_code": "PRODUCT_REVISION_FAILED",
                "error_message_safe": "商品表格修订失败，请重新检查审核决定和原始文件。",
            })

    def _state(self, job_id: str) -> dict[str, Any]:
        path = self.audit_service.job_directory(job_id) / "product-workflow.json"
        if not path.is_file():
            raise FileNotFoundError("product workflow state not found")
        return json.loads(path.read_text(encoding="utf-8"))


def _validate_review_decision(review: dict[str, Any], decision: ProductReviewDecision) -> None:
    review_type = review.get("review_type")
    if review_type == "category" and decision.action not in {"confirm_category", "reject"}:
        raise ValueError("category review requires confirm_category or reject")
    if review_type == "field_mapping" and decision.action not in {"confirm_mapping", "keep_extra", "reject"}:
        raise ValueError("field mapping review requires confirm_mapping, keep_extra, or reject")
    if review_type == "duplicate_header" and decision.action not in {"keep_extra", "reject"}:
        raise ValueError("duplicate header review cannot be auto-mapped")
    if decision.action in {"confirm_category", "confirm_mapping"}:
        target = decision.category_id if decision.action == "confirm_category" else decision.field_id
        candidates = {
            candidate.get("field_id")
            for candidate in review.get("payload", {}).get("candidates", [])
        }
        if candidates and target not in candidates:
            raise ValueError("review decision target is not one of the recorded candidates")
    if decision.action == "confirm_mapping" and decision.raw_header != review.get("payload", {}).get("raw_header"):
        raise ValueError("confirmed raw_header does not match the reviewed physical header")


def _write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temporary.replace(path)


def _render_product_development(
    source: Path,
    destination: Path,
    result: Any,
    *,
    job_id: str,
    rules: RuleSet,
) -> str:
    """Local fallback mirroring the strict product-sheet operation used by .NET."""
    shutil.copy2(source, destination)
    book = load_workbook(destination, data_only=False, keep_links=False)
    issue_colors = {
        (issue.category_id, issue.excel_row, issue.field_id): issue.color
        for issue in result.issues
        if issue.category_id and issue.field_id
    }
    used_names = {name.casefold() for name in book.sheetnames}
    for category in result.category_sheets:
        name = category.worksheet_name
        if name in book.sheetnames:
            del book[name]
            used_names.discard(name.casefold())
        sheet = book.create_sheet(name)
        used_names.add(name.casefold())
        fields = [item.field for item in category.plan.fields]
        _write_product_sheet(sheet, fields, category.rows, category.source_excel_rows, category.category_id, issue_colors, result.merchant_extra_header_color)
        if category.sku_rows:
            sku_name = _unique_local_sheet_name(f"{name}-SKU", category.category_id, used_names)
            sku_sheet = book.create_sheet(sku_name)
            sku_fields = [
                field for field in fields
                if field.source.value in {"fixed", "platform_specification"}
            ]
            _write_product_sheet(
                sku_sheet,
                sku_fields,
                category.sku_rows,
                category.source_excel_rows,
                category.category_id,
                issue_colors,
                result.merchant_extra_header_color,
            )
            used_names.add(sku_name.casefold())
    if "__ExcelAuditorMetadata" in book.sheetnames:
        del book["__ExcelAuditorMetadata"]
    metadata = book.create_sheet("__ExcelAuditorMetadata")
    metadata.append(["key", "value"])
    for key, value in [
        ("job_id", job_id),
        ("schema_id", rules.schema_id),
        ("schema_version", rules.schema_version),
        ("schema_sha256", rules.content_sha256),
        ("standard_snapshot_id", ""),
        ("standard_sha256", ""),
        ("input_sha256", sha256_file(source)),
        ("result_content_sha256", ""),
        ("operation_count", len(result.category_sheets)),
    ]:
        metadata.append([key, value])
    metadata.sheet_state = "veryHidden"
    book.save(destination)
    book.close()
    metadata_entry = _worksheet_package_entry(destination, "__ExcelAuditorMetadata")
    _set_result_content_hash(destination, metadata_entry)
    load_workbook(destination, read_only=True, data_only=False).close()
    return sha256_file(destination)


def _write_product_sheet(
    sheet: Any,
    fields: list[Any],
    rows: list[dict[str, Any]],
    source_rows: list[int],
    category_id: str,
    issue_colors: dict[tuple[str, int, str], str],
    merchant_extra_header_color: str,
) -> None:
    header_colors = {
        "fixed": "DDEBF7",
        "platform_attribute": "E2F0D9",
        "platform_specification": "FFF2CC",
        "merchant_extra": merchant_extra_header_color,
    }
    for column, field in enumerate(fields, start=1):
        cell = sheet.cell(1, column, field.title)
        cell.fill = PatternFill("solid", fgColor=header_colors[field.source.value])
    for output_row, (source_row, values) in enumerate(zip(source_rows, rows, strict=True), start=2):
        for column, field in enumerate(fields, start=1):
            cell = sheet.cell(output_row, column, values.get(field.field_id))
            if field.number_format:
                cell.number_format = field.number_format
            color = issue_colors.get((category_id, source_row, field.field_id))
            if color:
                cell.fill = PatternFill("solid", fgColor=color)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{sheet.cell(1, len(fields)).coordinate[:-1]}{max(1, len(rows) + 1)}"


def _unique_local_sheet_name(base: str, category_id: str, used: set[str]) -> str:
    cleaned = "".join("-" if character in "[]:*?/\\" else character for character in base)[:31]
    if cleaned.casefold() not in used:
        return cleaned
    suffix = f"-{category_id}"[:12]
    candidate = f"{cleaned[:31-len(suffix)]}{suffix}"
    counter = 2
    while candidate.casefold() in used:
        marker = f"-{counter}"
        candidate = f"{cleaned[:31-len(suffix)-len(marker)]}{suffix}{marker}"
        counter += 1
    return candidate
