from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from typing import Any

from sqlalchemy import DateTime, ForeignKey, Index, Integer, JSON, String, Text, UniqueConstraint, create_engine, delete, insert, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column

from .models import AuditReport, RuleSet
from .snapshots import StandardSnapshot
from .ids import new_ulid


class Base(DeclarativeBase):
    pass


class SchemaRow(Base):
    __tablename__ = "schemas"
    __table_args__ = (UniqueConstraint("tenant_id", "schema_id", name="uq_tenant_schema"),)
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(200), index=True)
    schema_id: Mapped[str] = mapped_column(String(200), index=True)
    name: Mapped[str] = mapped_column(String(300))
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    current_version: Mapped[str | None] = mapped_column(String(32), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))


class SchemaVersionRow(Base):
    __tablename__ = "schema_versions"
    __table_args__ = (UniqueConstraint("tenant_id", "schema_id", "version", name="uq_tenant_schema_version"),)
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(200), index=True)
    schema_id: Mapped[str] = mapped_column(String(200), index=True)
    version: Mapped[str] = mapped_column(String(32))
    status: Mapped[str] = mapped_column(String(24), default="published")
    config_json: Mapped[dict[str, Any]] = mapped_column(JSON)
    config_sha256: Mapped[str] = mapped_column(String(64))
    published_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    created_by: Mapped[str | None] = mapped_column(String(200), nullable=True)
    change_summary: Mapped[str | None] = mapped_column(Text, nullable=True)


class StandardConnectionRow(Base):
    __tablename__ = "standard_connections"
    __table_args__ = (UniqueConstraint("tenant_id", "connection_id", name="uq_tenant_standard_connection"),)
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(200), index=True)
    connection_id: Mapped[str] = mapped_column(String(200), index=True)
    source_type: Mapped[str] = mapped_column(String(32))
    config_ciphertext: Mapped[str] = mapped_column(Text)
    secret_ref: Mapped[str | None] = mapped_column(String(300), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))


class StandardSnapshotRow(Base):
    __tablename__ = "standard_snapshots"
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    source_type: Mapped[str] = mapped_column(String(32))
    object_key: Mapped[str] = mapped_column(Text)
    content_sha256: Mapped[str] = mapped_column(String(64), index=True)
    record_count: Mapped[int] = mapped_column(Integer)
    fetched_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    metadata_json: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    tenant_id: Mapped[str] = mapped_column(String(200), index=True)
    schema_id: Mapped[str] = mapped_column(String(200), index=True)
    schema_version: Mapped[str] = mapped_column(String(32))
    status: Mapped[str] = mapped_column(String(24), default="ready")
    expires_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)


class ComparisonJobRow(Base):
    __tablename__ = "comparison_jobs"
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    tenant_id: Mapped[str] = mapped_column(String(200), index=True)
    user_id: Mapped[str] = mapped_column(String(200))
    schema_id: Mapped[str | None] = mapped_column(String(200), nullable=True)
    schema_version: Mapped[str | None] = mapped_column(String(32), nullable=True)
    standard_snapshot_id: Mapped[str | None] = mapped_column(ForeignKey("standard_snapshots.id"), nullable=True)
    status: Mapped[str] = mapped_column(String(32), index=True)
    progress: Mapped[int] = mapped_column(Integer, default=0)
    input_sha256: Mapped[str | None] = mapped_column(String(64), nullable=True)
    output_sha256: Mapped[str | None] = mapped_column(String(64), nullable=True)
    summary_json: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    error_code: Mapped[str | None] = mapped_column(String(80), nullable=True)
    error_message_safe: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))
    completed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    retention_until: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    deleted_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    parameters_json: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    input_object_key: Mapped[str | None] = mapped_column(Text, nullable=True)
    output_object_key: Mapped[str | None] = mapped_column(Text, nullable=True)
    report_object_keys: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    renderer_version: Mapped[str | None] = mapped_column(String(100), nullable=True)
    started_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)


class DifferenceIndexRow(Base):
    __tablename__ = "comparison_differences"
    __table_args__ = (Index("ix_differences_job_type", "job_id", "type"),)
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    job_id: Mapped[str] = mapped_column(ForeignKey("comparison_jobs.id", ondelete="CASCADE"), primary_key=True)
    type: Mapped[str] = mapped_column(String(64))
    severity: Mapped[str] = mapped_column(String(16))
    sheet_id: Mapped[str] = mapped_column(String(200))
    cell_ref: Mapped[str | None] = mapped_column(String(32), nullable=True)
    canonical_field: Mapped[str | None] = mapped_column(String(200), nullable=True)
    business_key_hash: Mapped[str | None] = mapped_column(String(64), nullable=True)
    message_safe: Mapped[str] = mapped_column(Text)
    render_action: Mapped[str] = mapped_column(String(64))
    repair_status: Mapped[str] = mapped_column(String(32))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))


class AuditEventRow(Base):
    __tablename__ = "audit_events"
    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    tenant_id: Mapped[str | None] = mapped_column(String(200), nullable=True, index=True)
    actor_id: Mapped[str | None] = mapped_column(String(200), nullable=True)
    action: Mapped[str] = mapped_column(String(100), index=True)
    resource_type: Mapped[str] = mapped_column(String(100))
    resource_id: Mapped[str] = mapped_column(String(200))
    metadata_json: Mapped[dict[str, Any]] = mapped_column(JSON, default=dict)
    occurred_at: Mapped[datetime] = mapped_column(DateTime(timezone=True))


class DatabaseRepository:
    def __init__(self, url: str, create_schema: bool = True) -> None:
        self.engine = create_engine(url, pool_pre_ping=True)
        if create_schema:
            Base.metadata.create_all(self.engine)

    def publish_rule(self, rules: RuleSet, actor_id: str | None = None, tenant_id: str = "local") -> None:
        with Session(self.engine) as session, session.begin():
            existing = session.scalar(select(SchemaVersionRow).where(SchemaVersionRow.tenant_id == tenant_id, SchemaVersionRow.schema_id == rules.schema_id, SchemaVersionRow.version == rules.schema_version))
            if existing:
                if existing.config_sha256 != rules.content_sha256:
                    raise FileExistsError("published rule versions are immutable")
                return
            now = _now()
            schema = session.scalar(select(SchemaRow).where(SchemaRow.tenant_id == tenant_id, SchemaRow.schema_id == rules.schema_id))
            if schema is None:
                schema = SchemaRow(id=new_ulid("schema_"), tenant_id=tenant_id, schema_id=rules.schema_id, name=rules.name, description=None, current_version=rules.schema_version, created_at=now, updated_at=now)
                session.add(schema)
            else:
                schema.name = rules.name
                schema.current_version = rules.schema_version
                schema.updated_at = now
            session.add(SchemaVersionRow(id=new_ulid("sv_"), tenant_id=tenant_id, schema_id=rules.schema_id, version=rules.schema_version, config_json=rules.model_dump(mode="json"), config_sha256=rules.content_sha256, published_at=now, created_by=actor_id))
            session.add(_event("schema.published", "schema_version", f"{rules.schema_id}@{rules.schema_version}", actor_id, {"config_sha256": rules.content_sha256}, tenant_id))

    def ping(self) -> None:
        with self.engine.connect() as connection:
            connection.execute(select(1))

    def get_rule(self, schema_id: str, version: str, tenant_id: str = "local") -> RuleSet:
        with Session(self.engine) as session:
            row = session.scalar(select(SchemaVersionRow).where(SchemaVersionRow.tenant_id == tenant_id, SchemaVersionRow.schema_id == schema_id, SchemaVersionRow.version == version))
            if row is None:
                raise FileNotFoundError(f"rule version not found: {schema_id}@{version}")
            return RuleSet.model_validate(row.config_json)

    def list_rule_versions(self, schema_id: str, tenant_id: str = "local") -> list[dict[str, str]]:
        with Session(self.engine) as session:
            rows = session.scalars(select(SchemaVersionRow).where(SchemaVersionRow.tenant_id == tenant_id, SchemaVersionRow.schema_id == schema_id).order_by(SchemaVersionRow.version)).all()
            return [{"schema_id": row.schema_id, "version": row.version, "config_sha256": row.config_sha256, "status": row.status} for row in rows]

    def create_job(self, job_id: str, tenant_id: str = "local", user_id: str = "local") -> None:
        with Session(self.engine) as session, session.begin():
            session.add(ComparisonJobRow(id=job_id, tenant_id=tenant_id, user_id=user_id, status="queued", progress=0, created_at=_now()))
            session.add(_event("comparison.created", "comparison_job", job_id, user_id, {"tenant_id": tenant_id}, tenant_id))

    def update_job(self, payload: dict[str, Any]) -> None:
        with Session(self.engine) as session, session.begin():
            row = session.get(ComparisonJobRow, payload["job_id"])
            if row is None:
                return
            for key in ["status", "progress", "schema_id", "schema_version", "standard_snapshot_id", "input_sha256", "output_sha256", "error_code", "error_message_safe"]:
                if key in payload:
                    setattr(row, key, payload[key])
            if "summary" in payload:
                row.summary_json = payload["summary"]
            if payload.get("completed_at"):
                row.completed_at = datetime.fromisoformat(payload["completed_at"])
            if payload.get("retention_until"):
                row.retention_until = datetime.fromisoformat(payload["retention_until"])
            if payload.get("deleted_at"):
                row.deleted_at = datetime.fromisoformat(payload["deleted_at"])

    def save_snapshot(self, snapshot: StandardSnapshot, source_type: str, tenant_id: str, object_key: str | None = None) -> None:
        with Session(self.engine) as session, session.begin():
            session.add(StandardSnapshotRow(id=snapshot.snapshot_id, tenant_id=tenant_id, schema_id=str(snapshot.metadata.get("schema_id", "")), schema_version=str(snapshot.metadata.get("schema_version", "")), status="ready", source_type=source_type, object_key=object_key or str(snapshot.path), content_sha256=snapshot.sha256, record_count=snapshot.record_count, fetched_at=snapshot.fetched_at, metadata_json=snapshot.metadata))

    def save_differences(self, report: AuditReport) -> None:
        now = _now()
        with Session(self.engine) as session, session.begin():
            batch: list[dict[str, Any]] = []
            for item in report.differences:
                key_hash = None
                if item.business_key is not None:
                    stable = json.dumps(item.business_key, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                    key_hash = hashlib.sha256(stable.encode("utf-8")).hexdigest()
                batch.append({
                    "id": item.difference_id,
                    "job_id": report.job_id,
                    "type": item.type.value,
                    "severity": item.severity,
                    "sheet_id": item.sheet_id,
                    "cell_ref": item.cell,
                    "canonical_field": item.canonical_field,
                    "business_key_hash": key_hash,
                    "message_safe": item.message,
                    "render_action": item.render_action,
                    "repair_status": item.repair_status,
                    "created_at": now,
                })
                if len(batch) >= 1_000:
                    session.execute(insert(DifferenceIndexRow), batch)
                    batch.clear()
            if batch:
                session.execute(insert(DifferenceIndexRow), batch)

    def mark_repair_results(self, job_id: str, results: dict[str, str]) -> None:
        with Session(self.engine) as session, session.begin():
            rows = session.scalars(select(DifferenceIndexRow).where(DifferenceIndexRow.job_id == job_id, DifferenceIndexRow.repair_status == "planned")).all()
            for row in rows:
                if row.id in results:
                    row.repair_status = results[row.id]

    def audit(self, action: str, resource_type: str, resource_id: str, actor_id: str | None = None, metadata: dict[str, Any] | None = None, tenant_id: str | None = None) -> None:
        with Session(self.engine) as session, session.begin():
            session.add(_event(action, resource_type, resource_id, actor_id, metadata, tenant_id))

    def purge_job(self, job_id: str, tenant_id: str) -> None:
        """Permanently remove expired operational rows while retaining audit events."""
        with Session(self.engine) as session, session.begin():
            job = session.scalar(select(ComparisonJobRow).where(ComparisonJobRow.id == job_id, ComparisonJobRow.tenant_id == tenant_id))
            if job is None:
                return
            snapshot_id = job.standard_snapshot_id
            session.execute(delete(DifferenceIndexRow).where(DifferenceIndexRow.job_id == job_id))
            session.delete(job)
            session.flush()
            if snapshot_id is not None:
                still_referenced = session.scalar(select(ComparisonJobRow.id).where(ComparisonJobRow.standard_snapshot_id == snapshot_id).limit(1))
                if still_referenced is None:
                    snapshot = session.get(StandardSnapshotRow, snapshot_id)
                    if snapshot is not None:
                        session.delete(snapshot)


def _event(action: str, resource_type: str, resource_id: str, actor_id: str | None = None, metadata: dict[str, Any] | None = None, tenant_id: str | None = None) -> AuditEventRow:
    return AuditEventRow(id=new_ulid("audit_"), tenant_id=tenant_id, actor_id=actor_id, action=action, resource_type=resource_type, resource_id=resource_id, metadata_json=metadata or {}, occurred_at=_now())


def _now() -> datetime:
    return datetime.now(timezone.utc)
