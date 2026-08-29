"""Initial rule, snapshot, comparison, difference, and audit tables."""

from alembic import op
import sqlalchemy as sa


revision = "20260828_0001"
down_revision = None
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "schema_versions",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("schema_id", sa.String(200), nullable=False),
        sa.Column("version", sa.String(32), nullable=False),
        sa.Column("status", sa.String(24), nullable=False),
        sa.Column("config_json", sa.JSON(), nullable=False),
        sa.Column("config_sha256", sa.String(64), nullable=False),
        sa.Column("published_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("tenant_id", "schema_id", "version", name="uq_tenant_schema_version"),
    )
    op.create_index("ix_schema_versions_tenant_id", "schema_versions", ["tenant_id"])
    op.create_index("ix_schema_versions_schema_id", "schema_versions", ["schema_id"])
    op.create_table(
        "standard_snapshots",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("source_type", sa.String(32), nullable=False),
        sa.Column("object_key", sa.Text(), nullable=False),
        sa.Column("content_sha256", sa.String(64), nullable=False),
        sa.Column("record_count", sa.Integer(), nullable=False),
        sa.Column("fetched_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("metadata_json", sa.JSON(), nullable=False),
    )
    op.create_index("ix_standard_snapshots_content_sha256", "standard_snapshots", ["content_sha256"])
    op.create_table(
        "comparison_jobs",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("user_id", sa.String(200), nullable=False),
        sa.Column("schema_id", sa.String(200), nullable=True),
        sa.Column("schema_version", sa.String(32), nullable=True),
        sa.Column("standard_snapshot_id", sa.String(64), sa.ForeignKey("standard_snapshots.id"), nullable=True),
        sa.Column("status", sa.String(32), nullable=False),
        sa.Column("progress", sa.Integer(), nullable=False),
        sa.Column("input_sha256", sa.String(64), nullable=True),
        sa.Column("output_sha256", sa.String(64), nullable=True),
        sa.Column("summary_json", sa.JSON(), nullable=False),
        sa.Column("error_code", sa.String(80), nullable=True),
        sa.Column("error_message_safe", sa.Text(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("completed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("retention_until", sa.DateTime(timezone=True), nullable=True),
        sa.Column("deleted_at", sa.DateTime(timezone=True), nullable=True),
    )
    op.create_index("ix_comparison_jobs_status", "comparison_jobs", ["status"])
    op.create_index("ix_comparison_jobs_tenant_id", "comparison_jobs", ["tenant_id"])
    op.create_table(
        "comparison_differences",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("job_id", sa.String(64), sa.ForeignKey("comparison_jobs.id", ondelete="CASCADE"), primary_key=True),
        sa.Column("type", sa.String(64), nullable=False),
        sa.Column("severity", sa.String(16), nullable=False),
        sa.Column("sheet_id", sa.String(200), nullable=False),
        sa.Column("cell_ref", sa.String(32), nullable=True),
        sa.Column("canonical_field", sa.String(200), nullable=True),
        sa.Column("business_key_hash", sa.String(64), nullable=True),
        sa.Column("message_safe", sa.Text(), nullable=False),
        sa.Column("render_action", sa.String(64), nullable=False),
        sa.Column("repair_status", sa.String(32), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
    )
    op.create_index("ix_differences_job_type", "comparison_differences", ["job_id", "type"])
    op.create_table(
        "audit_events",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=True),
        sa.Column("actor_id", sa.String(200), nullable=True),
        sa.Column("action", sa.String(100), nullable=False),
        sa.Column("resource_type", sa.String(100), nullable=False),
        sa.Column("resource_id", sa.String(200), nullable=False),
        sa.Column("metadata_json", sa.JSON(), nullable=False),
        sa.Column("occurred_at", sa.DateTime(timezone=True), nullable=False),
    )
    op.create_index("ix_audit_events_action", "audit_events", ["action"])
    op.create_index("ix_audit_events_tenant_id", "audit_events", ["tenant_id"])


def downgrade() -> None:
    op.drop_index("ix_audit_events_tenant_id", table_name="audit_events")
    op.drop_index("ix_audit_events_action", table_name="audit_events")
    op.drop_table("audit_events")
    op.drop_index("ix_differences_job_type", table_name="comparison_differences")
    op.drop_table("comparison_differences")
    op.drop_index("ix_comparison_jobs_status", table_name="comparison_jobs")
    op.drop_index("ix_comparison_jobs_tenant_id", table_name="comparison_jobs")
    op.drop_table("comparison_jobs")
    op.drop_index("ix_standard_snapshots_content_sha256", table_name="standard_snapshots")
    op.drop_table("standard_snapshots")
    op.drop_index("ix_schema_versions_schema_id", table_name="schema_versions")
    op.drop_index("ix_schema_versions_tenant_id", table_name="schema_versions")
    op.drop_table("schema_versions")
