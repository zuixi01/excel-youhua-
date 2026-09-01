"""Add versioned product-catalog workflow and human-review persistence."""

from alembic import op
import sqlalchemy as sa


revision = "20260901_0003"
down_revision = "20260828_0002"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "product_catalog_snapshots",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("connection_id", sa.String(200), nullable=False),
        sa.Column("category_id", sa.String(300), nullable=False),
        sa.Column("content_sha256", sa.String(64), nullable=False),
        sa.Column("fields_json", sa.JSON(), nullable=False),
        sa.Column("source_metadata_json", sa.JSON(), nullable=False),
        sa.Column("captured_at", sa.DateTime(timezone=True), nullable=False),
    )
    for column in ("tenant_id", "connection_id", "category_id", "content_sha256"):
        op.create_index(f"ix_product_catalog_snapshots_{column}", "product_catalog_snapshots", [column])
    op.create_table(
        "product_workflow_revisions",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("job_id", sa.String(64), sa.ForeignKey("comparison_jobs.id", ondelete="CASCADE"), nullable=False),
        sa.Column("revision_number", sa.Integer(), nullable=False),
        sa.Column("parent_revision_id", sa.String(64), sa.ForeignKey("product_workflow_revisions.id"), nullable=True),
        sa.Column("status", sa.String(32), nullable=False),
        sa.Column("catalog_snapshot_ids_json", sa.JSON(), nullable=False),
        sa.Column("result_json", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("created_by", sa.String(200), nullable=True),
        sa.UniqueConstraint("job_id", "revision_number", name="uq_product_job_revision"),
    )
    op.create_index("ix_product_workflow_revisions_job_id", "product_workflow_revisions", ["job_id"])
    op.create_index("ix_product_workflow_revisions_status", "product_workflow_revisions", ["status"])
    op.create_table(
        "product_review_items",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("job_id", sa.String(64), sa.ForeignKey("comparison_jobs.id", ondelete="CASCADE"), nullable=False),
        sa.Column("revision_id", sa.String(64), sa.ForeignKey("product_workflow_revisions.id", ondelete="CASCADE"), nullable=False),
        sa.Column("review_key", sa.String(500), nullable=False),
        sa.Column("review_type", sa.String(64), nullable=False),
        sa.Column("status", sa.String(24), nullable=False),
        sa.Column("payload_json", sa.JSON(), nullable=False),
        sa.Column("decision_json", sa.JSON(), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("decided_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("decided_by", sa.String(200), nullable=True),
        sa.UniqueConstraint("revision_id", "review_key", name="uq_product_revision_review_key"),
    )
    for column in ("job_id", "revision_id", "review_type", "status"):
        op.create_index(f"ix_product_review_items_{column}", "product_review_items", [column])


def downgrade() -> None:
    for column in ("status", "review_type", "revision_id", "job_id"):
        op.drop_index(f"ix_product_review_items_{column}", table_name="product_review_items")
    op.drop_table("product_review_items")
    op.drop_index("ix_product_workflow_revisions_status", table_name="product_workflow_revisions")
    op.drop_index("ix_product_workflow_revisions_job_id", table_name="product_workflow_revisions")
    op.drop_table("product_workflow_revisions")
    for column in ("content_sha256", "category_id", "connection_id", "tenant_id"):
        op.drop_index(f"ix_product_catalog_snapshots_{column}", table_name="product_catalog_snapshots")
    op.drop_table("product_catalog_snapshots")
