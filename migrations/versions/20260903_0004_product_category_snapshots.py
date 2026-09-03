"""Persist immutable product category-list snapshots."""

from alembic import op
import sqlalchemy as sa


revision = "20260903_0004"
down_revision = "20260901_0003"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "product_category_snapshots",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("connection_id", sa.String(200), nullable=False),
        sa.Column("content_sha256", sa.String(64), nullable=False),
        sa.Column("categories_json", sa.JSON(), nullable=False),
        sa.Column("source_metadata_json", sa.JSON(), nullable=False),
        sa.Column("captured_at", sa.DateTime(timezone=True), nullable=False),
    )
    for column in ("tenant_id", "connection_id", "content_sha256"):
        op.create_index(
            f"ix_product_category_snapshots_{column}",
            "product_category_snapshots",
            [column],
        )
    # Batch mode keeps this migration portable to the SQLite database used by
    # local development and tests while emitting a normal ALTER on PostgreSQL.
    with op.batch_alter_table("product_workflow_revisions") as batch_op:
        batch_op.add_column(sa.Column("category_snapshot_id", sa.String(64), nullable=True))
        batch_op.create_foreign_key(
            "fk_product_workflow_revisions_category_snapshot_id",
            "product_category_snapshots",
            ["category_snapshot_id"],
            ["id"],
        )


def downgrade() -> None:
    with op.batch_alter_table("product_workflow_revisions") as batch_op:
        batch_op.drop_constraint(
            "fk_product_workflow_revisions_category_snapshot_id",
            type_="foreignkey",
        )
        batch_op.drop_column("category_snapshot_id")
    for column in ("content_sha256", "connection_id", "tenant_id"):
        op.drop_index(
            f"ix_product_category_snapshots_{column}",
            table_name="product_category_snapshots",
        )
    op.drop_table("product_category_snapshots")
