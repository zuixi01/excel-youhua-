"""Add schema catalog, managed connection registry, and lifecycle metadata."""

from alembic import op
import sqlalchemy as sa


revision = "20260828_0002"
down_revision = "20260828_0001"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "schemas",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("schema_id", sa.String(200), nullable=False),
        sa.Column("name", sa.String(300), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("current_version", sa.String(32), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("tenant_id", "schema_id", name="uq_tenant_schema"),
    )
    op.create_index("ix_schemas_tenant_id", "schemas", ["tenant_id"])
    op.create_index("ix_schemas_schema_id", "schemas", ["schema_id"])
    op.create_table(
        "standard_connections",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("tenant_id", sa.String(200), nullable=False),
        sa.Column("connection_id", sa.String(200), nullable=False),
        sa.Column("source_type", sa.String(32), nullable=False),
        sa.Column("config_ciphertext", sa.Text(), nullable=False),
        sa.Column("secret_ref", sa.String(300), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.UniqueConstraint("tenant_id", "connection_id", name="uq_tenant_standard_connection"),
    )
    op.create_index("ix_standard_connections_tenant_id", "standard_connections", ["tenant_id"])
    op.create_index("ix_standard_connections_connection_id", "standard_connections", ["connection_id"])
    with op.batch_alter_table("schema_versions") as batch:
        batch.add_column(sa.Column("created_by", sa.String(200), nullable=True))
        batch.add_column(sa.Column("change_summary", sa.Text(), nullable=True))
    with op.batch_alter_table("standard_snapshots") as batch:
        batch.add_column(sa.Column("tenant_id", sa.String(200), nullable=False, server_default="local"))
        batch.add_column(sa.Column("schema_id", sa.String(200), nullable=False, server_default=""))
        batch.add_column(sa.Column("schema_version", sa.String(32), nullable=False, server_default=""))
        batch.add_column(sa.Column("status", sa.String(24), nullable=False, server_default="ready"))
        batch.add_column(sa.Column("expires_at", sa.DateTime(timezone=True), nullable=True))
        batch.create_index("ix_standard_snapshots_tenant_id", ["tenant_id"])
        batch.create_index("ix_standard_snapshots_schema_id", ["schema_id"])
    with op.batch_alter_table("comparison_jobs") as batch:
        batch.add_column(sa.Column("parameters_json", sa.JSON(), nullable=False, server_default="{}"))
        batch.add_column(sa.Column("input_object_key", sa.Text(), nullable=True))
        batch.add_column(sa.Column("output_object_key", sa.Text(), nullable=True))
        batch.add_column(sa.Column("report_object_keys", sa.JSON(), nullable=False, server_default="{}"))
        batch.add_column(sa.Column("renderer_version", sa.String(100), nullable=True))
        batch.add_column(sa.Column("started_at", sa.DateTime(timezone=True), nullable=True))


def downgrade() -> None:
    with op.batch_alter_table("comparison_jobs") as batch:
        for name in ("started_at", "renderer_version", "report_object_keys", "output_object_key", "input_object_key", "parameters_json"):
            batch.drop_column(name)
    with op.batch_alter_table("standard_snapshots") as batch:
        batch.drop_index("ix_standard_snapshots_schema_id")
        batch.drop_index("ix_standard_snapshots_tenant_id")
        for name in ("expires_at", "status", "schema_version", "schema_id", "tenant_id"):
            batch.drop_column(name)
    with op.batch_alter_table("schema_versions") as batch:
        batch.drop_column("change_summary")
        batch.drop_column("created_by")
    op.drop_index("ix_standard_connections_connection_id", table_name="standard_connections")
    op.drop_index("ix_standard_connections_tenant_id", table_name="standard_connections")
    op.drop_table("standard_connections")
    op.drop_index("ix_schemas_schema_id", table_name="schemas")
    op.drop_index("ix_schemas_tenant_id", table_name="schemas")
    op.drop_table("schemas")
