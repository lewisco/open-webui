"""Add sharepoint_site and sharepoint_file tables

Revision ID: c3d4e5f6a7b8
Revises: b2c3d4e5f6a7
Create Date: 2026-02-15 00:00:00.000000

"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
import json

from open_webui.migrations.util import get_existing_tables

revision: str = "c3d4e5f6a7b8"
down_revision: Union[str, None] = "b2c3d4e5f6a7"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def _convert_column_to_json(table: str, column: str):
    conn = op.get_bind()
    dialect = conn.dialect.name

    if dialect == "sqlite":
        op.add_column(table, sa.Column(f"{column}_json", sa.JSON(), nullable=True))

        rows = conn.execute(sa.text(f'SELECT id, {column} FROM "{table}"')).fetchall()

        for row in rows:
            uid, raw = row
            if raw is None:
                parsed = None
            else:
                try:
                    parsed = json.loads(raw)
                except Exception:
                    parsed = None

            conn.execute(
                sa.text(f'UPDATE "{table}" SET {column}_json = :val WHERE id = :id'),
                {"val": json.dumps(parsed) if parsed else None, "id": uid},
            )

        op.drop_column(table, column)
        op.alter_column(table, f"{column}_json", new_column_name=column)

    else:
        op.alter_column(
            table,
            column,
            type_=sa.JSON(),
            postgresql_using=f"{column}::json",
        )


def upgrade() -> None:
    existing_tables = set(get_existing_tables())

    if "sharepoint_site" not in existing_tables:
        op.create_table(
            "sharepoint_site",
            sa.Column("id", sa.Text(), primary_key=True),
            sa.Column("site_id", sa.Text(), nullable=False),
            sa.Column("drive_id", sa.Text(), nullable=False),
            sa.Column("site_url", sa.Text(), nullable=True),
            sa.Column("site_name", sa.Text(), nullable=True),
            sa.Column("drive_name", sa.Text(), nullable=True),
            sa.Column("selected_items", sa.JSON(), nullable=True),
            sa.Column("kb_id", sa.Text(), nullable=True),
            sa.Column("kb_name", sa.Text(), nullable=True),
            sa.Column("sync_mode", sa.Text(), server_default="none"),
            sa.Column("delta_link", sa.Text(), nullable=True),
            sa.Column("last_sync_at", sa.BigInteger(), nullable=True),
            sa.Column("sync_status", sa.Text(), server_default="idle"),
            sa.Column("sync_error", sa.Text(), nullable=True),
            sa.Column("created_at", sa.BigInteger(), nullable=False),
            sa.Column("updated_at", sa.BigInteger(), nullable=False),
        )
    else:
        # Peewee creates selected_items as TEXT; convert to native JSON
        _convert_column_to_json("sharepoint_site", "selected_items")

    if "sharepoint_file" not in existing_tables:
        op.create_table(
            "sharepoint_file",
            sa.Column("id", sa.Text(), primary_key=True),
            sa.Column(
                "site_config_id",
                sa.Text(),
                sa.ForeignKey("sharepoint_site.id", ondelete="CASCADE"),
                nullable=False,
            ),
            sa.Column("sp_item_id", sa.Text(), nullable=False),
            sa.Column("owui_file_id", sa.Text(), nullable=True),
            sa.Column("filename", sa.Text(), nullable=True),
            sa.Column("sp_item_path", sa.Text(), nullable=True),
            sa.Column("sp_etag", sa.Text(), nullable=True),
            sa.Column("sp_last_modified", sa.Text(), nullable=True),
            sa.Column("allowed_users", sa.JSON(), nullable=True),
            sa.Column("allowed_groups", sa.JSON(), nullable=True),
            sa.Column("sync_status", sa.Text(), server_default="pending"),
            sa.Column("sync_error", sa.Text(), nullable=True),
            sa.Column("excluded", sa.Boolean(), server_default="0"),
            sa.Column("created_at", sa.BigInteger(), nullable=False),
            sa.Column("updated_at", sa.BigInteger(), nullable=False),
            sa.UniqueConstraint(
                "site_config_id",
                "sp_item_id",
                name="uq_sharepoint_file_site_item",
            ),
        )

        op.create_index(
            "ix_sp_file_owui_id",
            "sharepoint_file",
            ["owui_file_id"],
        )
        op.create_index(
            "ix_sp_file_site_config",
            "sharepoint_file",
            ["site_config_id"],
        )
    else:
        # Peewee creates these as TEXT; convert to native JSON
        _convert_column_to_json("sharepoint_file", "allowed_users")
        _convert_column_to_json("sharepoint_file", "allowed_groups")


def downgrade() -> None:
    op.drop_index("ix_sp_file_site_config", table_name="sharepoint_file")
    op.drop_index("ix_sp_file_owui_id", table_name="sharepoint_file")
    op.drop_table("sharepoint_file")
    op.drop_table("sharepoint_site")
