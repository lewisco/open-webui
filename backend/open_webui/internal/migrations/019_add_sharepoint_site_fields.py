"""Peewee migrations -- 019_add_sharepoint_site_fields.py.

Adds sync_all, display_name, and sync_enabled columns to sharepoint_site.
Migrates existing sites where selected_items IS NULL to sync_all = TRUE.

Uses raw SQL because sharepoint_site is created by SQLAlchemy, not peewee,
so peewee_migrate's model registry does not know about it.
"""

from contextlib import suppress

import peewee as pw
from peewee_migrate import Migrator

with suppress(ImportError):
    import playhouse.postgres_ext as pw_pext


def migrate(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your migrations here."""

    # Add columns via raw SQL since the table isn't in peewee's model registry.
    # Use IF NOT EXISTS / IF NOT FOUND guards for idempotency.
    if isinstance(database, pw.SqliteDatabase):
        # SQLite doesn't have IF NOT EXISTS for ALTER TABLE ADD COLUMN,
        # but it will error if the column already exists — catch that.
        for stmt in [
            "ALTER TABLE sharepoint_site ADD COLUMN sync_all BOOLEAN NOT NULL DEFAULT 0",
            "ALTER TABLE sharepoint_site ADD COLUMN display_name VARCHAR(255)",
            "ALTER TABLE sharepoint_site ADD COLUMN sync_enabled BOOLEAN NOT NULL DEFAULT 1",
        ]:
            try:
                database.execute_sql(stmt)
            except Exception:
                pass  # Column likely already exists
    else:
        # PostgreSQL: use DO block for idempotent column addition
        database.execute_sql("""
            DO $$
            BEGIN
                IF NOT EXISTS (
                    SELECT 1 FROM information_schema.columns
                    WHERE table_name = 'sharepoint_site' AND column_name = 'sync_all'
                ) THEN
                    ALTER TABLE sharepoint_site ADD COLUMN sync_all BOOLEAN NOT NULL DEFAULT FALSE;
                END IF;

                IF NOT EXISTS (
                    SELECT 1 FROM information_schema.columns
                    WHERE table_name = 'sharepoint_site' AND column_name = 'display_name'
                ) THEN
                    ALTER TABLE sharepoint_site ADD COLUMN display_name VARCHAR(255);
                END IF;

                IF NOT EXISTS (
                    SELECT 1 FROM information_schema.columns
                    WHERE table_name = 'sharepoint_site' AND column_name = 'sync_enabled'
                ) THEN
                    ALTER TABLE sharepoint_site ADD COLUMN sync_enabled BOOLEAN NOT NULL DEFAULT TRUE;
                END IF;
            END $$;
        """)

    if not fake:
        # Existing sites with selected_items IS NULL were implicitly syncing
        # everything. Preserve that behavior by setting sync_all = TRUE.
        if isinstance(database, pw.SqliteDatabase):
            database.execute_sql(
                "UPDATE sharepoint_site SET sync_all = 1 WHERE selected_items IS NULL"
            )
        else:
            database.execute_sql(
                "UPDATE sharepoint_site SET sync_all = TRUE WHERE selected_items IS NULL"
            )


def rollback(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your rollback migrations here."""

    database.execute_sql("ALTER TABLE sharepoint_site DROP COLUMN IF EXISTS sync_all")
    database.execute_sql("ALTER TABLE sharepoint_site DROP COLUMN IF EXISTS display_name")
    database.execute_sql("ALTER TABLE sharepoint_site DROP COLUMN IF EXISTS sync_enabled")
