"""Peewee migrations -- 019_add_sharepoint_site_fields.py.

Adds sync_all, display_name, and sync_enabled columns to sharepoint_site.
Migrates existing sites where selected_items IS NULL to sync_all = TRUE.
"""

from contextlib import suppress

import peewee as pw
from peewee_migrate import Migrator

with suppress(ImportError):
    import playhouse.postgres_ext as pw_pext


def migrate(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your migrations here."""

    migrator.add_fields(
        "sharepoint_site",
        sync_all=pw.BooleanField(default=False),
        display_name=pw.CharField(max_length=255, null=True),
        sync_enabled=pw.BooleanField(default=True),
    )

    if not fake:
        # Existing sites with selected_items IS NULL were implicitly syncing
        # everything. Preserve that behavior by setting sync_all = TRUE.
        database.execute_sql(
            "UPDATE sharepoint_site SET sync_all = ? WHERE selected_items IS NULL",
            (True,),
        )


def rollback(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your rollback migrations here."""

    migrator.remove_fields("sharepoint_site", "sync_all", "display_name", "sync_enabled")
