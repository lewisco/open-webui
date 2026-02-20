"""Peewee migrations -- 019_add_sharepoint_site_fields.py.

Some examples (model - class or model name)::

    > Model = migrator.orm['table_name']            # Return model in current state by name
    > Model = migrator.ModelClass                   # Return model in current state by name

    > migrator.sql(sql)                             # Run custom SQL
    > migrator.run(func, *args, **kwargs)           # Run python function with the given args
    > migrator.create_model(Model)                  # Create a model (could be used as decorator)
    > migrator.remove_model(model, cascade=True)    # Remove a model
    > migrator.add_fields(model, **fields)          # Add fields to a model
    > migrator.change_fields(model, **fields)       # Change fields
    > migrator.remove_fields(model, *field_names, cascade=True)
    > migrator.rename_field(model, old_field_name, new_field_name)
    > migrator.rename_table(model, new_table_name)
    > migrator.add_index(model, *col_names, unique=False)
    > migrator.add_not_null(model, *field_names)
    > migrator.add_default(model, field_name, default)
    > migrator.add_constraint(model, name, sql)
    > migrator.drop_index(model, *col_names)
    > migrator.drop_not_null(model, *field_names)
    > migrator.drop_constraints(model, *constraints)

"""

from contextlib import suppress

import peewee as pw
from peewee_migrate import Migrator

with suppress(ImportError):
    import playhouse.postgres_ext as pw_pext


def migrate(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your migrations here."""

    if isinstance(database, pw.SqliteDatabase):
        migrate_sqlite(migrator, database, fake=fake)
    else:
        migrate_external(migrator, database, fake=fake)


def migrate_sqlite(migrator: Migrator, database: pw.Database, *, fake=False):
    @migrator.create_model
    class SharePointSite(pw.Model):
        id = pw.CharField(max_length=255, unique=True)
        site_id = pw.CharField(max_length=255)
        drive_id = pw.CharField(max_length=255)
        site_url = pw.CharField(max_length=255, null=True)
        site_name = pw.CharField(max_length=255, null=True)
        drive_name = pw.CharField(max_length=255, null=True)
        selected_items = pw.TextField(null=True)
        sync_all = pw.BooleanField(default=False)
        display_name = pw.CharField(max_length=255, null=True)
        sync_enabled = pw.BooleanField(default=True)
        kb_id = pw.CharField(max_length=255, null=True)
        kb_name = pw.CharField(max_length=255, null=True)
        sync_mode = pw.CharField(max_length=255, default="none")
        delta_link = pw.TextField(null=True)
        last_sync_at = pw.BigIntegerField(null=True)
        sync_status = pw.CharField(max_length=255, default="idle")
        sync_error = pw.TextField(null=True)
        created_at = pw.BigIntegerField()
        updated_at = pw.BigIntegerField()

        class Meta:
            table_name = "sharepoint_site"

    @migrator.create_model
    class SharePointFile(pw.Model):
        id = pw.CharField(max_length=255, unique=True)
        site_config_id = pw.ForeignKeyField(
            SharePointSite, field=SharePointSite.id, on_delete="CASCADE"
        )
        sp_item_id = pw.CharField(max_length=255)
        owui_file_id = pw.CharField(max_length=255, null=True)
        filename = pw.CharField(max_length=255, null=True)
        sp_item_path = pw.TextField(null=True)
        sp_etag = pw.CharField(max_length=255, null=True)
        sp_last_modified = pw.CharField(max_length=255, null=True)
        allowed_users = pw.TextField(null=True)
        allowed_groups = pw.TextField(null=True)
        sync_status = pw.CharField(max_length=255, default="pending")
        sync_error = pw.TextField(null=True)
        excluded = pw.BooleanField(default=False)
        created_at = pw.BigIntegerField()
        updated_at = pw.BigIntegerField()

        class Meta:
            table_name = "sharepoint_file"


def migrate_external(migrator: Migrator, database: pw.Database, *, fake=False):
    @migrator.create_model
    class SharePointSite(pw.Model):
        id = pw.TextField(unique=True)
        site_id = pw.TextField()
        drive_id = pw.TextField()
        site_url = pw.TextField(null=True)
        site_name = pw.TextField(null=True)
        drive_name = pw.TextField(null=True)
        selected_items = pw.TextField(null=True)
        sync_all = pw.BooleanField(default=False)
        display_name = pw.TextField(null=True)
        sync_enabled = pw.BooleanField(default=True)
        kb_id = pw.TextField(null=True)
        kb_name = pw.TextField(null=True)
        sync_mode = pw.TextField(default="none")
        delta_link = pw.TextField(null=True)
        last_sync_at = pw.BigIntegerField(null=True)
        sync_status = pw.TextField(default="idle")
        sync_error = pw.TextField(null=True)
        created_at = pw.BigIntegerField()
        updated_at = pw.BigIntegerField()

        class Meta:
            table_name = "sharepoint_site"

    @migrator.create_model
    class SharePointFile(pw.Model):
        id = pw.TextField(unique=True)
        site_config_id = pw.ForeignKeyField(
            SharePointSite, field=SharePointSite.id, on_delete="CASCADE"
        )
        sp_item_id = pw.TextField()
        owui_file_id = pw.TextField(null=True)
        filename = pw.TextField(null=True)
        sp_item_path = pw.TextField(null=True)
        sp_etag = pw.TextField(null=True)
        sp_last_modified = pw.TextField(null=True)
        allowed_users = pw.TextField(null=True)
        allowed_groups = pw.TextField(null=True)
        sync_status = pw.TextField(default="pending")
        sync_error = pw.TextField(null=True)
        excluded = pw.BooleanField(default=False)
        created_at = pw.BigIntegerField()
        updated_at = pw.BigIntegerField()

        class Meta:
            table_name = "sharepoint_file"


def rollback(migrator: Migrator, database: pw.Database, *, fake=False):
    """Write your rollback migrations here."""

    migrator.remove_model("sharepoint_file")
    migrator.remove_model("sharepoint_site")
