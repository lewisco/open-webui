import logging
import time
from typing import Optional
import uuid

from sqlalchemy import func
from sqlalchemy.orm import Session
from open_webui.internal.db import Base, get_db_context

from pydantic import BaseModel, ConfigDict
from sqlalchemy import (
    BigInteger,
    Boolean,
    Column,
    ForeignKey,
    Text,
    JSON,
    UniqueConstraint,
    or_,
)

log = logging.getLogger(__name__)

####################
# SharePointSite DB Schema
####################


class SharePointSite(Base):
    __tablename__ = "sharepoint_site"

    id = Column(Text, unique=True, primary_key=True)

    site_id = Column(Text, nullable=False)
    drive_id = Column(Text, nullable=False)
    site_url = Column(Text, nullable=True)
    site_name = Column(Text, nullable=True)
    drive_name = Column(Text, nullable=True)

    selected_items = Column(JSON, nullable=True)

    kb_id = Column(Text, nullable=True)
    kb_name = Column(Text, nullable=True)
    sync_mode = Column(Text, default="none")

    delta_link = Column(Text, nullable=True)
    last_sync_at = Column(BigInteger, nullable=True)
    sync_status = Column(Text, default="idle")
    sync_error = Column(Text, nullable=True)

    created_at = Column(BigInteger)
    updated_at = Column(BigInteger)


class SharePointSiteModel(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    site_id: str
    drive_id: str
    site_url: Optional[str] = None
    site_name: Optional[str] = None
    drive_name: Optional[str] = None

    selected_items: Optional[list] = None

    kb_id: Optional[str] = None
    kb_name: Optional[str] = None
    sync_mode: str = "none"

    delta_link: Optional[str] = None
    last_sync_at: Optional[int] = None
    sync_status: str = "idle"
    sync_error: Optional[str] = None
    file_count: int = 0
    error_count: int = 0
    excluded_count: int = 0

    created_at: int
    updated_at: int


####################
# SharePointFile DB Schema
####################


class SharePointFile(Base):
    __tablename__ = "sharepoint_file"

    id = Column(Text, unique=True, primary_key=True)

    site_config_id = Column(
        Text,
        ForeignKey("sharepoint_site.id", ondelete="CASCADE"),
        nullable=False,
    )
    sp_item_id = Column(Text, nullable=False)
    owui_file_id = Column(Text, nullable=True)
    filename = Column(Text, nullable=True)
    sp_item_path = Column(Text, nullable=True)
    sp_etag = Column(Text, nullable=True)
    sp_last_modified = Column(Text, nullable=True)

    allowed_users = Column(JSON, nullable=True)
    allowed_groups = Column(JSON, nullable=True)

    sync_status = Column(Text, default="pending")
    sync_error = Column(Text, nullable=True)
    excluded = Column(Boolean, default=False)

    created_at = Column(BigInteger)
    updated_at = Column(BigInteger)

    __table_args__ = (
        UniqueConstraint(
            "site_config_id",
            "sp_item_id",
            name="uq_sharepoint_file_site_item",
        ),
    )


class SharePointFileModel(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    site_config_id: str
    sp_item_id: str
    owui_file_id: Optional[str] = None
    filename: Optional[str] = None
    sp_item_path: Optional[str] = None
    sp_etag: Optional[str] = None
    sp_last_modified: Optional[str] = None

    allowed_users: Optional[list] = None
    allowed_groups: Optional[list] = None

    sync_status: str = "pending"
    sync_error: Optional[str] = None
    excluded: bool = False

    created_at: int
    updated_at: int


####################
# Forms
####################


class SharePointSiteForm(BaseModel):
    site_url: str
    site_id: str
    site_name: Optional[str] = None
    drive_id: str
    drive_name: Optional[str] = None
    selected_items: Optional[list] = None
    kb_name: Optional[str] = None
    sync_mode: str = "none"


class SharePointConfigForm(BaseModel):
    ENABLE_SHAREPOINT_SYNC: Optional[bool] = None
    SHAREPOINT_TENANT_ID: Optional[str] = None
    SHAREPOINT_CLIENT_ID: Optional[str] = None
    SHAREPOINT_CLIENT_SECRET: Optional[str] = None
    SHAREPOINT_SYNC_INTERVAL: Optional[int] = None


class SharePointSiteUpdateForm(BaseModel):
    selected_items: Optional[list] = None
    sync_mode: Optional[str] = None
    kb_name: Optional[str] = None


class SharePointSyncForm(BaseModel):
    site_id: Optional[str] = None
    force: bool = False
    clear_exclusions: bool = False


####################
# SharePointTable Operations
####################


class SharePointTable:
    # ---- Site CRUD ----

    def insert_new_site(
        self, form_data: SharePointSiteForm, kb_id: str, db: Optional[Session] = None
    ) -> Optional[SharePointSiteModel]:
        with get_db_context(db) as db:
            site = SharePointSite(
                id=str(uuid.uuid4()),
                site_id=form_data.site_id,
                drive_id=form_data.drive_id,
                site_url=form_data.site_url,
                site_name=form_data.site_name,
                drive_name=form_data.drive_name,
                selected_items=form_data.selected_items,
                kb_id=kb_id,
                kb_name=form_data.kb_name,
                sync_mode=form_data.sync_mode,
                sync_status="idle",
                created_at=int(time.time()),
                updated_at=int(time.time()),
            )
            try:
                db.add(site)
                db.commit()
                db.refresh(site)
                return SharePointSiteModel.model_validate(site)
            except Exception as e:
                log.exception(e)
                return None

    def get_sites(self, db: Optional[Session] = None) -> list[SharePointSiteModel]:
        with get_db_context(db) as db:
            sites = (
                db.query(SharePointSite)
                .order_by(SharePointSite.created_at.desc())
                .all()
            )
            return [SharePointSiteModel.model_validate(s) for s in sites]

    def get_site_by_id(
        self, id: str, db: Optional[Session] = None
    ) -> Optional[SharePointSiteModel]:
        try:
            with get_db_context(db) as db:
                site = db.query(SharePointSite).filter_by(id=id).first()
                return SharePointSiteModel.model_validate(site) if site else None
        except Exception as e:
            log.debug(f"get_site_by_id error: {e}")
            return None

    def update_site(
        self, id: str, update_data: dict, db: Optional[Session] = None
    ) -> Optional[SharePointSiteModel]:
        try:
            with get_db_context(db) as db:
                db.query(SharePointSite).filter_by(id=id).update(
                    {**update_data, "updated_at": int(time.time())}
                )
                db.commit()
                return self.get_site_by_id(id, db=db)
        except Exception as e:
            log.exception(e)
            return None

    def delete_site_by_id(self, id: str, db: Optional[Session] = None) -> bool:
        try:
            with get_db_context(db) as db:
                # Files cascade-delete via FK
                db.query(SharePointSite).filter_by(id=id).delete()
                db.commit()
                return True
        except Exception as e:
            log.exception(e)
            return False

    # ---- File CRUD ----

    def upsert_file(
        self,
        site_config_id: str,
        sp_item_id: str,
        data: dict,
        db: Optional[Session] = None,
    ) -> Optional[SharePointFileModel]:
        with get_db_context(db) as db:
            existing = (
                db.query(SharePointFile)
                .filter_by(site_config_id=site_config_id, sp_item_id=sp_item_id)
                .first()
            )
            if existing:
                for key, value in data.items():
                    setattr(existing, key, value)
                existing.updated_at = int(time.time())
                db.commit()
                db.refresh(existing)
                return SharePointFileModel.model_validate(existing)
            else:
                file = SharePointFile(
                    id=str(uuid.uuid4()),
                    site_config_id=site_config_id,
                    sp_item_id=sp_item_id,
                    **data,
                    created_at=int(time.time()),
                    updated_at=int(time.time()),
                )
                db.add(file)
                db.commit()
                db.refresh(file)
                return SharePointFileModel.model_validate(file)

    def get_files_by_site(
        self,
        site_config_id: str,
        skip: int = 0,
        limit: int = 200,
        db: Optional[Session] = None,
    ) -> list[SharePointFileModel]:
        with get_db_context(db) as db:
            query = (
                db.query(SharePointFile)
                .filter_by(site_config_id=site_config_id)
                .order_by(SharePointFile.filename)
                .offset(skip)
                .limit(min(limit, 1000))
            )
            return [SharePointFileModel.model_validate(f) for f in query.all()]

    def get_file_by_sp_item_id(
        self, site_config_id: str, sp_item_id: str, db: Optional[Session] = None
    ) -> Optional[SharePointFileModel]:
        try:
            with get_db_context(db) as db:
                f = (
                    db.query(SharePointFile)
                    .filter_by(site_config_id=site_config_id, sp_item_id=sp_item_id)
                    .first()
                )
                return SharePointFileModel.model_validate(f) if f else None
        except Exception as e:
            log.debug(f"get_file_by_sp_item_id error: {e}")
            return None

    def get_file_by_owui_id(
        self, owui_file_id: str, db: Optional[Session] = None
    ) -> Optional[SharePointFileModel]:
        try:
            with get_db_context(db) as db:
                f = (
                    db.query(SharePointFile)
                    .filter_by(owui_file_id=owui_file_id)
                    .first()
                )
                return SharePointFileModel.model_validate(f) if f else None
        except Exception as e:
            log.debug(f"get_file_by_owui_id error: {e}")
            return None

    def get_files_by_owui_ids(
        self, owui_file_ids: list[str], db: Optional[Session] = None
    ) -> list[SharePointFileModel]:
        """Batch lookup SharePoint files by their Open WebUI file IDs."""
        if not owui_file_ids:
            return []
        try:
            with get_db_context(db) as db:
                files = (
                    db.query(SharePointFile)
                    .filter(SharePointFile.owui_file_id.in_(owui_file_ids))
                    .all()
                )
                return [SharePointFileModel.model_validate(f) for f in files]
        except Exception as e:
            log.debug(f"get_files_by_owui_ids error: {e}")
            return []

    def get_sites_by_ids(
        self, site_ids: list[str], db: Optional[Session] = None
    ) -> list[SharePointSiteModel]:
        """Batch lookup SharePoint sites by IDs."""
        if not site_ids:
            return []
        try:
            with get_db_context(db) as db:
                sites = (
                    db.query(SharePointSite)
                    .filter(SharePointSite.id.in_(site_ids))
                    .all()
                )
                return [SharePointSiteModel.model_validate(s) for s in sites]
        except Exception as e:
            log.debug(f"get_sites_by_ids error: {e}")
            return []

    def get_excluded_files_by_site(
        self, site_config_id: str, db: Optional[Session] = None
    ) -> list[SharePointFileModel]:
        with get_db_context(db) as db:
            files = (
                db.query(SharePointFile)
                .filter_by(site_config_id=site_config_id, excluded=True)
                .all()
            )
            return [SharePointFileModel.model_validate(f) for f in files]

    def mark_file_excluded(
        self, file_id: str, db: Optional[Session] = None
    ) -> bool:
        try:
            with get_db_context(db) as db:
                db.query(SharePointFile).filter_by(id=file_id).update(
                    {"excluded": True, "updated_at": int(time.time())}
                )
                db.commit()
                return True
        except Exception as e:
            log.debug(f"mark_file_excluded error: {e}")
            return False

    def mark_file_included(
        self, file_id: str, db: Optional[Session] = None
    ) -> bool:
        try:
            with get_db_context(db) as db:
                db.query(SharePointFile).filter_by(id=file_id).update(
                    {"excluded": False, "updated_at": int(time.time())}
                )
                db.commit()
                return True
        except Exception as e:
            log.debug(f"mark_file_included error: {e}")
            return False

    def clear_exclusions_by_site(
        self, site_config_id: str, db: Optional[Session] = None
    ) -> bool:
        try:
            with get_db_context(db) as db:
                db.query(SharePointFile).filter_by(
                    site_config_id=site_config_id, excluded=True
                ).update({"excluded": False, "updated_at": int(time.time())})
                db.commit()
                return True
        except Exception as e:
            log.debug(f"clear_exclusions_by_site error: {e}")
            return False

    def get_file_counts_by_site(
        self, site_config_id: str, db: Optional[Session] = None
    ) -> dict:
        """Return file_count, error_count, and excluded_count in a single query."""
        try:
            with get_db_context(db) as db:
                result = (
                    db.query(
                        func.count()
                        .filter(SharePointFile.sync_status == "synced")
                        .label("file_count"),
                        func.count()
                        .filter(SharePointFile.sync_status == "error")
                        .label("error_count"),
                        func.count()
                        .filter(SharePointFile.excluded == True)
                        .label("excluded_count"),
                    )
                    .filter(SharePointFile.site_config_id == site_config_id)
                    .one()
                )
                return {
                    "file_count": result.file_count,
                    "error_count": result.error_count,
                    "excluded_count": result.excluded_count,
                }
        except Exception as e:
            log.debug(f"get_file_counts_by_site error: {e}")
            return {"file_count": 0, "error_count": 0, "excluded_count": 0}

    def get_all_file_counts(
        self, db: Optional[Session] = None
    ) -> dict[str, dict]:
        """Return file counts grouped by site_config_id in a single query."""
        try:
            with get_db_context(db) as db:
                results = (
                    db.query(
                        SharePointFile.site_config_id,
                        func.count()
                        .filter(SharePointFile.sync_status == "synced")
                        .label("file_count"),
                        func.count()
                        .filter(SharePointFile.sync_status == "error")
                        .label("error_count"),
                        func.count()
                        .filter(SharePointFile.excluded == True)
                        .label("excluded_count"),
                    )
                    .group_by(SharePointFile.site_config_id)
                    .all()
                )
                return {
                    r.site_config_id: {
                        "file_count": r.file_count,
                        "error_count": r.error_count,
                        "excluded_count": r.excluded_count,
                    }
                    for r in results
                }
        except Exception as e:
            log.debug(f"get_all_file_counts error: {e}")
            return {}

    def delete_file_by_sp_item_id(
        self, site_config_id: str, sp_item_id: str, db: Optional[Session] = None
    ) -> bool:
        try:
            with get_db_context(db) as db:
                db.query(SharePointFile).filter_by(
                    site_config_id=site_config_id, sp_item_id=sp_item_id
                ).delete()
                db.commit()
                return True
        except Exception as e:
            log.debug(f"delete_file_by_sp_item_id error: {e}")
            return False

    def get_problem_files_by_site(
        self, site_config_id: str, db: Optional[Session] = None
    ) -> list[SharePointFileModel]:
        with get_db_context(db) as db:
            files = (
                db.query(SharePointFile)
                .filter(
                    SharePointFile.site_config_id == site_config_id,
                    or_(
                        SharePointFile.sync_status == "error",
                        SharePointFile.excluded == True,
                    ),
                )
                .order_by(SharePointFile.filename)
                .all()
            )
            return [SharePointFileModel.model_validate(f) for f in files]

    def get_error_files_by_site(
        self, site_config_id: str, db: Optional[Session] = None
    ) -> list[SharePointFileModel]:
        with get_db_context(db) as db:
            files = (
                db.query(SharePointFile)
                .filter_by(site_config_id=site_config_id, sync_status="error")
                .all()
            )
            return [SharePointFileModel.model_validate(f) for f in files]


SharePoints = SharePointTable()
