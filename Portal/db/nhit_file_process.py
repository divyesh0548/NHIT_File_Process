"""
Database access for Portal file-process settings (header keywords).

Uses the same RDS database connection as other Portal jobs (RDS_HOST, RDS_PORT,
RDS_DB_NAME, RDS_USER, RDS_PASSWORD) but only reads/writes the `nhit_file_process`
table.

This module intentionally does NOT use RDS_TABLE_NAME (checkpostmaster). That
table is reserved for Valid/Invalid Lookup vehicle master data only.
"""

from pathlib import Path
import os
import re

from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.exc import IntegrityError

REPO_ROOT = Path(__file__).resolve().parents[2]
load_dotenv(REPO_ROOT / ".env")

LC_ETC_FILE_TYPE = "LC/ETC"
# Fixed table for LC/ETC header keywords — never RDS_TABLE_NAME / checkpostmaster.
NHIT_FILE_PROCESS_TABLE = "nhit_file_process"
_RESERVED_LOOKUP_TABLE = "checkpostmaster"
_TABLE_NAME_PATTERN = re.compile(r"^[a-z_][a-z0-9_]*$")

DEFAULT_LC_ETC_HEADER_KEYWORDS = [
    "Agency Txn Id",
    "Settlement Amount",
    "Plaza ID",
    "Violation Amts",
]


def _assert_keywords_table_isolated():
    """Guard against ever pointing keyword storage at the lookup master table."""
    table_name = NHIT_FILE_PROCESS_TABLE.lower()
    if table_name == _RESERVED_LOOKUP_TABLE:
        raise RuntimeError(
            "Header keyword storage cannot use the checkpostmaster table. "
            "Keywords are stored only in nhit_file_process."
        )
    if not _TABLE_NAME_PATTERN.match(table_name):
        raise RuntimeError(f"Invalid keyword table name: {NHIT_FILE_PROCESS_TABLE}")


def _db_config():
    return {
        "host": os.environ.get("RDS_HOST"),
        "port": os.environ.get("RDS_PORT", "5432"),
        "name": os.environ.get("RDS_DB_NAME"),
        "user": os.environ.get("RDS_USER"),
        "password": os.environ.get("RDS_PASSWORD"),
    }


def _missing_db_config():
    cfg = _db_config()
    missing = [key for key, value in cfg.items() if key != "port" and not value]
    return missing


def get_engine():
    _assert_keywords_table_isolated()
    missing = _missing_db_config()
    if missing:
        raise EnvironmentError(
            "Missing required database configuration in .env: " + ", ".join(missing)
        )
    cfg = _db_config()
    return create_engine(
        f"postgresql+psycopg2://{cfg['user']}:{cfg['password']}@{cfg['host']}:{cfg['port']}/{cfg['name']}"
    )


def ensure_table_exists(engine=None):
    _assert_keywords_table_isolated()
    engine = engine or get_engine()
    create_sql = f"""
        CREATE TABLE IF NOT EXISTS {NHIT_FILE_PROCESS_TABLE} (
            id SERIAL PRIMARY KEY,
            file_type TEXT NOT NULL,
            header_keywords TEXT NOT NULL,
            CONSTRAINT uq_nhit_file_process_type_keyword UNIQUE (file_type, header_keywords)
        )
    """
    with engine.begin() as conn:
        conn.execute(text(create_sql))


def list_header_keyword_records(file_type=LC_ETC_FILE_TYPE, engine=None):
    engine = engine or get_engine()
    ensure_table_exists(engine)
    query = text(
        f"""
        SELECT id, file_type, header_keywords
        FROM {NHIT_FILE_PROCESS_TABLE}
        WHERE file_type = :file_type
        ORDER BY id
        """
    )
    with engine.connect() as conn:
        rows = conn.execute(query, {"file_type": file_type}).mappings().all()
    return [dict(row) for row in rows]


def list_header_keyword_strings(file_type=LC_ETC_FILE_TYPE, engine=None):
    return [row["header_keywords"] for row in list_header_keyword_records(file_type, engine=engine)]


def ensure_lc_etc_keywords_seeded(engine=None):
    engine = engine or get_engine()
    ensure_table_exists(engine)
    existing = list_header_keyword_strings(LC_ETC_FILE_TYPE, engine=engine)
    if existing:
        return existing
    for keyword in DEFAULT_LC_ETC_HEADER_KEYWORDS:
        add_header_keyword(LC_ETC_FILE_TYPE, keyword, engine=engine)
    return list_header_keyword_strings(LC_ETC_FILE_TYPE, engine=engine)


def get_lc_etc_header_keyword_records():
    engine = get_engine()
    ensure_lc_etc_keywords_seeded(engine)
    return list_header_keyword_records(LC_ETC_FILE_TYPE, engine=engine)


def get_lc_etc_header_keyword_strings():
    engine = get_engine()
    return ensure_lc_etc_keywords_seeded(engine)


def add_header_keyword(file_type, keyword, engine=None):
    keyword = (keyword or "").strip()
    if not keyword:
        raise ValueError("Header keyword cannot be empty.")
    engine = engine or get_engine()
    ensure_table_exists(engine)
    insert_sql = text(
        f"""
        INSERT INTO {NHIT_FILE_PROCESS_TABLE} (file_type, header_keywords)
        VALUES (:file_type, :header_keywords)
        RETURNING id, file_type, header_keywords
        """
    )
    try:
        with engine.begin() as conn:
            row = conn.execute(
                insert_sql,
                {"file_type": file_type, "header_keywords": keyword},
            ).mappings().one()
    except IntegrityError as exc:
        raise ValueError(f"Keyword already exists: {keyword}") from exc
    return dict(row)


def update_header_keyword(keyword_id, keyword, file_type=LC_ETC_FILE_TYPE, engine=None):
    keyword = (keyword or "").strip()
    if not keyword:
        raise ValueError("Header keyword cannot be empty.")
    engine = engine or get_engine()
    ensure_table_exists(engine)
    update_sql = text(
        f"""
        UPDATE {NHIT_FILE_PROCESS_TABLE}
        SET header_keywords = :header_keywords
        WHERE id = :keyword_id AND file_type = :file_type
        RETURNING id, file_type, header_keywords
        """
    )
    try:
        with engine.begin() as conn:
            row = conn.execute(
                update_sql,
                {
                    "keyword_id": keyword_id,
                    "file_type": file_type,
                    "header_keywords": keyword,
                },
            ).mappings().first()
    except IntegrityError as exc:
        raise ValueError(f"Keyword already exists: {keyword}") from exc
    if not row:
        raise ValueError("Header keyword not found.")
    return dict(row)


def delete_header_keyword(keyword_id, file_type=LC_ETC_FILE_TYPE, engine=None):
    engine = engine or get_engine()
    ensure_table_exists(engine)
    delete_sql = text(
        f"""
        DELETE FROM {NHIT_FILE_PROCESS_TABLE}
        WHERE id = :keyword_id AND file_type = :file_type
        RETURNING id
        """
    )
    with engine.begin() as conn:
        row = conn.execute(
            delete_sql,
            {"keyword_id": keyword_id, "file_type": file_type},
        ).first()
    if not row:
        raise ValueError("Header keyword not found.")
    return True
