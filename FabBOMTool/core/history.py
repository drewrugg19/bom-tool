"""
core/history.py
Manages run history stored in a local SQLite database.
"""

from contextlib import contextmanager
import sqlite3
import json
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "history.db"


@contextmanager
def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), timeout=30)
    con.row_factory = sqlite3.Row
    try:
        with con:
            yield con
    finally:
        con.close()


def init_db():
    with _conn() as con:
        # Acquire the SQLite writer lock BEFORE inspecting or changing schema.
        # Other processes wait, then inspect the committed schema.
        con.execute("BEGIN IMMEDIATE")
        con.execute("""
            CREATE TABLE IF NOT EXISTS runs (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                run_date      TEXT    NOT NULL,
                export_file   TEXT    NOT NULL,
                pdf_count     INTEGER NOT NULL DEFAULT 0,
                row_count     INTEGER NOT NULL DEFAULT 0,
                total_inches  REAL    NOT NULL DEFAULT 0,
                ok_rows       INTEGER NOT NULL DEFAULT 0,
                warn_rows     INTEGER NOT NULL DEFAULT 0,
                err_rows      INTEGER NOT NULL DEFAULT 0,
                mode          TEXT    NOT NULL DEFAULT 'Company',
                project       TEXT    NOT NULL DEFAULT '',
                summary       TEXT    NOT NULL DEFAULT '',
                output_path   TEXT    NOT NULL DEFAULT ''
            )
        """)
        columns = {row[1] for row in con.execute("PRAGMA table_info(runs)")}
        if "run_type" not in columns:
            con.execute("ALTER TABLE runs ADD COLUMN run_type TEXT NOT NULL DEFAULT 'BOM / Inches'")
        if "pdf_filenames" not in columns:
            con.execute("ALTER TABLE runs ADD COLUMN pdf_filenames TEXT NOT NULL DEFAULT '[]'")
        if "run_metadata" not in columns:
            con.execute("ALTER TABLE runs ADD COLUMN run_metadata TEXT NOT NULL DEFAULT '{}'")
        con.commit()


def save_run(result: dict, pdf_paths: list, mode: str, project: str, run_type: str = "BOM / Inches", run_metadata: dict | None = None) -> int:
    init_db()
    with _conn() as con:
        cur = con.execute("""
            INSERT INTO runs
              (run_date, export_file, pdf_count, row_count, total_inches,
               ok_rows, warn_rows, err_rows, mode, project, summary, output_path, run_type, pdf_filenames, run_metadata)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            result.get("output_filename", ""),
            len(pdf_paths),
            result.get("row_count", len(result["rows"]) if isinstance(result.get("rows"), list) else result.get("rows", 0)),
            result.get("total_inches", 0),
            result.get("ok_rows", 0),
            result.get("warn_rows", 0),
            result.get("err_rows", 0),
            mode,
            project or "",
            result.get("summary", ""),
            result.get("output", ""),
            run_type,
            json.dumps([Path(path).name for path in pdf_paths]),
            json.dumps(run_metadata or {}),
        ))
        con.commit()
        return cur.lastrowid


def get_history(limit: int = 50) -> list:
    init_db()
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM runs ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [_record(r) for r in rows]


def get_run(run_id: int) -> dict | None:
    init_db()
    with _conn() as con:
        row = con.execute("SELECT * FROM runs WHERE id=?", (run_id,)).fetchone()
    return _record(row) if row else None


def delete_run(run_id: int) -> bool:
    init_db()
    with _conn() as con:
        con.execute("DELETE FROM runs WHERE id=?", (run_id,))
        con.commit()
    return True


def _record(row) -> dict:
    record = dict(row)
    record["run_type"] = record.get("run_type") or "BOM / Inches"
    record["pdf_filenames"] = json.loads(record.get("pdf_filenames") or "[]")
    record["run_metadata"] = json.loads(record.get("run_metadata") or "{}")
    return record
