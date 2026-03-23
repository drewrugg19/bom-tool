"""
core/history.py
Manages run history stored in a local SQLite database.
"""

import sqlite3
import json
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "history.db"


def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH))
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with _conn() as con:
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
        con.commit()


def save_run(result: dict, pdf_paths: list, mode: str, project: str) -> int:
    init_db()
    with _conn() as con:
        cur = con.execute("""
            INSERT INTO runs
              (run_date, export_file, pdf_count, row_count, total_inches,
               ok_rows, warn_rows, err_rows, mode, project, summary, output_path)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            result.get("output_filename", ""),
            len(pdf_paths),
            result.get("rows", 0),
            result.get("total_inches", 0),
            result.get("ok_rows", 0),
            result.get("warn_rows", 0),
            result.get("err_rows", 0),
            mode,
            project or "",
            result.get("summary", ""),
            result.get("output", ""),
        ))
        con.commit()
        return cur.lastrowid


def get_history(limit: int = 50) -> list:
    init_db()
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM runs ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def get_run(run_id: int) -> dict | None:
    init_db()
    with _conn() as con:
        row = con.execute("SELECT * FROM runs WHERE id=?", (run_id,)).fetchone()
    return dict(row) if row else None


def delete_run(run_id: int) -> bool:
    init_db()
    with _conn() as con:
        con.execute("DELETE FROM runs WHERE id=?", (run_id,))
        con.commit()
    return True
