"""MySQL database module for cold email tracking."""

from __future__ import annotations

import os
from pathlib import Path

import mysql.connector
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

_db_available = False
_db_error: str | None = None


def is_available() -> bool:
    return _db_available


def last_error() -> str | None:
    return _db_error


def _connect():
    return mysql.connector.connect(
        host=os.getenv("DB_HOST", "127.0.0.1"),
        port=int(os.getenv("DB_PORT", 3306)),
        user=os.getenv("DB_USER", "root"),
        password=os.getenv("DB_PASSWORD", "root"),
        database=os.getenv("DB_NAME", "cold_email_tracker"),
    )


def init_db() -> bool:
    """Create the contacts table if it doesn't exist. Returns True if connected."""
    global _db_available, _db_error
    try:
        conn = _connect()
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                id           INT AUTO_INCREMENT PRIMARY KEY,
                company_name VARCHAR(255),
                person_name  VARCHAR(255),
                email        VARCHAR(255),
                type         VARCHAR(50),
                status       VARCHAR(100) DEFAULT 'Delivered',
                notes        TEXT,
                sent_at      DATETIME     DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        cur.close()
        conn.close()
        _db_available = True
        _db_error = None
        return True
    except Exception as exc:
        _db_available = False
        _db_error = str(exc)
        return False


def add_contact(
    company_name: str,
    person_name: str,
    email: str,
    type_: str,
    status: str = "Delivered",
    notes: str = "",
) -> int | None:
    """Insert a new contact row. Returns the new row id, or None if DB unavailable."""
    if not _db_available:
        return None
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO contacts (company_name, person_name, email, type, status, notes)
               VALUES (%s, %s, %s, %s, %s, %s)""",
            (company_name, person_name, email, type_, status, notes),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        cur.close()
        conn.close()


def get_all_contacts() -> list[dict]:
    """Return all contacts ordered by most recent first."""
    if not _db_available:
        raise RuntimeError(_db_error or "Tracking database is not connected.")
    conn = _connect()
    cur = conn.cursor(dictionary=True)
    try:
        cur.execute("SELECT * FROM contacts ORDER BY sent_at DESC")
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()
    for row in rows:
        if row.get("sent_at"):
            row["sent_at"] = row["sent_at"].strftime("%Y-%m-%d %H:%M")
    return rows


def update_contact(id_: int, **fields) -> None:
    """Update any fields for a contact. Pass only the fields you want to change."""
    if not _db_available:
        raise RuntimeError(_db_error or "Tracking database is not connected.")
    allowed = {"company_name", "person_name", "email", "type", "status", "notes"}
    updates = {k: v for k, v in fields.items() if k in allowed}
    if not updates:
        return
    set_clause = ", ".join(f"{k}=%s" for k in updates)
    values = list(updates.values()) + [id_]
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute(f"UPDATE contacts SET {set_clause} WHERE id=%s", values)
        conn.commit()
    finally:
        cur.close()
        conn.close()


def delete_contact(id_: int) -> None:
    """Delete a contact by id."""
    if not _db_available:
        raise RuntimeError(_db_error or "Tracking database is not connected.")
    conn = _connect()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM contacts WHERE id=%s", (id_,))
        conn.commit()
    finally:
        cur.close()
        conn.close()
