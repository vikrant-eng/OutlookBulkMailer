"""
Database module for Outlook Bulk Mailer.
Manages SQLite storage for campaigns and email history.
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import Optional


DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "mailer.db")


def get_connection() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initialize database tables."""
    conn = get_connection()
    cur = conn.cursor()

    cur.executescript("""
        CREATE TABLE IF NOT EXISTS campaigns (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL,
            subject     TEXT NOT NULL,
            body        TEXT NOT NULL,
            created_at  TEXT NOT NULL,
            sent_at     TEXT,
            total       INTEGER DEFAULT 0,
            sent        INTEGER DEFAULT 0,
            failed      INTEGER DEFAULT 0,
            status      TEXT DEFAULT 'draft'
        );

        CREATE TABLE IF NOT EXISTS recipients (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            campaign_id  INTEGER NOT NULL REFERENCES campaigns(id) ON DELETE CASCADE,
            email        TEXT NOT NULL,
            name         TEXT,
            status       TEXT DEFAULT 'pending',
            error_msg    TEXT,
            sent_at      TEXT
        );

        CREATE TABLE IF NOT EXISTS attachments (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            campaign_id  INTEGER NOT NULL REFERENCES campaigns(id) ON DELETE CASCADE,
            filename     TEXT NOT NULL,
            filepath     TEXT NOT NULL
        );
    """)
    conn.commit()
    conn.close()


def create_campaign(name: str, subject: str, body: str, recipients: list[dict],
                    attachments: list[dict] | None = None) -> int:
    conn = get_connection()
    cur = conn.cursor()
    now = datetime.now().isoformat()

    cur.execute(
        "INSERT INTO campaigns (name, subject, body, created_at, total, status) VALUES (?,?,?,?,?,?)",
        (name, subject, body, now, len(recipients), "pending")
    )
    campaign_id = cur.lastrowid

    cur.executemany(
        "INSERT INTO recipients (campaign_id, email, name, status) VALUES (?,?,?,?)",
        [(campaign_id, r["email"], r.get("name", ""), "pending") for r in recipients]
    )

    if attachments:
        cur.executemany(
            "INSERT INTO attachments (campaign_id, filename, filepath) VALUES (?,?,?)",
            [(campaign_id, a["filename"], a["filepath"]) for a in attachments]
        )

    conn.commit()
    conn.close()
    return campaign_id


def get_campaigns() -> list[dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM campaigns ORDER BY created_at DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_campaign(campaign_id: int) -> Optional[dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM campaigns WHERE id=?", (campaign_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def get_recipients(campaign_id: int, status: Optional[str] = None) -> list[dict]:
    conn = get_connection()
    cur = conn.cursor()
    if status:
        cur.execute("SELECT * FROM recipients WHERE campaign_id=? AND status=?", (campaign_id, status))
    else:
        cur.execute("SELECT * FROM recipients WHERE campaign_id=?", (campaign_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_attachments(campaign_id: int) -> list[dict]:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM attachments WHERE campaign_id=?", (campaign_id,))
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def update_recipient_status(recipient_id: int, status: str, error_msg: str = ""):
    conn = get_connection()
    cur = conn.cursor()
    now = datetime.now().isoformat() if status == "sent" else None
    cur.execute(
        "UPDATE recipients SET status=?, error_msg=?, sent_at=? WHERE id=?",
        (status, error_msg, now, recipient_id)
    )
    # Update campaign counters
    cur.execute("SELECT campaign_id FROM recipients WHERE id=?", (recipient_id,))
    row = cur.fetchone()
    if row:
        cid = row["campaign_id"]
        cur.execute(
            """UPDATE campaigns SET
                sent  = (SELECT COUNT(*) FROM recipients WHERE campaign_id=? AND status='sent'),
                failed = (SELECT COUNT(*) FROM recipients WHERE campaign_id=? AND status='failed')
               WHERE id=?""",
            (cid, cid, cid)
        )
    conn.commit()
    conn.close()


def finish_campaign(campaign_id: int, status: str = "completed"):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "UPDATE campaigns SET status=?, sent_at=? WHERE id=?",
        (status, datetime.now().isoformat(), campaign_id)
    )
    conn.commit()
    conn.close()


def delete_campaign(campaign_id: int):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM campaigns WHERE id=?", (campaign_id,))
    conn.commit()
    conn.close()
