"""
SQLite persistence layer for SLA dashboard.

Two features:
  1. Snapshots — every time tickets are fetched from SharePoint, the full
     dataset is saved as a numbered snapshot so users can browse history.
  2. Overrides — a per-ticket state override table that drives the
     "Resolve Ticket" button in the dashboard.

Uses SQLite (built into Python) — no server or admin rights required.
Database file: sla_dashboard.db in the same folder as this module.
"""

import json
import logging
from typing import Optional
import os
import sqlite3
from datetime import datetime

logger = logging.getLogger(__name__)

_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sla_dashboard.db")


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_conn() -> sqlite3.Connection:
    """Open and return a SQLite connection with row_factory set."""
    conn = sqlite3.connect(_DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")   # better concurrent read perf
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def init_db() -> bool:
    """
    Create all tables / indexes if they do not already exist.
    Safe to call on every startup.
    Returns True on success, False on failure.
    """
    try:
        conn = _get_conn()
        with conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS snapshots (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    fetched_at   TEXT    NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now')),
                    ticket_count INTEGER NOT NULL DEFAULT 0,
                    file_name    TEXT    NOT NULL DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS ticket_records (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    snapshot_id   INTEGER NOT NULL
                                  REFERENCES snapshots(id) ON DELETE CASCADE,
                    ticket_number TEXT    NOT NULL DEFAULT '',
                    data          TEXT    NOT NULL DEFAULT '{}'
                );

                CREATE INDEX IF NOT EXISTS idx_tr_snapshot
                    ON ticket_records(snapshot_id);

                CREATE INDEX IF NOT EXISTS idx_tr_number
                    ON ticket_records(ticket_number);

                CREATE TABLE IF NOT EXISTS ticket_overrides (
                    ticket_number TEXT PRIMARY KEY,
                    state         TEXT NOT NULL DEFAULT 'Resolved',
                    resolved_at   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now')),
                    comment       TEXT NOT NULL DEFAULT ''
                );
            """)
        conn.close()
        logger.info("Database schema ready (%s)", _DB_PATH)
        return True
    except Exception as exc:
        logger.error("init_db failed: %s", exc)
        return False


def save_snapshot(records: list, file_name: str = "") -> Optional[int]:
    """
    Persist a list of sanitised ticket dicts as a new snapshot.
    Automatically prunes snapshots beyond the last 500.
    Returns the new snapshot_id, or None on failure.
    """
    try:
        conn = _get_conn()
        with conn:
            cur = conn.execute(
                "INSERT INTO snapshots (ticket_count, file_name) VALUES (?, ?)",
                (len(records), file_name or ""),
            )
            snapshot_id = cur.lastrowid

            if records:
                conn.executemany(
                    "INSERT INTO ticket_records (snapshot_id, ticket_number, data) VALUES (?, ?, ?)",
                    [
                        (snapshot_id, str(r.get("Number") or ""), json.dumps(r, default=str))
                        for r in records
                    ],
                )

            # Keep only the 500 most-recent snapshots
            conn.execute("""
                DELETE FROM snapshots
                WHERE id NOT IN (
                    SELECT id FROM snapshots
                    ORDER BY fetched_at DESC
                    LIMIT 500
                )
            """)

        conn.close()
        logger.info("Snapshot %s saved (%s tickets)", snapshot_id, len(records))
        return snapshot_id
    except Exception as exc:
        logger.error("save_snapshot failed: %s", exc)
        return None


def list_snapshots() -> list:
    """
    Return a list of snapshot metadata dicts ordered most-recent first.
    Deduplicated by file_name (latest snapshot per file).
    Each dict: {id, fetched_at, ticket_count, file_name}
    """
    try:
        conn = _get_conn()
        rows = conn.execute("""
            SELECT id, fetched_at, ticket_count, file_name
            FROM (
                SELECT id, fetched_at, ticket_count, COALESCE(file_name, '') AS file_name,
                       ROW_NUMBER() OVER (
                           PARTITION BY COALESCE(file_name, '')
                           ORDER BY fetched_at DESC
                       ) AS rn
                FROM snapshots
            )
            WHERE rn = 1
            ORDER BY fetched_at DESC
            LIMIT 200
        """).fetchall()
        conn.close()
        return [dict(r) for r in rows]
    except Exception as exc:
        logger.error("list_snapshots failed: %s", exc)
        return []


def get_snapshot_tickets(snapshot_id: int) -> Optional[dict]:
    """
    Return dict with snapshot metadata + ticket data list.
    Shape: {id, fetched_at, file_name, ticket_count, tickets: [...]}
    Returns None if snapshot not found.
    """
    try:
        conn = _get_conn()
        row = conn.execute(
            "SELECT id, fetched_at, ticket_count, COALESCE(file_name, '') AS file_name "
            "FROM snapshots WHERE id = ?",
            (snapshot_id,),
        ).fetchone()
        if not row:
            conn.close()
            return None
        meta = dict(row)
        records = conn.execute(
            "SELECT data FROM ticket_records WHERE snapshot_id = ? ORDER BY id",
            (snapshot_id,),
        ).fetchall()
        conn.close()
        meta["tickets"] = [json.loads(r["data"]) for r in records]
        return meta
    except Exception as exc:
        logger.error("get_snapshot_tickets(%s) failed: %s", snapshot_id, exc)
        return None


def get_overrides() -> dict:
    """
    Return {ticket_number: {state, resolved_at, comment}} for every active override.
    """
    try:
        conn = _get_conn()
        rows = conn.execute(
            "SELECT ticket_number, state, resolved_at, COALESCE(comment, '') AS comment "
            "FROM ticket_overrides"
        ).fetchall()
        conn.close()
        return {r["ticket_number"]: dict(r) for r in rows}
    except Exception as exc:
        logger.error("get_overrides failed: %s", exc)
        return {}


def set_override(ticket_number: str, state: str = "Resolved", comment: str = "") -> bool:
    """Insert or update a ticket state override."""
    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = _get_conn()
        with conn:
            conn.execute(
                """
                INSERT INTO ticket_overrides (ticket_number, state, resolved_at, comment)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(ticket_number) DO UPDATE
                    SET state = excluded.state,
                        resolved_at = excluded.resolved_at,
                        comment = excluded.comment
                """,
                (ticket_number, state, now, comment or ""),
            )
        conn.close()
        return True
    except Exception as exc:
        logger.error("set_override(%s) failed: %s", ticket_number, exc)
        return False


def remove_override(ticket_number: str) -> bool:
    """Delete a ticket state override (undo manual resolve)."""
    try:
        conn = _get_conn()
        with conn:
            conn.execute(
                "DELETE FROM ticket_overrides WHERE ticket_number = ?",
                (ticket_number,),
            )
        conn.close()
        return True
    except Exception as exc:
        logger.error("remove_override(%s) failed: %s", ticket_number, exc)
        return False
