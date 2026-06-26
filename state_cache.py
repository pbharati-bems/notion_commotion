"""
state_cache.py — SQLite-backed task state snapshot and event diff engine.

The cache stores the last-known state of every Notion task page so that
mis_runner.py can diff the latest data against it and emit typed change events.

Schema (mis_state.db):
    task_state  — one row per Notion page_id
        page_id              TEXT  PK
        db_env_var           TEXT  — which database this task belongs to
        status               TEXT  — last-known status string
        assignee_emails      TEXT  — JSON list of emails (sorted, for stable comparison)
        due_date             TEXT  — YYYY-MM-DD or empty
        last_seen_at         TEXT  — UTC ISO datetime of last diff() call
        notified_started     INT   — 1 once a StatusStarted email has been sent
        notified_due_soon    TEXT  — due_date string for which DueSoon was sent

Events emitted by diff():
    TaskCreated      — page_id was not in the cache (brand-new task)
    AssigneeChanged  — the set of assignee emails changed
    StatusStarted    — status changed to the DB's in_progress_value (fired once per task)
"""
import json
import logging
import os
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

log = logging.getLogger(__name__)


def _default_state_path() -> Path:
    override = os.environ.get("STATE_CACHE_DB")
    if override:
        return Path(override)
    return Path(__file__).parent / "mis_state.db"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS task_state (
    page_id           TEXT PRIMARY KEY,
    db_env_var        TEXT NOT NULL,
    status            TEXT NOT NULL DEFAULT '',
    assignee_emails   TEXT NOT NULL DEFAULT '[]',
    due_date          TEXT NOT NULL DEFAULT '',
    last_seen_at      TEXT NOT NULL DEFAULT '',
    notified_started  INTEGER NOT NULL DEFAULT 0,
    notified_due_soon TEXT NOT NULL DEFAULT ''
);
"""


class StateCache:

    def __init__(self, db_path: Path | None = None):
        if db_path is None:
            db_path = _default_state_path()
        self._path = db_path
        self._conn = sqlite3.connect(str(db_path))
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA journal_mode=WAL;")
        self._conn.executescript(_SCHEMA)
        self._conn.commit()
        log.debug("StateCache opened: %s", db_path)

    def close(self):
        self._conn.close()

    # ------------------------------------------------------------------
    # is_first_run — True when no rows exist yet for this database
    # ------------------------------------------------------------------

    def is_first_run(self, db_env_var: str) -> bool:
        """
        Returns True if this database has never been snapshotted.
        On first run we call snapshot_silent() instead of diff() so we
        don't flood everyone with "TaskCreated" for every existing task.
        """
        row = self._conn.execute(
            "SELECT 1 FROM task_state WHERE db_env_var = ? LIMIT 1",
            (db_env_var,),
        ).fetchone()
        return row is None

    # ------------------------------------------------------------------
    # snapshot_silent — populate cache on first run, no events emitted
    # ------------------------------------------------------------------

    def snapshot_silent(self, db_env_var: str, tasks: list,
                        in_progress_value: str = "In progress"):
        """
        Bulk-insert tasks without emitting any events.
        Tasks already in "in progress" are marked notified_started=1
        so a StatusStarted email is not sent retroactively on the next run.
        """
        now = datetime.utcnow().isoformat()
        cur = self._conn.cursor()
        for task in tasks:
            pid    = task["id"]
            emails = json.dumps(sorted(
                a.get("email", "") for a in task.get("assignees", []) if a.get("email")
            ))
            already_started = 1 if task.get("status") == in_progress_value else 0
            cur.execute(
                """INSERT OR IGNORE INTO task_state
                   (page_id, db_env_var, status, assignee_emails,
                    due_date, last_seen_at, notified_started, notified_due_soon)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (pid, db_env_var, task.get("status", ""),
                 emails, task.get("due_date") or "",
                 now, already_started, ""),
            )
        self._conn.commit()
        log.info("[%s] snapshot_silent: %d task(s) stored (no emails sent)", db_env_var, len(tasks))

    # ------------------------------------------------------------------
    # diff — compare fresh Notion snapshot against stored state
    # ------------------------------------------------------------------

    def diff(self, db_env_var: str, fresh_tasks: list,
             in_progress_value: str = "In progress") -> list:
        """
        Compare fresh_tasks (from Notion) against the stored snapshot.

        Returns a list of event dicts:
            {"type": "TaskCreated",     "task": task_dict}
            {"type": "AssigneeChanged", "task": task_dict}
            {"type": "StatusStarted",   "task": task_dict}
        """
        now    = datetime.utcnow().isoformat()
        events: list = []
        cur    = self._conn.cursor()

        for task in fresh_tasks:
            pid        = task["id"]
            new_emails = json.dumps(sorted(
                a.get("email", "") for a in task.get("assignees", []) if a.get("email")
            ))
            new_status = task.get("status") or ""
            new_due    = task.get("due_date") or ""

            row = cur.execute(
                "SELECT * FROM task_state WHERE page_id = ?", (pid,)
            ).fetchone()

            if row is None:
                # Brand-new task — never seen before
                events.append({"type": "TaskCreated", "task": task})
                cur.execute(
                    """INSERT INTO task_state
                       (page_id, db_env_var, status, assignee_emails,
                        due_date, last_seen_at, notified_started, notified_due_soon)
                       VALUES (?,?,?,?,?,?,0,'')""",
                    (pid, db_env_var, new_status, new_emails, new_due, now),
                )
            else:
                # Existing task — check for relevant field changes
                if row["assignee_emails"] != new_emails:
                    events.append({"type": "AssigneeChanged", "task": task})

                if (new_status == in_progress_value
                        and row["status"] != in_progress_value
                        and not row["notified_started"]):
                    events.append({"type": "StatusStarted", "task": task})
                    cur.execute(
                        "UPDATE task_state SET notified_started = 1 WHERE page_id = ?", (pid,)
                    )

                # Always refresh the stored state
                cur.execute(
                    """UPDATE task_state
                       SET status = ?, assignee_emails = ?, due_date = ?, last_seen_at = ?
                       WHERE page_id = ?""",
                    (new_status, new_emails, new_due, now, pid),
                )

        self._conn.commit()
        log.info("[%s] diff → %d event(s) from %d task(s)",
                 db_env_var, len(events), len(fresh_tasks))
        return events

    # ------------------------------------------------------------------
    # due_soon_tasks — tasks due in N days not yet notified
    # ------------------------------------------------------------------

    def due_soon_tasks(self, days: int = 2) -> list:
        """
        Return rows whose due_date == today + {days} days and for which
        a DueSoon notification has not been sent for that date yet.

        Caller must call mark_due_soon_notified() after a successful send.
        """
        target = (date.today() + timedelta(days=days)).isoformat()
        rows   = self._conn.execute(
            """SELECT page_id, db_env_var, status, assignee_emails, due_date
               FROM task_state
               WHERE due_date = ? AND notified_due_soon != ?""",
            (target, target),
        ).fetchall()
        return [dict(r) for r in rows]

    def mark_due_soon_notified(self, page_id: str, due_date: str):
        """Record that the DueSoon email has been sent for this task's current due date."""
        self._conn.execute(
            "UPDATE task_state SET notified_due_soon = ? WHERE page_id = ?",
            (due_date, page_id),
        )
        self._conn.commit()

    # ------------------------------------------------------------------
    # stats — quick summary for logging
    # ------------------------------------------------------------------

    def stats(self) -> dict:
        row = self._conn.execute(
            "SELECT COUNT(*) as total, "
            "SUM(notified_started) as started, "
            "COUNT(DISTINCT db_env_var) as dbs "
            "FROM task_state"
        ).fetchone()
        return dict(row) if row else {}
