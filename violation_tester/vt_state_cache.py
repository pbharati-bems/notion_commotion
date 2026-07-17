"""
vt_state_cache.py — SQLite-backed snapshot store for the Violation Tester.

Own database file (vt_state.db), fully separate from the main system's
mis_state.db. Stores the last-known Planning Status and Resource Conflict
flag per task so vt_event_detector.py can report only new degradations
instead of re-alerting on an unchanged violation every run.

Schema (vt_state.db):
    vt_task_state — one row per Notion page_id
        page_id            TEXT  PK
        planning_status     TEXT  — last-known status string
        resource_conflict   INT   — 1 if a conflict was flagged last run
        last_seen_at        TEXT  — UTC ISO datetime of last diff() call
"""
import logging
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

log = logging.getLogger(__name__)

# On overdue-prevention grounds, "worse" is ranked by how close a task is to
# actually missing its deadline. Data Error always alerts regardless of rank,
# since it signals a correctness problem, not a scheduling one.
_STATUS_RANK = {"On Track": 0, "Blocked": 1, "At Risk": 2, "Violated": 3}

_SCHEMA = """
CREATE TABLE IF NOT EXISTS vt_task_state (
    page_id            TEXT PRIMARY KEY,
    planning_status    TEXT NOT NULL DEFAULT '',
    resource_conflict  INTEGER NOT NULL DEFAULT 0,
    last_seen_at       TEXT NOT NULL DEFAULT ''
);
"""


def _default_state_path() -> Path:
    override = os.environ.get("VT_STATE_CACHE_DB")
    if override:
        return Path(override)
    return Path(__file__).parent / "vt_state.db"


class VTStateCache:

    def __init__(self, db_path: Path | None = None):
        self._path = db_path or _default_state_path()
        self._conn = sqlite3.connect(str(self._path))
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA journal_mode=WAL;")
        self._conn.executescript(_SCHEMA)
        self._conn.commit()
        log.debug("VTStateCache opened: %s", self._path)

    def close(self):
        self._conn.close()

    def is_first_run(self) -> bool:
        """True if this cache has never stored a snapshot yet. On first run,
        callers should use snapshot_silent() instead of diff() so an entire
        existing backlog of violations doesn't fire as "new" all at once."""
        row = self._conn.execute("SELECT 1 FROM vt_task_state LIMIT 1").fetchone()
        return row is None

    def snapshot_silent(self, results: list[dict]) -> None:
        """Store current statuses without producing any change events."""
        now = datetime.now(timezone.utc).isoformat()
        for r in results:
            self._conn.execute(
                """
                INSERT INTO vt_task_state (page_id, planning_status, resource_conflict, last_seen_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(page_id) DO UPDATE SET
                    planning_status = excluded.planning_status,
                    resource_conflict = excluded.resource_conflict,
                    last_seen_at = excluded.last_seen_at
                """,
                (r["page_id"], r["planning_status"], int(r["resource_conflict"]), now),
            )
        self._conn.commit()

    def diff(self, results: list[dict]) -> list[dict]:
        """
        Compare each result against its last stored snapshot. Returns the
        subset that represents a genuine degradation:
          - Planning Status moved to a higher-risk rank (On Track -> At Risk,
            At Risk -> Violated, etc.)
          - Planning Status newly became "Data Error" (always surfaced)
          - Resource Conflict flipped False -> True

        Always updates the stored snapshot for every task passed in,
        regardless of whether it was surfaced as a change.
        """
        now = datetime.now(timezone.utc).isoformat()
        degraded = []

        for r in results:
            row = self._conn.execute(
                "SELECT planning_status, resource_conflict FROM vt_task_state WHERE page_id = ?",
                (r["page_id"],),
            ).fetchone()

            if row is not None:
                old_status = row["planning_status"]
                old_conflict = bool(row["resource_conflict"])
                new_status = r["planning_status"]
                new_conflict = r["resource_conflict"]

                became_data_error = new_status == "Data Error" and old_status != "Data Error"
                rank_worsened = (
                    old_status in _STATUS_RANK
                    and new_status in _STATUS_RANK
                    and _STATUS_RANK[new_status] > _STATUS_RANK[old_status]
                )
                conflict_worsened = new_conflict and not old_conflict

                if became_data_error or rank_worsened or conflict_worsened:
                    degraded.append({**r, "previous_status": old_status})

            self._conn.execute(
                """
                INSERT INTO vt_task_state (page_id, planning_status, resource_conflict, last_seen_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(page_id) DO UPDATE SET
                    planning_status = excluded.planning_status,
                    resource_conflict = excluded.resource_conflict,
                    last_seen_at = excluded.last_seen_at
                """,
                (r["page_id"], r["planning_status"], int(r["resource_conflict"]), now),
            )

        self._conn.commit()
        return degraded
