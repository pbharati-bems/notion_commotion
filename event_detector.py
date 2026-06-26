"""
event_detector.py — Query Notion and produce typed MIS events via StateCache diff.

Three modes, each called by mis_runner.py:

    detect_changes(token, db_entries, cache)
        → queries all non-done tasks per DB, diffs against cache
        → returns {db_env_var: [TaskCreated | AssigneeChanged | StatusStarted events]}

    detect_due_soon(cache, days=2)
        → reads the local cache only (no Notion call)
        → returns rows for tasks due in {days} days not yet notified

    detect_slippage(token, db_entries, min_days=1)
        → queries all non-done tasks, filters by slippage_days >= min_days
        → returns (slipped_tasks, stalled_tasks)

    detect_overdue(token, db_entries, min_days=1)
        → queries all non-done tasks, filters by days_overdue >= min_days
        → returns (overdue_tasks, stalled_tasks)
"""
import logging

_STALLED_STATUSES = frozenset({"on hold", "blocked"})

log = logging.getLogger(__name__)


def detect_changes(token: str, db_entries: list, cache) -> dict:
    """
    For each configured DB:
      1. Fetch all non-done tasks from Notion (via get_all_tasks_for_mis).
      2. If this is the first run for that DB → snapshot silently, emit NO events.
      3. Otherwise → diff against cache and return typed events.

    Returns:
        {db_env_var: [event_dict, ...]}
    where each event_dict is {"type": str, "task": dict}
    """
    from notion_client import get_all_tasks_for_mis

    result: dict = {}

    for entry in db_entries:
        ev_var     = entry["env_var"]
        db_id      = entry["db_id"]
        fields     = entry["fields"]
        db_name    = entry["db_name"]
        in_prog    = fields.get("in_progress_value", "In progress")

        log.info("[%s] Querying Notion for all non-done tasks...", ev_var)
        try:
            tasks = get_all_tasks_for_mis(token, db_id, fields, db_name)
        except Exception as exc:
            log.error("[%s] Notion query failed: %s", ev_var, exc)
            result[ev_var] = []
            continue

        log.info("[%s] %d task(s) fetched.", ev_var, len(tasks))

        if cache.is_first_run(ev_var):
            log.info(
                "[%s] First run — silently snapshotting %d task(s). "
                "No emails sent this round.", ev_var, len(tasks),
            )
            cache.snapshot_silent(ev_var, tasks, in_progress_value=in_prog)
            result[ev_var] = []
        else:
            events = cache.diff(ev_var, tasks, in_progress_value=in_prog)
            result[ev_var] = events

    return result


def detect_due_soon(cache, days: int = 2) -> list:
    """
    Read the local cache (no Notion call) and return rows for tasks
    whose due_date == today + {days} days and that have not yet
    received a DueSoon notification for that date.

    Returns:
        [{"page_id": str, "db_env_var": str, "due_date": str, ...}, ...]
    """
    rows = cache.due_soon_tasks(days=days)
    log.info("[DueSoon] %d task(s) due in %d day(s) need notification.", len(rows), days)
    return rows


def detect_overdue(token: str, db_entries: list, min_days: int = 1) -> tuple:
    """
    Fetch all non-done tasks across DBs where overdue=true and return those
    whose days_overdue >= min_days, sorted by days_overdue descending.
    Also collects all On Hold / Blocked tasks from the same DBs.

    Returns:
        (overdue_tasks, stalled_tasks)   both sorted by days_overdue desc / name
    """
    from notion_client import get_all_tasks_for_mis

    overdue:  list = []
    stalled:  list = []
    seen_ids: set  = set()

    for entry in db_entries:
        if not entry.get("overdue", True):
            log.info("[%s] overdue=false — skipping.", entry["env_var"])
            continue
        log.info("[%s] Querying Notion for overdue tasks...", entry["env_var"])
        try:
            tasks = get_all_tasks_for_mis(
                token, entry["db_id"], entry["fields"], entry["db_name"]
            )
        except Exception as exc:
            log.error("[%s] Overdue query failed: %s", entry["env_var"], exc)
            continue

        for t in tasks:
            owners = t.get("assignees") or []
            t["owner_name"] = (
                ", ".join(o["name"] for o in owners if o.get("name")) or "—"
            )
            days = t.get("days_overdue", 0)
            if (t.get("overdue") or days > 0) and days >= min_days:
                overdue.append(t)
            if (t.get("status") or "").lower() in _STALLED_STATUSES:
                if t["id"] not in seen_ids:
                    seen_ids.add(t["id"])
                    stalled.append(t)

    overdue.sort(key=lambda t: t.get("days_overdue", 0), reverse=True)
    stalled.sort(key=lambda t: t.get("name", ""))
    log.info("[Overdue] %d overdue, %d stalled task(s) found (min_days=%d).",
             len(overdue), len(stalled), min_days)
    return overdue, stalled


def detect_slippage(token: str, db_entries: list, min_days: int = 1) -> tuple:
    """
    Fetch all non-done tasks across all DBs and return those whose
    slippage_days >= min_days, sorted by slippage_days descending.
    Also collects all On Hold / Blocked tasks from the same DBs.

    Returns:
        (slipped_tasks, stalled_tasks)
    """
    from notion_client import get_all_tasks_for_mis

    slipped:  list = []
    stalled:  list = []
    seen_ids: set  = set()

    for entry in db_entries:
        if not entry.get("slippage", True):
            log.info("[%s] slippage=false — skipping.", entry["env_var"])
            continue
        log.info("[%s] Querying Notion for slippage...", entry["env_var"])
        try:
            tasks = get_all_tasks_for_mis(
                token, entry["db_id"], entry["fields"], entry["db_name"]
            )
        except Exception as exc:
            log.error("[%s] Slippage query failed: %s", entry["env_var"], exc)
            continue

        for t in tasks:
            owners = t.get("assignees") or []
            t["owner_name"] = (
                ", ".join(o["name"] for o in owners if o.get("name")) or "—"
            )
            if (t.get("slippage_days") or 0) >= min_days:
                slipped.append(t)
            if (t.get("status") or "").lower() in _STALLED_STATUSES:
                if t["id"] not in seen_ids:
                    seen_ids.add(t["id"])
                    stalled.append(t)

    slipped.sort(key=lambda t: t.get("slippage_days", 0), reverse=True)
    stalled.sort(key=lambda t: t.get("name", ""))
    log.info("[Slippage] %d slipped, %d stalled task(s) found (min_days=%d).",
             len(slipped), len(stalled), min_days)
    return slipped, stalled
