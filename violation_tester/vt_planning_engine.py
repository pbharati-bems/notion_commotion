"""
vt_planning_engine.py — pure calculation engine for the Violation Tester.

No I/O of any kind. Takes the normalized task records produced by
vt_notion_client.fetch_all_tasks() and returns each task's Effective Earliest
Start (EES), Resource Conflict flag, and Planning Status — nothing here ever
reads or writes Notion, a file, or a database.

For a Blocked/Violated task whose EES was pushed later by a predecessor,
each result's "bottleneck" field names that predecessor, its owner, and how
many days it pushed this task's start back — the root-cause behind the
status, not just the status itself.

Every Blocked/Violated result also carries a "reassignment_suggestion": the
On Track task elsewhere in the project with the most spare schedule margin,
preferring one that shares an assignee, then one of the same Task Type, then
any task project-wide — or None if no slack exists anywhere to draw on. This
is computed fresh from the current dataset each run, never hardcoded.

Planning Status decision matrix, evaluated top to bottom (first match wins):
    1. Data Error   — missing/inverted dates, dependency cycle, or unresolved
                       predecessor, or a mismatch against Notion's own
                       "Dependencies Met" field
    2. Violated     — projected finish slips past Due Date by more than the
                       Allowed Slippage, or the due date has already passed
    3. Blocked      — a predecessor isn't Complete yet and EES is still ahead
    4. At Risk      — slip is within the Allowed Slippage buffer, or a
                       Resource Conflict is present
    5. On Track     — none of the above
"""
from collections import defaultdict, deque
from datetime import date, datetime, timedelta

DONE_STATUS = "Complete"
_STATUS_ORDER = ["On Track", "Blocked", "At Risk", "Violated"]


def _parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        return None


def _resolve_start_date(task: dict) -> date | None:
    """Start Date (with fallback): Planned Start, else Created Time."""
    start = _parse_date(task.get("start_date"))
    if start is not None:
        return start
    return _parse_date(task.get("created_time"))


def _resolve_duration(task: dict, start: date | None, due: date | None) -> int | None:
    """Duration in days: parsed Efforts/Lead Time field, else Due − Start."""
    if task.get("duration_days") is not None:
        return task["duration_days"]
    if start is not None and due is not None:
        return (due - start).days
    return None


def _build_graph(tasks_by_id: dict) -> tuple[dict, dict]:
    """Returns (in_degree, successors) built from each task's blocked_by list,
    counting only predecessor ids that are actually present in this dataset."""
    in_degree = {tid: 0 for tid in tasks_by_id}
    successors = {tid: [] for tid in tasks_by_id}
    for tid, task in tasks_by_id.items():
        valid_preds = [p for p in task["blocked_by"] if p in tasks_by_id]
        in_degree[tid] = len(valid_preds)
        for p in valid_preds:
            successors[p].append(tid)
    return in_degree, successors


def _topological_order(tasks_by_id: dict) -> tuple[list, list]:
    """Kahn's algorithm. Returns (resolved_order, stuck_ids) — stuck_ids are
    tasks that never reach in-degree zero, i.e. they sit in or depend on a
    dependency cycle. Terminates in all cases; never loops forever."""
    in_degree, successors = _build_graph(tasks_by_id)
    queue = deque(tid for tid, deg in in_degree.items() if deg == 0)
    order = []
    remaining = dict(in_degree)
    while queue:
        tid = queue.popleft()
        order.append(tid)
        for succ in successors[tid]:
            remaining[succ] -= 1
            if remaining[succ] == 0:
                queue.append(succ)
    stuck_ids = [tid for tid in tasks_by_id if tid not in order]
    return order, stuck_ids


def _predecessors_complete(task: dict, tasks_by_id: dict) -> bool:
    valid_preds = [p for p in task["blocked_by"] if p in tasks_by_id]
    if not valid_preds:
        return True
    return all(tasks_by_id[p]["status"] == DONE_STATUS for p in valid_preds)


def _compute_ees_and_finish(order: list, tasks_by_id: dict, resolved: dict) -> None:
    """Forward pass (CPM-style). Mutates `resolved` in place, filling in
    ees/finish for every task in topological order."""
    for tid in order:
        task = tasks_by_id[tid]
        own_start = resolved[tid]["start_date"]
        valid_preds = [p for p in task["blocked_by"] if p in tasks_by_id]

        candidate = own_start
        bottleneck_pred = None
        for p in valid_preds:
            p_task = tasks_by_id[p]
            if p_task["status"] == DONE_STATUS:
                # No separate "actual finish" field exists in this DB yet —
                # Due Date is used as the completed-on-schedule stand-in.
                p_finish = resolved[p]["due_date"] or resolved[p]["finish"]
            else:
                p_finish = resolved[p]["finish"]
            if p_finish is not None and (candidate is None or p_finish > candidate):
                candidate = p_finish
                bottleneck_pred = p

        resolved[tid]["ees"] = candidate
        resolved[tid]["bottleneck_pred_id"] = bottleneck_pred
        if candidate is not None and resolved[tid]["duration"] is not None:
            resolved[tid]["finish"] = candidate + timedelta(days=resolved[tid]["duration"])
        else:
            resolved[tid]["finish"] = None


def _compute_resource_conflicts(order: list, tasks_by_id: dict, resolved: dict, max_concurrent: int) -> dict:
    """Interval-overlap check per assignee, counting only duration_tag=='Effort'
    tasks (Lead-time/procurement waits don't occupy a person's bandwidth)."""
    by_assignee = defaultdict(list)
    for tid in order:
        task = tasks_by_id[tid]
        if task.get("duration_tag") != "Effort":
            continue
        if resolved[tid]["ees"] is None or resolved[tid]["finish"] is None:
            continue
        for name in task["assignees"]:
            by_assignee[name].append(tid)

    conflicts = defaultdict(set)
    for name, tids in by_assignee.items():
        for i in range(len(tids)):
            for j in range(i + 1, len(tids)):
                a, b = tids[i], tids[j]
                a_s, a_e = resolved[a]["ees"], resolved[a]["finish"]
                b_s, b_e = resolved[b]["ees"], resolved[b]["finish"]
                if a_s < b_e and b_s < a_e:
                    conflicts[a].add(b)
                    conflicts[b].add(a)

    flags = {}
    for tid in order:
        overlapping = conflicts.get(tid, set())
        concurrent_load = 1 + len(overlapping)
        flags[tid] = {
            "resource_conflict": concurrent_load > max_concurrent,
            "conflicts_with": sorted(overlapping),
        }
    return flags


def _planning_status(
    task: dict,
    resolved_entry: dict,
    today: date,
    has_conflict: bool,
    is_stuck: bool,
    deps_mismatch: bool,
) -> tuple[str, str | None]:
    due = resolved_entry["due_date"]
    start = resolved_entry["start_date"]
    ees = resolved_entry["ees"]
    finish = resolved_entry["finish"]

    if is_stuck:
        return "Data Error", "part of or depends on a dependency cycle"
    if due is None or start is None:
        return "Data Error", "missing Start Date or Due Date"
    if due < start:
        return "Data Error", "Due Date precedes Start Date"
    if deps_mismatch:
        return "Data Error", "engine's dependency check disagrees with Notion's Dependencies Met field"
    if finish is None:
        return "Data Error", "duration could not be determined"

    allowed_slippage = task.get("allowed_slippage") or 0
    slip = (finish - due).days

    if slip > allowed_slippage or today > due:
        return "Violated", None
    if not _predecessors_complete(task, resolved_entry["_tasks_by_id"]) and ees is not None and ees > today:
        return "Blocked", None
    if (0 < slip <= allowed_slippage) or has_conflict:
        return "At Risk", None
    return "On Track", None


def _slack_days(result: dict) -> int:
    """Spare schedule margin: how many days early a task is projected to
    finish before its due date. Zero if it has none (or isn't computable)."""
    if result["projected_slip_days"] is None:
        return 0
    return max(0, -result["projected_slip_days"])


def _suggest_reassignment(task: dict, tasks_by_id: dict, on_track_pool: list[dict]) -> dict | None:
    """Scans On Track tasks for spare bandwidth to recommend pulling from,
    preferring a shared assignee, then the same Task Type, then anything
    project-wide with slack. Excludes this task's own predecessors — pulling
    bandwidth from the very thing blocking it isn't a fix. Returns None if no
    slack exists anywhere."""
    own_preds = set(task["blocked_by"])

    def has_slack(r):
        return _slack_days(r) > 0 and r["page_id"] not in own_preds

    same_owner = []
    if task["assignees"]:
        same_owner = [
            r for r in on_track_pool
            if has_slack(r) and set(tasks_by_id[r["page_id"]]["assignees"]) & set(task["assignees"])
        ]

    same_type = [
        r for r in on_track_pool
        if has_slack(r) and tasks_by_id[r["page_id"]]["task_type"] == task["task_type"]
    ]

    any_slack = [r for r in on_track_pool if has_slack(r)]

    if same_owner:
        pool, basis = same_owner, "owner"
    elif same_type:
        pool, basis = same_type, "type"
    elif any_slack:
        pool, basis = any_slack, "project"
    else:
        return None

    best = max(pool, key=_slack_days)
    return {
        "candidate_task": best["title"],
        "candidate_type": tasks_by_id[best["page_id"]]["task_type"],
        "candidate_owner": best["owner"],
        "slack_days": _slack_days(best),
        "basis": basis,
    }


def evaluate(tasks: list[dict], config: dict | None = None, today: date | None = None) -> list[dict]:
    """
    Run the full Planning Engine over a list of normalized task records
    (as returned by vt_notion_client.fetch_all_tasks). Returns one result
    dict per input task — nothing is silently dropped; data problems are
    reported as Planning Status "Data Error" with a reason.
    """
    config = config or {}
    max_concurrent = (config.get("resource_capacity") or {}).get("max_concurrent_effort_tasks", 1)
    today = today or date.today()

    tasks_by_id = {t["page_id"]: t for t in tasks}
    order, stuck_ids = _topological_order(tasks_by_id)
    stuck_set = set(stuck_ids)

    resolved = {}
    for tid, task in tasks_by_id.items():
        start = _resolve_start_date(task)
        due = _parse_date(task.get("due_date"))
        resolved[tid] = {
            "start_date": start,
            "due_date": due,
            "duration": _resolve_duration(task, start, due),
            "ees": None,
            "finish": None,
            "bottleneck_pred_id": None,
            "_tasks_by_id": tasks_by_id,
        }

    _compute_ees_and_finish(order, tasks_by_id, resolved)
    conflict_flags = _compute_resource_conflicts(order, tasks_by_id, resolved, max_concurrent)

    results = []
    for tid, task in tasks_by_id.items():
        entry = resolved[tid]
        is_stuck = tid in stuck_set
        conflict = conflict_flags.get(tid, {"resource_conflict": False, "conflicts_with": []})

        deps_mismatch = False
        if not is_stuck and task.get("dependencies_met") is not None:
            engine_deps_met = _predecessors_complete(task, tasks_by_id)
            deps_mismatch = engine_deps_met != task["dependencies_met"]

        status, reason = _planning_status(
            task, entry, today,
            has_conflict=conflict["resource_conflict"],
            is_stuck=is_stuck,
            deps_mismatch=deps_mismatch,
        )

        slip = None
        if entry["finish"] is not None and entry["due_date"] is not None:
            slip = (entry["finish"] - entry["due_date"]).days

        bottleneck = None
        pred_id = entry.get("bottleneck_pred_id")
        if pred_id and status in ("Blocked", "Violated") and entry["ees"] is not None and entry["start_date"] is not None:
            pred_task = tasks_by_id[pred_id]
            bottleneck = {
                "predecessor": pred_task["title"],
                "owner": ", ".join(pred_task["assignees"]) or "Unassigned",
                "delay_days": (entry["ees"] - entry["start_date"]).days,
            }

        results.append({
            "page_id": tid,
            "title": task["title"],
            "url": task["url"],
            "status": task["status"],
            "priority": task["priority"],
            "task_type": task["task_type"],
            "owner": ", ".join(task["assignees"]) or "Unassigned",
            "start_date": entry["start_date"],
            "due_date": entry["due_date"],
            "duration_days": entry["duration"],
            "ees": entry["ees"],
            "projected_finish": entry["finish"],
            "allowed_slippage": task.get("allowed_slippage"),
            "projected_slip_days": slip,
            "resource_conflict": conflict["resource_conflict"],
            "conflicts_with": [tasks_by_id[c]["title"] for c in conflict["conflicts_with"] if c in tasks_by_id],
            "planning_status": status,
            "data_error_reason": reason,
            "bottleneck": bottleneck,
            "waiting_on": [tasks_by_id[p]["title"] for p in task["blocked_by"] if p in tasks_by_id],
        })

    on_track_pool = [r for r in results if r["planning_status"] == "On Track"]
    for r in results:
        r["reassignment_suggestion"] = (
            _suggest_reassignment(tasks_by_id[r["page_id"]], tasks_by_id, on_track_pool)
            if r["planning_status"] in ("Blocked", "Violated")
            else None
        )

    return results
