"""
Offline unit tests for vt_planning_engine — synthetic fixtures only, no Notion
API calls. Mirrors the real DB's known tricky cases: a mutual-dependency
cycle (like "Loop Test Part A/B"), a missing-date record, and a resource
overlap. Run directly: python -m violation_tester.tests.test_planning_engine
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import vt_planning_engine as engine  # noqa: E402


def make_task(page_id, title, status="Not Started", priority="Medium",
              start_date=None, due_date=None, duration_days=5, duration_tag="Effort",
              assignees=None, blocked_by=None, dependencies_met=None, allowed_slippage=2,
              task_type="Development"):
    return {
        "page_id": page_id,
        "url": f"https://notion.so/{page_id}",
        "title": title,
        "status": status,
        "priority": priority,
        "task_type": task_type,
        "start_date": start_date,
        "due_date": due_date,
        "due_date_history": [],
        "duration_days": duration_days,
        "duration_tag": duration_tag,
        "assignees": assignees or [],
        "blocked_by": blocked_by or [],
        "blocks": [],
        "dependencies_met": dependencies_met,
        "allowed_slippage": allowed_slippage,
        "created_time": "2026-01-01T00:00:00.000Z",
    }


def _status_of(results, page_id):
    return next(r for r in results if r["page_id"] == page_id)


def test_on_track():
    tasks = [make_task("t1", "On Track Task", start_date="2026-07-01", due_date="2026-07-20",
                        duration_days=5, allowed_slippage=2)]
    results = engine.evaluate(tasks, today=date(2026, 7, 15))
    result = _status_of(results, "t1")
    assert result["planning_status"] == "On Track"
    assert result["task_type"] == "Development"
    assert result["owner"] == "Unassigned"


def test_violated_when_past_due():
    tasks = [make_task("t1", "Late Task", start_date="2026-06-20", due_date="2026-07-01",
                        duration_days=5, allowed_slippage=2)]
    results = engine.evaluate(tasks, today=date(2026, 7, 15))
    assert _status_of(results, "t1")["planning_status"] == "Violated"


def test_blocked_by_incomplete_predecessor():
    pred = make_task("p1", "Predecessor", status="In Progress",
                      start_date="2026-07-10", due_date="2026-07-25", duration_days=10)
    succ = make_task("s1", "Successor", start_date="2026-07-01", due_date="2026-07-25",
                      duration_days=3, allowed_slippage=5, blocked_by=["p1"])
    results = engine.evaluate([pred, succ], today=date(2026, 7, 15))
    succ_result = _status_of(results, "s1")
    assert succ_result["planning_status"] == "Blocked"
    assert succ_result["ees"] == date(2026, 7, 20)  # pushed to predecessor's projected finish
    assert succ_result["waiting_on"] == ["Predecessor"]


def test_bottleneck_names_predecessor_and_owner():
    pred = make_task("p1", "Predecessor", status="In Progress", assignees=["Asha Rao"],
                      start_date="2026-07-10", due_date="2026-07-25", duration_days=10)
    succ = make_task("s1", "Successor", start_date="2026-07-01", due_date="2026-07-25",
                      duration_days=3, allowed_slippage=5, blocked_by=["p1"])
    results = engine.evaluate([pred, succ], today=date(2026, 7, 15))
    bottleneck = _status_of(results, "s1")["bottleneck"]
    assert bottleneck["predecessor"] == "Predecessor"
    assert bottleneck["owner"] == "Asha Rao"
    assert bottleneck["delay_days"] == 19  # EES (Jul 20) - own planned start (Jul 1)


def test_bottleneck_owner_defaults_to_unassigned():
    pred = make_task("p1", "Predecessor", status="In Progress",
                      start_date="2026-07-10", due_date="2026-07-25", duration_days=10)
    succ = make_task("s1", "Successor", start_date="2026-07-01", due_date="2026-07-25",
                      duration_days=3, allowed_slippage=5, blocked_by=["p1"])
    results = engine.evaluate([pred, succ], today=date(2026, 7, 15))
    assert _status_of(results, "s1")["bottleneck"]["owner"] == "Unassigned"


def test_bottleneck_absent_when_violated_without_predecessor():
    tasks = [make_task("t1", "Late Task", start_date="2026-06-20", due_date="2026-07-01",
                        duration_days=5, allowed_slippage=2)]
    results = engine.evaluate(tasks, today=date(2026, 7, 15))
    assert _status_of(results, "t1")["bottleneck"] is None


def test_at_risk_within_slippage_buffer():
    tasks = [make_task("t1", "Slightly Slipping", start_date="2026-07-01", due_date="2026-07-10",
                        duration_days=12, allowed_slippage=5)]
    results = engine.evaluate(tasks, today=date(2026, 7, 5))
    result = _status_of(results, "t1")
    assert result["planning_status"] == "At Risk"
    assert result["projected_slip_days"] == 3


def test_reassignment_prefers_shared_owner_over_same_type():
    blocked = make_task("blk", "Blocked Task", assignees=["Priya"], task_type="Development",
                         start_date="2026-07-01", due_date="2026-07-05", blocked_by=["p"])
    pred = make_task("p", "Predecessor", status="In Progress",
                      start_date="2026-07-01", due_date="2026-07-20", duration_days=15)
    same_owner_slack = make_task("so", "Priya's Other Task", assignees=["Priya"], task_type="Procurement",
                                  start_date="2026-07-01", due_date="2026-07-20", duration_days=5)
    same_type_slack = make_task("st", "Same Type Task", task_type="Development",
                                 start_date="2026-07-01", due_date="2026-07-25", duration_days=5)
    results = engine.evaluate([blocked, pred, same_owner_slack, same_type_slack], today=date(2026, 7, 3))
    suggestion = _status_of(results, "blk")["reassignment_suggestion"]
    assert suggestion["basis"] == "owner"
    assert suggestion["candidate_task"] == "Priya's Other Task"


def test_reassignment_falls_back_to_same_type():
    blocked = make_task("blk", "Blocked Task", task_type="Development",
                         start_date="2026-07-01", due_date="2026-07-05", blocked_by=["p"])
    pred = make_task("p", "Predecessor", status="In Progress",
                      start_date="2026-07-01", due_date="2026-07-20", duration_days=15)
    same_type_slack = make_task("st", "Same Type Task", task_type="Development",
                                 start_date="2026-07-01", due_date="2026-07-25", duration_days=5)
    other_type_slack = make_task("ot", "Other Type Task", task_type="Procurement",
                                  start_date="2026-07-01", due_date="2026-07-30", duration_days=5)
    results = engine.evaluate([blocked, pred, same_type_slack, other_type_slack], today=date(2026, 7, 3))
    suggestion = _status_of(results, "blk")["reassignment_suggestion"]
    assert suggestion["basis"] == "type"
    assert suggestion["candidate_task"] == "Same Type Task"


def test_reassignment_none_when_no_slack_anywhere():
    blocked = make_task("blk", "Blocked Task", start_date="2026-07-01", due_date="2026-07-05", blocked_by=["p"])
    pred = make_task("p", "Predecessor", status="In Progress",
                      start_date="2026-07-01", due_date="2026-07-20", duration_days=15)
    on_schedule = make_task("os", "On Schedule Task", start_date="2026-07-01", due_date="2026-07-06", duration_days=5)
    results = engine.evaluate([blocked, pred, on_schedule], today=date(2026, 7, 3))
    assert _status_of(results, "blk")["reassignment_suggestion"] is None


def test_reassignment_absent_for_on_track_tasks():
    tasks = [make_task("t1", "Fine Task", start_date="2026-07-01", due_date="2026-07-20", duration_days=5)]
    results = engine.evaluate(tasks, today=date(2026, 7, 5))
    assert _status_of(results, "t1")["reassignment_suggestion"] is None


def test_dependency_cycle_is_data_error():
    a = make_task("a", "Loop Test Part A", start_date="2026-07-01", due_date="2026-07-10", blocked_by=["b"])
    b = make_task("b", "Loop Test Part B", start_date="2026-07-01", due_date="2026-07-10", blocked_by=["a"])
    results = engine.evaluate([a, b], today=date(2026, 7, 5))
    for tid in ("a", "b"):
        r = _status_of(results, tid)
        assert r["planning_status"] == "Data Error"
        assert "cycle" in r["data_error_reason"]


def test_missing_due_date_is_data_error():
    tasks = [make_task("t1", "No Due Date", start_date="2026-07-01", due_date=None)]
    results = engine.evaluate(tasks, today=date(2026, 7, 15))
    result = _status_of(results, "t1")
    assert result["planning_status"] == "Data Error"
    assert "missing" in result["data_error_reason"]


def test_resource_conflict_flagged_and_forces_at_risk():
    x = make_task("x", "Task X", start_date="2026-07-01", due_date="2026-07-30",
                   duration_days=5, assignees=["Alice"])
    y = make_task("y", "Task Y", start_date="2026-07-03", due_date="2026-07-30",
                   duration_days=5, assignees=["Alice"])
    results = engine.evaluate([x, y], today=date(2026, 7, 2))
    rx, ry = _status_of(results, "x"), _status_of(results, "y")
    assert rx["resource_conflict"] and ry["resource_conflict"]
    assert rx["planning_status"] == "At Risk"
    assert ry["planning_status"] == "At Risk"


def test_dependencies_met_mismatch_is_data_error():
    tasks = [make_task("t1", "Mismatched", start_date="2026-07-01", due_date="2026-07-20",
                        blocked_by=[], dependencies_met=False)]
    results = engine.evaluate(tasks, today=date(2026, 7, 5))
    result = _status_of(results, "t1")
    assert result["planning_status"] == "Data Error"
    assert "disagrees" in result["data_error_reason"]


def run_all():
    tests = [v for k, v in globals().items() if k.startswith("test_") and callable(v)]
    passed, failed = 0, []
    for t in tests:
        try:
            t()
            passed += 1
        except AssertionError as e:
            failed.append((t.__name__, str(e)))
    print(f"{passed}/{len(tests)} passed")
    for name, err in failed:
        print(f"FAIL {name}: {err}")
    return not failed


if __name__ == "__main__":
    ok = run_all()
    sys.exit(0 if ok else 1)
