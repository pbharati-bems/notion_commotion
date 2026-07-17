"""
Offline unit tests for vt_state_cache — uses a temp SQLite file, no Notion
API calls. Run directly: python -m violation_tester.tests.test_state_cache
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from vt_state_cache import VTStateCache  # noqa: E402


def _make_result(page_id, status, resource_conflict=False):
    return {"page_id": page_id, "planning_status": status, "resource_conflict": resource_conflict}


def test_first_run_is_silent():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        assert cache.is_first_run()
        cache.snapshot_silent([_make_result("t1", "On Track")])
        assert not cache.is_first_run()
        cache.close()


def test_diff_reports_worsening_status():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        cache.snapshot_silent([_make_result("t1", "On Track")])
        degraded = cache.diff([_make_result("t1", "At Risk")])
        assert len(degraded) == 1
        assert degraded[0]["previous_status"] == "On Track"
        cache.close()


def test_diff_ignores_unchanged_status():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        cache.snapshot_silent([_make_result("t1", "At Risk")])
        degraded = cache.diff([_make_result("t1", "At Risk")])
        assert degraded == []
        cache.close()


def test_diff_ignores_improving_status():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        cache.snapshot_silent([_make_result("t1", "Violated")])
        degraded = cache.diff([_make_result("t1", "On Track")])
        assert degraded == []
        cache.close()


def test_diff_reports_new_resource_conflict():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        cache.snapshot_silent([_make_result("t1", "On Track", resource_conflict=False)])
        degraded = cache.diff([_make_result("t1", "On Track", resource_conflict=True)])
        assert len(degraded) == 1
        cache.close()


def test_diff_always_reports_new_data_error():
    with tempfile.TemporaryDirectory() as d:
        cache = VTStateCache(Path(d) / "vt_state.db")
        cache.snapshot_silent([_make_result("t1", "On Track")])
        degraded = cache.diff([_make_result("t1", "Data Error")])
        assert len(degraded) == 1
        cache.close()


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
