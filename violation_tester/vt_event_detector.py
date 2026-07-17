"""
vt_event_detector.py — wires vt_planning_engine + vt_state_cache together.

Given the current set of normalized task records, runs the Planning Engine
and diffs the result against the last snapshot, returning only the tasks
whose status just got worse. On the very first run (empty cache) it seeds
the snapshot silently instead of reporting every task as "new" at once.
"""
import vt_planning_engine as engine
from vt_state_cache import VTStateCache


def detect_changes(tasks: list[dict], config: dict, cache: VTStateCache) -> list[dict]:
    """
    Returns a list of degraded task results (see VTStateCache.diff), or an
    empty list on the first-ever run for this cache (which instead seeds the
    snapshot silently so nothing floods out for pre-existing conditions).
    """
    results = engine.evaluate(tasks, config)

    if cache.is_first_run():
        cache.snapshot_silent(results)
        return []

    return cache.diff(results)
