#!/usr/bin/env python3
"""
mis_runner.py — MIS email automation entry point.

Usage:
  python mis_runner.py --events       Detect task changes  (Req 1 + Req 4)
  python mis_runner.py --daily        Due-in-2-days reminders  (Req 5)
  python mis_runner.py --slippage     Slippage digest  (Req 3)
  python mis_runner.py --overdue      Overdue digest
  python mis_runner.py --combined     Combined Slippage & Overdue digest (replaces --slippage + --overdue)
  python mis_runner.py --all          Run all modes in sequence
  python mis_runner.py --dry-run      Add to any flag: log what WOULD be sent, no emails

Examples:
  python mis_runner.py --events --dry-run        # preview change detection
  python mis_runner.py --all                     # full MIS run
  python mis_runner.py --combined --dry-run      # preview combined digest
  python mis_runner.py --slippage --min-slip 3   # only tasks slipped >= 3 days

Cron schedule (IST = UTC+5:30):
  # Every 15 min on workdays 9 AM – 7 PM — detect task changes
  */15 3-13 * * 1-6  /path/.venv/bin/python /path/mis_runner.py --events

  # Daily 8:00 AM IST — due-in-2-days reminders
  30 2 * * 1-6  /path/.venv/bin/python /path/mis_runner.py --daily

  # Daily 9:00 AM IST — combined slippage+overdue digest
  30 3 * * 1-6  /path/.venv/bin/python /path/mis_runner.py --combined
"""
import argparse
import json
import logging
import os
import sys
from pathlib import Path

import requests as _req
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

_REQUIRED_ENV = [
    "NOTION_TOKEN",
    "AZURE_TENANT_ID",
    "AZURE_CLIENT_ID",
    "AZURE_CLIENT_SECRET",
]


# ---------------------------------------------------------------------------
# Config loaders
# ---------------------------------------------------------------------------

def _load_app_config(path: Path) -> dict:
    try:
        cfg = json.loads(path.read_text())
    except Exception as exc:
        log.error("Cannot read %s: %s", path.name, exc)
        sys.exit(1)
    if not cfg.get("sender_email"):
        log.error("%s missing 'sender_email'", path.name)
        sys.exit(1)
    return cfg


def _resolve_db_entries(notion_token: str, db_config_path: Path) -> list:
    """
    Load databases config, resolve each env_var to a DB ID, and fetch
    the human-readable DB name from Notion.
    Returns list of dicts: {env_var, db_id, db_name, fields}
    """
    try:
        raw_dbs = json.loads(db_config_path.read_text())
    except Exception as exc:
        log.error("Cannot read %s: %s", db_config_path.name, exc)
        sys.exit(1)

    headers = {
        "Authorization":  f"Bearer {notion_token}",
        "Notion-Version": "2022-06-28",
    }
    entries = []
    for entry in raw_dbs:
        env_var = entry.get("env_var", "")
        db_id   = os.getenv(env_var)
        if not db_id:
            log.warning("Skipping %r — not set in .env", env_var)
            continue

        # Resolve friendly name from Notion
        db_name = env_var
        try:
            r = _req.get(
                f"https://api.notion.com/v1/databases/{db_id}",
                headers=headers, timeout=10,
            )
            if r.ok:
                db_name = (
                    "".join(t.get("plain_text", "") for t in r.json().get("title", []))
                    or env_var
                )
        except Exception:
            pass

        entries.append({
            "env_var":  env_var,
            "db_id":    db_id,
            "db_name":  db_name,
            "fields":   entry.get("fields", {}),
            "digest":   entry.get("digest",   True),
            "slippage": entry.get("slippage", True),
            "overdue":  entry.get("overdue",  True),
        })

    if not entries:
        log.error("No databases resolved. Check databases.json and .env.")
        sys.exit(1)

    return entries


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="MIS email automation — task tracking, reminders, slippage.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--events",   action="store_true",
                        help="Detect task changes: TaskCreated, AssigneeChanged, StatusStarted")
    parser.add_argument("--daily",    action="store_true",
                        help="Send due-in-2-days reminders to task owners/PMs")
    parser.add_argument("--slippage", action="store_true",
                        help="Send slippage digest to executives, leaders, PMs")
    parser.add_argument("--overdue",   action="store_true",
                        help="Send overdue digest to executives, leaders, PMs")
    parser.add_argument("--combined", action="store_true",
                        help="Send combined Slippage & Overdue digest (replaces --slippage + --overdue)")
    parser.add_argument("--all",      action="store_true",
                        help="Run --events + --daily + --combined in sequence")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Log what would be sent — no emails actually sent")
    parser.add_argument("--min-slip", type=int, default=None,
                        help="Min slippage days for the digest (overrides config.json)")
    parser.add_argument("--test",     action="store_true",
                        help="Use test config (config.test.json, databases.test.json, mis_state.test.db)")
    args = parser.parse_args()

    if not any([args.events, args.daily, args.slippage, args.overdue, args.combined, args.all]):
        parser.print_help()
        return 1

    if args.all:
        args.events = args.daily = args.combined = True

    # ── Test-mode config selection ────────────────────────────────────────
    _base = Path(__file__).parent
    if args.test:
        app_config_path = _base / "config.test.json"
        db_config_path  = _base / "databases.test.json"
        os.environ["STATE_CACHE_DB"] = str(_base / "mis_state.test.db")
        log.info("TEST MODE — config.test.json | databases.test.json | mis_state.test.db")
    else:
        app_config_path = _base / "config.json"
        db_config_path  = _base / "databases.json"

    # ── Environment checks ────────────────────────────────────────────────
    missing = [k for k in _REQUIRED_ENV if not os.getenv(k)]
    if missing:
        log.error("Missing environment variables: %s", ", ".join(missing))
        return 1

    notion_token = os.environ["NOTION_TOKEN"]
    cfg          = _load_app_config(app_config_path)
    db_entries   = _resolve_db_entries(notion_token, db_config_path)

    min_slip = (
        args.min_slip
        if args.min_slip is not None
        else cfg.get("mis", {}).get("slippage_min_days", 1)
    )
    due_days = cfg.get("mis", {}).get("due_soon_days", 2)

    log.info(
        "MIS run started — modes: %s | dry_run=%s | %d DB(s)",
        "+".join(m for m in ("events","daily","slippage","overdue","combined") if getattr(args, m, False)),
        args.dry_run,
        len(db_entries),
    )

    # ── Acquire Graph API token once ──────────────────────────────────────
    graph_token = None
    if not args.dry_run:
        from mailer import _acquire_token
        try:
            graph_token = _acquire_token(
                os.environ["AZURE_TENANT_ID"],
                os.environ["AZURE_CLIENT_ID"],
                os.environ["AZURE_CLIENT_SECRET"],
            )
            log.info("Microsoft Graph token acquired.")
        except Exception as exc:
            log.error("Failed to acquire Graph token: %s", exc)
            return 1

    from state_cache    import StateCache
    from event_detector import detect_changes, detect_due_soon, detect_slippage, detect_overdue
    from event_dispatcher import (
        dispatch_events, dispatch_due_soon,
        dispatch_slippage, dispatch_overdue, dispatch_combined,
    )

    cache      = StateCache()
    all_errors: list = []

    # ── --events  (Req 1 + Req 4) ─────────────────────────────────────────
    if args.events:
        log.info("─── MODE: --events ───────────────────────────────────────")
        by_db = detect_changes(notion_token, db_entries, cache)

        for entry in db_entries:
            ev_var = entry["env_var"]
            events = by_db.get(ev_var, [])

            if not events:
                log.info("[%s] No changes detected.", ev_var)
                continue

            log.info("[%s] %d event(s) to dispatch.", ev_var, len(events))

            # Log event summary always (useful even without dry-run)
            for ev in events:
                log.info(
                    "  %-20s  %s", ev["type"], ev["task"].get("name", "?")
                )

            if args.dry_run:
                log.info("[%s] DRY-RUN — no emails sent.", ev_var)
                continue

            errs = dispatch_events(
                events       = events,
                cfg          = cfg,
                token        = graph_token,
                notion_token = notion_token,
                db_id        = entry["db_id"],
                db_env_var   = ev_var,
                fields       = entry["fields"],
                db_name      = entry["db_name"],
            )
            all_errors.extend(errs)

    # ── --daily  (Req 5) ──────────────────────────────────────────────────
    if args.daily:
        log.info("─── MODE: --daily ────────────────────────────────────────")
        due_rows = detect_due_soon(cache, days=due_days)

        if not due_rows:
            log.info("No tasks due in %d day(s). Nothing to send.", due_days)
        else:
            # Build page_id → task lookup from a fresh Notion fetch
            from notion_client import get_all_tasks_for_mis
            full_tasks: dict = {}
            for entry in db_entries:
                try:
                    tasks = get_all_tasks_for_mis(
                        notion_token, entry["db_id"], entry["fields"], entry["db_name"]
                    )
                    full_tasks.update({t["id"]: t for t in tasks})
                except Exception as exc:
                    log.error("[%s] Fetch failed for daily mode: %s", entry["env_var"], exc)

            for row in due_rows:
                task = full_tasks.get(row["page_id"])
                name = task["name"] if task else row["page_id"]
                log.info("  DueSoon  %s  (due %s)", name, row["due_date"])

            if args.dry_run:
                log.info("DRY-RUN — no emails sent.")
            else:
                errs = dispatch_due_soon(due_rows, full_tasks, cfg, graph_token, cache,
                                     dry_run=args.dry_run)
                all_errors.extend(errs)

    # ── --slippage  (Req 3) ───────────────────────────────────────────────
    if args.slippage:
        log.info("─── MODE: --slippage ─────────────────────────────────────")
        slipped, stalled_slip = detect_slippage(notion_token, db_entries, min_days=min_slip)

        if not slipped and not stalled_slip:
            log.info("No slippage or stalled tasks found (min_days=%d). Nothing to send.", min_slip)
        else:
            for t in slipped:
                log.info("  +%3dd  %s", t.get("slippage_days", 0), t["name"])
            for t in stalled_slip:
                log.info("  [%-8s]  %s", t.get("status", "?"), t["name"])

            if args.dry_run:
                log.info("DRY-RUN — no emails sent.")
            else:
                errs = dispatch_slippage(slipped, stalled_slip, cfg, graph_token)
                all_errors.extend(errs)

    # ── --overdue ─────────────────────────────────────────────────────────
    if args.overdue:
        log.info("─── MODE: --overdue ──────────────────────────────────────")
        min_overdue = cfg.get("mis", {}).get("overdue_min_days", 1)
        overdue_tasks, stalled_ov = detect_overdue(notion_token, db_entries, min_days=min_overdue)

        if not overdue_tasks and not stalled_ov:
            log.info("No overdue or stalled tasks found (min_days=%d). Nothing to send.", min_overdue)
        else:
            for t in overdue_tasks:
                log.info("  +%3dd overdue  %s", t.get("days_overdue", 0), t["name"])
            for t in stalled_ov:
                log.info("  [%-8s]  %s", t.get("status", "?"), t["name"])

            if args.dry_run:
                log.info("DRY-RUN — no emails sent.")
            else:
                errs = dispatch_overdue(overdue_tasks, stalled_ov, cfg, graph_token)
                all_errors.extend(errs)

    # ── --combined  (Slippage + Overdue in one email) ────────────────────
    if args.combined:
        log.info("─── MODE: --combined ─────────────────────────────────────")
        min_overdue = cfg.get("mis", {}).get("overdue_min_days", 1)

        slipped, stalled_slip = detect_slippage(notion_token, db_entries, min_days=min_slip)
        overdue, stalled_ov   = detect_overdue(notion_token, db_entries, min_days=min_overdue)

        # Merge stalled lists (deduplicated)
        seen_stalled: set  = set()
        all_stalled:  list = []
        for t in stalled_slip + stalled_ov:
            if t["id"] not in seen_stalled:
                seen_stalled.add(t["id"])
                all_stalled.append(t)

        if not slipped and not overdue and not all_stalled:
            log.info("No schedule exceptions found. Nothing to send.")
        else:
            for t in slipped:
                log.info("  slip +%3dd  %s", t.get("slippage_days", 0), t["name"])
            for t in overdue:
                log.info("  ovd  +%3dd  %s", t.get("days_overdue", 0), t["name"])
            for t in all_stalled:
                log.info("  [%-8s]  %s", t.get("status", "?"), t["name"])

            if args.dry_run:
                log.info("DRY-RUN — no emails sent.")
            else:
                errs = dispatch_combined(slipped, overdue, all_stalled, cfg, graph_token)
                all_errors.extend(errs)

    # ── Wrap-up ───────────────────────────────────────────────────────────
    cache.close()

    stats = ""
    try:
        from state_cache import StateCache as _SC
        _c = _SC()
        s  = _c.stats()
        _c.close()
        stats = (
            f" | cache: {s.get('total',0)} task(s) across "
            f"{s.get('dbs',0)} DB(s)"
        )
    except Exception:
        pass

    if all_errors:
        log.error("MIS run finished with %d error(s)%s:", len(all_errors), stats)
        for e in all_errors:
            log.error("  %s", e)
        return 1

    log.info("MIS run complete — no errors%s.", stats)
    return 0


if __name__ == "__main__":
    sys.exit(main())
