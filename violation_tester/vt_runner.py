#!/usr/bin/env python3
"""
vt_runner.py — Violation Tester entry point.

Usage:
  python vt_runner.py --calculate              Run the Planning Engine, print a summary. No email.
  python vt_runner.py --calculate --verbose    Also list every task's computed status.
  python vt_runner.py --digest                 Render the Supervisory Digest and save it locally (dry run — default, safe).
  python vt_runner.py --digest --send          Actually send the digest via Microsoft Graph. Real email, real recipients.

--calculate only ever reads Notion and writes to this folder's own
vt_state.db — it never sends mail. --digest without --send never sends mail
either; it only writes an HTML file under violation_tester/test_emails/.
--send is the only path in this entire codebase that can put a message in
someone's inbox, and it's opt-in every time you run it.
"""
import argparse
import json
import logging
import os
import sys
from collections import Counter
from pathlib import Path

from dotenv import dotenv_values, load_dotenv

import vt_mailer as mailer
import vt_notion_client as client
import vt_planning_engine as engine
from vt_state_cache import VTStateCache

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "vt_logs"

log = logging.getLogger("vt_runner")


def _load_config() -> tuple[dict, dict, dict]:
    env = dotenv_values(BASE_DIR / ".env.vt")
    dbcfg = json.loads((BASE_DIR / "vt_databases.json").read_text())
    cfg = json.loads((BASE_DIR / "vt_config.json").read_text())
    return env, dbcfg, cfg


def _setup_logging() -> None:
    LOG_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(LOG_DIR / "vt_runner.log"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def run_calculate(verbose: bool) -> list[dict]:
    env, dbcfg, cfg = _load_config()

    tasks = client.fetch_all_tasks(
        env["VT_NOTION_TOKEN"], env["VT_NOTION_DATABASE_ID"], dbcfg["fields"],
        project_filter=dbcfg.get("project_filter"),
    )
    log.info("Fetched %d valid task rows", len(tasks))

    results = engine.evaluate(tasks, cfg)
    counts = Counter(r["planning_status"] for r in results)
    log.info("Status counts: %s", dict(counts))

    if verbose:
        for r in sorted(results, key=lambda r: r["title"]):
            log.info(
                "  %-32s %-11s due=%s ees=%s slip=%s",
                r["title"], r["planning_status"], r["due_date"], r["ees"], r["projected_slip_days"],
            )

    cache = VTStateCache()
    try:
        if cache.is_first_run():
            cache.snapshot_silent(results)
            log.info("First run for this cache — snapshot seeded silently, no degradations reported.")
        else:
            degraded = cache.diff(results)
            if degraded:
                log.info("%d task(s) degraded since the last run:", len(degraded))
                for r in degraded:
                    log.info("  %s: %s -> %s", r["title"], r["previous_status"], r["planning_status"])
            else:
                log.info("No new degradations since the last run.")
    finally:
        cache.close()

    return results


def run_digest(send: bool) -> None:
    env, dbcfg, cfg = _load_config()
    load_dotenv(BASE_DIR / ".env.vt")  # populate os.environ for vt_mailer (Azure creds)
    os.environ["VT_MAILER_DRY_RUN"] = "0" if send else "1"

    tasks = client.fetch_all_tasks(
        env["VT_NOTION_TOKEN"], env["VT_NOTION_DATABASE_ID"], dbcfg["fields"],
        project_filter=dbcfg.get("project_filter"),
    )
    log.info("Fetched %d valid task rows", len(tasks))

    results = engine.evaluate(tasks, cfg)
    log.info("Status counts: %s", dict(Counter(r["planning_status"] for r in results)))

    sender = cfg["sender_email"]
    recipients = cfg["recipients"]
    log.info("Building Supervisory Digest (%s) — sender=%s recipients=%s",
              "SEND" if send else "DRY RUN", sender, recipients)

    out_path = mailer.build_and_send_digest(results, sender, recipients)
    if out_path:
        log.info("Dry run complete — wrote %s (open it in a browser to preview)", out_path)
    else:
        log.info("Digest sent via Microsoft Graph to %s", ", ".join(recipients))


def main() -> None:
    parser = argparse.ArgumentParser(description="Violation Tester runner")
    parser.add_argument("--calculate", action="store_true", help="Run the Planning Engine and print a summary. No email.")
    parser.add_argument("--verbose", action="store_true", help="List every task's computed status.")
    parser.add_argument("--digest", action="store_true", help="Render the Supervisory Digest. Dry run (saved locally) unless --send is also given.")
    parser.add_argument("--send", action="store_true", help="With --digest: actually send via Microsoft Graph instead of a dry run.")
    args = parser.parse_args()

    _setup_logging()

    if args.calculate:
        run_calculate(args.verbose)
    elif args.digest:
        run_digest(args.send)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
