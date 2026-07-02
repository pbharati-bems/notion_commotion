#!/usr/bin/env python3
"""
gen_preview_combined.py — Generate a local HTML preview of the combined
Slippage & Overdue digest without hitting Notion or sending any email.

Usage:
    python gen_preview_combined.py
    # Opens preview_combined_digest.html in the current directory.
"""
from pathlib import Path
from mailer import _build_combined_digest_html

# ── Project stubs ──────────────────────────────────────────────────────────
KUHL_KENT  = {"name": "KUHL/KENT FAN PROJECT - P4 IR",  "url": "https://notion.so/kuhl-kent"}
BEMS_TECH  = {"name": "BEMS Tech Tasks – Drives",        "url": "https://notion.so/bems-tech"}
FIELD_PROJ = {"name": "Field Projects & Commissioning",  "url": "https://notion.so/field-proj"}
NO_PROJ    = {"name": "No Project",                      "url": ""}

# ── Slipped tasks (status != On Hold → stay in main body) ─────────────────
slipped_tasks = [
    {
        "id": "t001",
        "name": "Blades & Shank",
        "url": "https://notion.so/blades-shank",
        "due_date": "2026-06-17", "original_date": "2026-05-03",
        "slippage_days": 45, "overdue": True, "days_overdue": 9,
        "days_since_update": 2,
        "all_dates": [
            "2026-05-03","2026-05-11","2026-05-12","2026-05-13",
            "2026-05-15","2026-05-27","2026-06-01","2026-06-17",
        ],
        "status": "In progress", "priority": "High", "teams": ["BEDL"],
        "owner_name": "Yatharth Dobhal", "reviewer_names": "—",
        "blocking": [],
        "project": KUHL_KENT, "db_name": "Task Tracker",
    },
    {
        "id": "t002",
        "name": "Surge test in Nashik",
        "url": "https://notion.so/surge-test-nashik",
        "due_date": "2026-06-08", "original_date": "2026-05-08",
        "slippage_days": 31, "overdue": True, "days_overdue": 18,
        "days_since_update": 63,
        "all_dates": [
            "2026-05-08","2026-05-13","2026-05-20","2026-05-25",
            "2026-05-27","2026-06-09","2026-06-02","2026-06-05","2026-06-08",
        ],
        "status": "Blocked", "priority": "High", "teams": ["BEMS", "BEDL"],
        "owner_name": "Pratyush Bharati, bakyaraj", "reviewer_names": "—",
        "blocking": [{"name": "Dispatch to CGCEL", "url": "https://notion.so/dispatch-cgcel"}],
        "project": KUHL_KENT, "db_name": "Task Tracker",
    },
    {
        "id": "t003",
        "name": "Motor Controller PCB Design – Rev C",
        "url": "https://notion.so/pcb-rev-c",
        "due_date": "2026-06-20", "original_date": "2026-05-29",
        "slippage_days": 22, "overdue": True, "days_overdue": 6,
        "days_since_update": 5,
        "all_dates": ["2026-05-29","2026-06-06","2026-06-12","2026-06-20"],
        "status": "In progress", "priority": "Medium", "teams": ["BEMS"],
        "owner_name": "Rajesh Kumar", "reviewer_names": "Ankit Shah",
        "blocking": [],
        "project": BEMS_TECH, "db_name": "Task Tracker",
    },
    {
        "id": "t004",
        "name": "Field Trial Documentation – Site A",
        "url": "https://notion.so/field-trial-doc",
        "due_date": "2026-07-03", "original_date": "2026-06-18",
        "slippage_days": 15, "overdue": False, "days_overdue": 0,
        "days_since_update": 1,
        "all_dates": ["2026-06-18","2026-06-25","2026-07-03"],
        "status": "In review", "priority": "Low", "teams": ["BEDL"],
        "owner_name": "Suresh Nair", "reviewer_names": "Yatharth Dobhal",
        "blocking": [],
        "project": FIELD_PROJ, "db_name": "Task Tracker",
    },
    {
        "id": "t005",
        "name": "Production Test Fixture v2",
        "url": "https://notion.so/test-fixture-v2",
        "due_date": "2026-06-23", "original_date": "2026-06-15",
        "slippage_days": 8, "overdue": True, "days_overdue": 3,
        "days_since_update": 9,
        "all_dates": ["2026-06-15","2026-06-20","2026-06-23"],
        "status": "In progress", "priority": "High", "teams": ["BEMS"],
        "owner_name": "Meera Pillai", "reviewer_names": "Rajesh Kumar",
        "blocking": [
            {"name": "Factory Acceptance Test", "url": "https://notion.so/fat"},
        ],
        "project": BEMS_TECH, "db_name": "Task Tracker",
    },
]

# ── Overdue-only tasks (slippage_days == 0 — never rescheduled, just late) ─
overdue_tasks = [
    {
        "id": "t006",
        "name": "Weekly Project Status Report – Week 25",
        "url": "https://notion.so/weekly-report-w25",
        "due_date": "2026-06-14", "original_date": "2026-06-14",
        "slippage_days": 0, "overdue": True, "days_overdue": 12,
        "days_since_update": 20,
        "all_dates": [],
        "status": "In progress", "priority": "Low", "teams": [],
        "owner_name": "bakyaraj", "reviewer_names": "—",
        "blocking": [],
        "project": NO_PROJ, "db_name": "Task Tracker",
    },
    {
        "id": "t007",
        "name": "Safety Audit Checklist – Panel Room Update",
        "url": "https://notion.so/safety-audit",
        "due_date": "2026-06-19", "original_date": "2026-06-19",
        "slippage_days": 0, "overdue": True, "days_overdue": 7,
        "days_since_update": 4,
        "all_dates": [],
        "status": "In progress", "priority": "Medium", "teams": ["BEDL"],
        "owner_name": "Arun Krishnamurthy", "reviewer_names": "—",
        "blocking": [],
        "project": FIELD_PROJ, "db_name": "Task Tracker",
    },
]

# ── Stalled tasks (On Hold + Blocked — collected by detector, shown at bottom) ─
stalled_tasks = [
    {
        "id": "t008",
        "name": "Crompton Visit with (V2.1)",
        "url": "https://notion.so/crompton-visit",
        "due_date": "2026-05-18", "original_date": "2026-05-15",
        "slippage_days": 3, "overdue": True, "days_overdue": 39,
        "days_since_update": 45,
        "all_dates": ["2026-05-15","2026-05-18"],
        "status": "On hold", "priority": "", "teams": [],
        "owner_name": "—", "reviewer_names": "—",
        "blocking": [],
        "project": FIELD_PROJ, "db_name": "Task Tracker",
    },
    {
        "id": "t009",
        "name": "Vendor Evaluation – Phase 2 (IGBT Sourcing)",
        "url": "https://notion.so/vendor-eval-p2",
        "due_date": "2026-07-10", "original_date": "2026-07-10",
        "slippage_days": 0, "overdue": False, "days_overdue": 0,
        "days_since_update": 60,
        "all_dates": [],
        "status": "On hold", "priority": "Medium", "teams": ["BEMS"],
        "owner_name": "Suresh Nair", "reviewer_names": "—",
        "blocking": [],
        "project": BEMS_TECH, "db_name": "Task Tracker",
    },
    # Blocked task also in stalled list (also appears in main body via slipped_tasks t002)
    {
        "id": "t002",
        "name": "Surge test in Nashik",
        "url": "https://notion.so/surge-test-nashik",
        "due_date": "2026-06-08", "original_date": "2026-05-08",
        "slippage_days": 31, "overdue": True, "days_overdue": 18,
        "days_since_update": 63,
        "all_dates": [
            "2026-05-08","2026-05-13","2026-05-20","2026-05-25",
            "2026-05-27","2026-06-09","2026-06-02","2026-06-05","2026-06-08",
        ],
        "status": "Blocked", "priority": "High", "teams": ["BEMS", "BEDL"],
        "owner_name": "Pratyush Bharati, bakyaraj", "reviewer_names": "—",
        "blocking": [{"name": "Dispatch to CGCEL", "url": "https://notion.so/dispatch-cgcel"}],
        "project": KUHL_KENT, "db_name": "Task Tracker",
    },
    {
        "id": "t010",
        "name": "Legacy HMI Panel Replacement",
        "url": "https://notion.so/legacy-hmi",
        "due_date": "2026-06-16", "original_date": "2026-06-16",
        "slippage_days": 0, "overdue": True, "days_overdue": 10,
        "days_since_update": 12,
        "all_dates": [],
        "status": "On hold", "priority": "Low", "teams": [],
        "owner_name": "Meera Pillai", "reviewer_names": "—",
        "blocking": [],
        "project": BEMS_TECH, "db_name": "Task Tracker",
    },
]

html = _build_combined_digest_html(slipped_tasks, overdue_tasks, stalled_tasks)

out = Path("test_emails") / "preview_combined_digest.html"
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(html, encoding="utf-8")
print(f"Preview written → {out.resolve()}")
