"""
event_dispatcher.py — Route MIS events to the correct recipients and email templates.

Called by mis_runner.py after event_detector.py produces events.

Recipient resolution rules (from config.json "roles"):
    TaskCreated     →  assignees  + roles.admin  + roles.pm[db_env_var]
    AssigneeChanged →  assignees  + roles.admin  + roles.pm[db_env_var]
    StatusStarted   →  assignees  + roles.admin  + roles.pm[db_env_var]
    DueSoon         →  assignees  + roles.admin  + roles.pm[db_env_var]
    Slippage digest →  roles.executive + roles.leader + all roles.pm values
    Overdue digest  →  roles.executive + roles.leader + all roles.pm values

Subject lines:
    [New Task]        {task_name} — due {date}
    [Reassigned]      {task_name} — due {date}
    [Project Started] {task_name} — {today}
    [Reminder]        {task_name} due {date}
    [Slippage Report] N task(s) with shifted due dates — {today}
    [Overdue Report]  N task(s) past due date — {today}
"""
import logging
from datetime import date

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Recipient helpers
# ---------------------------------------------------------------------------

def _admin_emails(cfg: dict) -> list:
    return cfg.get("roles", {}).get("admin", [])


def _pm_emails(cfg: dict, db_env_var: str) -> list:
    return cfg.get("roles", {}).get("pm", {}).get(db_env_var, [])


def _executive_emails(cfg: dict) -> list:
    return cfg.get("roles", {}).get("executive", [])


def _leader_emails(cfg: dict) -> list:
    return cfg.get("roles", {}).get("leader", [])


def _all_pm_emails(cfg: dict) -> list:
    pm_map = cfg.get("roles", {}).get("pm", {})
    return [email for lst in pm_map.values() for email in lst]


def _merge(*lists) -> list:
    """Merge email lists, deduplicate, strip blanks, preserve order."""
    seen:   set  = set()
    result: list = []
    for lst in lists:
        for email in (lst or []):
            e = (email or "").strip()
            if e and e not in seen:
                seen.add(e)
                result.append(e)
    return result


def _task_recipients(task: dict, cfg: dict, db_env_var: str) -> list:
    """Standard recipient list for task-level notifications."""
    assignee_emails = [a["email"] for a in task.get("assignees", []) if a.get("email")]
    return _merge(assignee_emails, _admin_emails(cfg), _pm_emails(cfg, db_env_var))


# ---------------------------------------------------------------------------
# dispatch_events — handle TaskCreated / AssigneeChanged / StatusStarted
# ---------------------------------------------------------------------------

def dispatch_events(
    events: list,
    cfg: dict,
    token: str,
    notion_token: str,
    db_id: str,
    db_env_var: str,
    fields: dict,
    db_name: str,
) -> list:
    """
    Send emails for a list of change events.
    Returns list of error strings (empty = all sent OK).

    Parameters:
        token         — pre-acquired Microsoft Graph API token
        notion_token  — Notion API token (for fetching sub-tasks on StatusStarted)
        db_id         — Notion database ID
        db_env_var    — env var key for this DB (used to look up PM list)
        fields        — database field-name config from databases.json
        db_name       — human-readable DB name
    """
    from mailer import (
        _build_task_assigned_html,
        _build_project_started_html,
        send_mis_email,
    )
    from notion_client import get_subtasks

    sender    = cfg["sender_email"]
    today_str = date.today().strftime("%d %b %Y")
    errors:  list = []

    for event in events:
        etype = event["type"]
        task  = event["task"]
        name  = task.get("name", "?")

        try:
            # ── TaskCreated / AssigneeChanged ────────────────────────────
            if etype in ("TaskCreated", "AssigneeChanged"):
                reason     = "created" if etype == "TaskCreated" else "reassigned"
                recipients = _task_recipients(task, cfg, db_env_var)

                if not recipients:
                    log.warning("[%s] %s — no recipients configured, skipping.", etype, name)
                    continue

                due_str = task.get("due_date") or "no due date"
                subject = (
                    f"[New Task] {name} — due {due_str}"
                    if reason == "created" else
                    f"[Reassigned] {name} — due {due_str}"
                )
                html = _build_task_assigned_html(task, reason)
                send_mis_email(token, sender, recipients, subject, html)
                log.info("[%s] \"%s\" → %d recipient(s)", etype, name, len(recipients))

            # ── StatusStarted ────────────────────────────────────────────
            elif etype == "StatusStarted":
                recipients = _task_recipients(task, cfg, db_env_var)

                if not recipients:
                    log.warning("[StatusStarted] %s — no recipients configured, skipping.", name)
                    continue

                # Fetch sub-tasks from the same DB (returns [] if none / field absent)
                sub_tasks = get_subtasks(notion_token, db_id, fields, task["id"], db_name)

                subject = f"[Project Started] {name} — {today_str}"
                html    = _build_project_started_html(task, sub_tasks)
                send_mis_email(token, sender, recipients, subject, html)
                log.info(
                    "[StatusStarted] \"%s\" → %d recipient(s), %d sub-task(s)",
                    name, len(recipients), len(sub_tasks),
                )

            else:
                log.warning("Unknown event type: %s — skipped.", etype)

        except Exception as exc:
            msg = f"{etype} | {name}: {exc}"
            errors.append(msg)
            log.error("[dispatch_events] %s", msg)

    return errors


# ---------------------------------------------------------------------------
# dispatch_due_soon — DueSoon reminders (Req 5)
# ---------------------------------------------------------------------------

def dispatch_due_soon(
    due_rows: list,
    full_tasks: dict,
    cfg: dict,
    token: str,
    cache,
    dry_run: bool = False,
) -> list:
    """
    Send a due-soon reminder for each cache row returned by detect_due_soon().

    Parameters:
        due_rows   — from cache.due_soon_tasks()
        full_tasks — {page_id: task_dict} built from get_all_tasks_for_mis()
        cache      — StateCache instance (to call mark_due_soon_notified)
        dry_run    — if True, log what would be sent but do NOT mark as notified

    Returns list of error strings.
    """
    from mailer import _build_due_soon_html, send_mis_email

    sender = cfg["sender_email"]
    errors: list = []

    for row in due_rows:
        pid  = row["page_id"]
        task = full_tasks.get(pid)

        if not task:
            # Task may have been closed/deleted — silence the reminder
            if not dry_run:
                cache.mark_due_soon_notified(pid, row["due_date"])
            log.debug("[DueSoon] page_id %s not in current snapshot (likely Done) — skipped.", pid)
            continue

        name = task.get("name", pid)
        try:
            recipients = _task_recipients(task, cfg, row["db_env_var"])

            if not recipients:
                log.warning("[DueSoon] \"%s\" — no recipients configured, skipping.", name)
                if not dry_run:
                    cache.mark_due_soon_notified(pid, row["due_date"])
                continue

            due_str = task.get("due_date") or "soon"
            subject = f"[Reminder] {name} due {due_str}"
            html    = _build_due_soon_html(task, days_until_due=2)
            send_mis_email(token, sender, recipients, subject, html)
            if not dry_run:
                cache.mark_due_soon_notified(pid, row["due_date"])
            log.info("[DueSoon] \"%s\" → %d recipient(s)", name, len(recipients))

        except Exception as exc:
            msg = f"DueSoon | {name}: {exc}"
            errors.append(msg)
            log.error("[dispatch_due_soon] %s", msg)

    return errors


# ---------------------------------------------------------------------------
# dispatch_slippage — slippage digest (Req 3)
# ---------------------------------------------------------------------------

def dispatch_slippage(slipped_tasks: list, stalled_tasks: list, cfg: dict, token: str) -> list:
    """
    Send a single slippage digest email to executives + leaders + all PMs.
    Returns list of error strings.
    """
    from mailer import _build_slippage_digest_html, send_mis_email

    if not slipped_tasks and not stalled_tasks:
        log.info("[Slippage] No slipped or stalled tasks — nothing to send.")
        return []

    recipients = _merge(
        _executive_emails(cfg),
        _leader_emails(cfg),
        _all_pm_emails(cfg),
    )

    if not recipients:
        log.warning(
            "[Slippage] No recipients in roles.executive / roles.leader / roles.pm. "
            "Add emails to config.json and re-run."
        )
        return []

    today_str = date.today().strftime("%d %b %Y")
    subject   = (
        f"[Slippage Report] {len(slipped_tasks)} task(s) with "
        f"shifted due dates — {today_str}"
    )
    sender = cfg["sender_email"]

    try:
        html = _build_slippage_digest_html(slipped_tasks, stalled_tasks)
        send_mis_email(token, sender, recipients, subject, html)
        log.info(
            "[Slippage] digest sent → %d recipient(s), %d slipped, %d stalled task(s).",
            len(recipients), len(slipped_tasks), len(stalled_tasks),
        )
    except Exception as exc:
        msg = f"Slippage digest: {exc}"
        log.error("[dispatch_slippage] %s", msg)
        return [msg]

    return []


# ---------------------------------------------------------------------------
# dispatch_overdue — overdue digest
# ---------------------------------------------------------------------------

def dispatch_overdue(overdue_tasks: list, stalled_tasks: list, cfg: dict, token: str) -> list:
    """
    Send a single overdue digest email to executives + leaders + all PMs.
    Returns list of error strings.
    """
    from mailer import _build_overdue_digest_html, send_mis_email

    if not overdue_tasks and not stalled_tasks:
        log.info("[Overdue] No overdue or stalled tasks — nothing to send.")
        return []

    recipients = _merge(
        _executive_emails(cfg),
        _leader_emails(cfg),
        _all_pm_emails(cfg),
    )

    if not recipients:
        log.warning(
            "[Overdue] No recipients in roles.executive / roles.leader / roles.pm. "
            "Add emails to config.json and re-run."
        )
        return []

    today_str = date.today().strftime("%d %b %Y")
    subject   = (
        f"[Overdue Report] {len(overdue_tasks)} task(s) past due date — {today_str}"
    )
    sender = cfg["sender_email"]

    try:
        html = _build_overdue_digest_html(overdue_tasks, stalled_tasks)
        send_mis_email(token, sender, recipients, subject, html)
        log.info(
            "[Overdue] digest sent → %d recipient(s), %d overdue, %d stalled task(s).",
            len(recipients), len(overdue_tasks), len(stalled_tasks),
        )
    except Exception as exc:
        msg = f"Overdue digest: {exc}"
        log.error("[dispatch_overdue] %s", msg)
        return [msg]

    return []
