import json
import logging
import os
import re
from datetime import date, datetime
from pathlib import Path

import msal
import requests

log = logging.getLogger(__name__)


def _dry_run_enabled() -> bool:
    return os.environ.get("MAILER_DRY_RUN", "").lower() in ("1", "true", "yes")


def _dry_run_dir() -> Path:
    return Path(os.environ.get("MAILER_DRY_RUN_DIR", "test_emails"))


_DRY_RUN_COUNTER = {"n": 0}


def _safe_slug(s: str, max_len: int = 60) -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", (s or "msg").strip())
    return s[:max_len].strip("_") or "msg"


def _graph_send(sender_email: str, token: str, payload: dict) -> None:
    """
    Send a Graph sendMail payload, or — when MAILER_DRY_RUN=1 — write the
    rendered HTML to MAILER_DRY_RUN_DIR (default: ./test_emails) and skip the
    network call. All mailer entry points route through here so a test run
    cannot accidentally fire a real email.
    """
    if _dry_run_enabled():
        out_dir = _dry_run_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        _DRY_RUN_COUNTER["n"] += 1
        msg = payload.get("message", {})
        subject = msg.get("subject", "no-subject")
        recipients = [
            r.get("emailAddress", {}).get("address", "")
            for r in msg.get("toRecipients", [])
        ]
        ts = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        fname = f"{ts}_{_DRY_RUN_COUNTER['n']:03d}_{_safe_slug(subject)}.html"
        out_path = out_dir / fname

        meta_banner = (
            f"<!-- DRY RUN — would have sent via {sender_email} -->\n"
            f'<div style="background:#fef3c7;border:1px solid #f59e0b;'
            f'padding:8px 12px;font-family:monospace;font-size:12px;color:#78350f;">'
            f"<strong>DRY RUN</strong> · subject: {subject} · "
            f"to: {', '.join(recipients) or '—'} · from: {sender_email}"
            f"</div>\n"
        )
        body_html = msg.get("body", {}).get("content", "")
        out_path.write_text(meta_banner + body_html, encoding="utf-8")
        log.info(
            "[DRY RUN] wrote %s (to=%s, subject=%r)",
            out_path,
            ",".join(recipients),
            subject,
        )
        return

    resp = requests.post(
        _GRAPH_SEND_MAIL.format(sender=sender_email),
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
    )
    resp.raise_for_status()


# MIS template accent colours (one per notification type)
_MIS_ACCENT = {
    "created": "#10b981",  # green   — new task assigned
    "reassigned": "#3b82f6",  # blue    — task reassigned
    "slippage": "#dc2626",  # red     — due-date slippage
    "started": "#8b5cf6",  # purple  — project / task started
    "due_soon": "#f59e0b",  # amber   — 2-day due reminder
}

_GRAPH_SEND_MAIL = "https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
_SCOPES = ["https://graph.microsoft.com/.default"]

# Sections whose report value is None are skipped automatically.
_ALL_SECTIONS = [
    ("no_due_date", "No Due Date", "#f59e0b", "Tasks with no due date assigned"),
    ("no_owner", "No Owner", "#8b5cf6", "Tasks with no owner assigned"),
    ("no_reviewer", "No Reviewer", "#3b82f6", "Tasks with no reviewer assigned"),
    ("overdue", "Overdue", "#ef4444", "Past due date and not yet Done"),
    ("max_slippage", "Maximum Slippage", "#dc2626", "Top tasks ranked by days slipped"),
]

_PRIORITY_COLORS = {"High": "#ef4444", "Medium": "#f59e0b", "Low": "#10b981"}


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


def _acquire_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    if _dry_run_enabled():
        return "DRY_RUN_TOKEN"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=_SCOPES)
    if "access_token" not in result:
        raise RuntimeError(
            f"Token acquisition failed: {result.get('error_description', result.get('error'))}"
        )
    return result["access_token"]


# ---------------------------------------------------------------------------
# Grouping
# ---------------------------------------------------------------------------


def _group_tasks(tasks: list[dict]) -> list[tuple]:
    """
    Returns [(proj_name, proj_url, direct_tasks, parent_groups)]
    where parent_groups = [(parent_name, parent_url, tasks)]
    Tasks with no parent sit in direct_tasks.
    Projects sorted by task count desc; "No Project" last.
    """
    groups: dict = {}
    for task in tasks:
        proj = task.get("project") or {}
        par = task.get("parent_task") or {}

        pk = proj.get("url") or "__none__"
        rk = par.get("url") or "__none__"

        if pk not in groups:
            groups[pk] = {
                "name": proj.get("name") or "No Project",
                "url": proj.get("url") or "",
                "direct": [],
                "parents": {},
            }

        if rk == "__none__":
            groups[pk]["direct"].append(task)
        else:
            pars = groups[pk]["parents"]
            if rk not in pars:
                pars[rk] = {
                    "name": par.get("name") or "",
                    "url": par.get("url") or "",
                    "tasks": [],
                }
            pars[rk]["tasks"].append(task)

    def _total(g):
        return len(g["direct"]) + sum(len(p["tasks"]) for p in g["parents"].values())

    result = []
    for _, g in sorted(
        groups.items(), key=lambda x: (x[0] == "__none__", -_total(x[1]))
    ):
        parent_groups = sorted(
            [(p["name"], p["url"], p["tasks"]) for p in g["parents"].values()],
            key=lambda x: -len(x[2]),
        )
        result.append((g["name"], g["url"], g["direct"], parent_groups))
    return result


# ---------------------------------------------------------------------------
# HTML primitives
# ---------------------------------------------------------------------------


def _badge(text: str, color: str, small: bool = False) -> str:
    import html as _html

    sz = "11px" if small else "12px"
    return (
        f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
        f'background:{color};color:#fff;font-size:{sz};font-weight:700;white-space:nowrap;">'
        f"{_html.escape(str(text))}</span>"
    )


def _task_row(
    task: dict, color: str, show_slippage: bool = False, indent: bool = False
) -> str:
    chips = []
    if task.get("status"):
        chips.append(_badge(task["status"], "#6b7280", small=True))
    if task.get("priority"):
        chips.append(
            _badge(
                task["priority"],
                _PRIORITY_COLORS.get(task["priority"], "#9ca3af"),
                small=True,
            )
        )
    for team in (task.get("teams") or [])[:2]:
        chips.append(_badge(team, "#0ea5e9", small=True))

    chips_html = (
        (
            "<div style='margin-top:4px;'>"
            + "".join(f'<span style="margin-right:4px;">{c}</span>' for c in chips)
            + "</div>"
        )
        if chips
        else ""
    )

    desc_html = (
        (
            f'<div style="font-size:11px;color:#6b7280;margin-top:3px;">'
            f"{task['description']}{'…' if len(task['description']) >= 120 else ''}</div>"
        )
        if task.get("description")
        else ""
    )

    slip_html = (
        (
            f'<div style="font-size:11px;color:#dc2626;margin-top:3px;">'
            f"Originally: {task.get('original_date', '?')} "
            f"&#8594; slipped <strong>{task['slippage_days']} day(s)</strong></div>"
        )
        if (show_slippage and task.get("slippage_days"))
        else ""
    )

    crumbs = []
    if task.get("project"):
        p = task["project"]
        crumbs.append(
            f'<a href="{p["url"]}" style="color:#d1d5db;text-decoration:none;">{p["name"]}</a>'
            if p.get("url")
            else f'<span style="color:#d1d5db;">{p["name"]}</span>'
        )
    if task.get("parent_task"):
        pt = task["parent_task"]
        crumbs.append(
            f'<a href="{pt["url"]}" style="color:#d1d5db;text-decoration:none;">{pt["name"]}</a>'
            if pt.get("url")
            else f'<span style="color:#d1d5db;">{pt["name"]}</span>'
        )
    crumb_html = (
        (
            f'<div style="font-size:10px;margin-bottom:2px;">'
            + " &rsaquo; ".join(crumbs)
            + "</div>"
        )
        if crumbs
        else ""
    )

    left_pad = "28px" if indent else "12px"
    due = task.get("due_date") or "—"
    return (
        f'<tr style="border-bottom:1px solid #f3f4f6;">'
        f'<td style="padding:10px 12px 10px {left_pad};border-left:3px solid {color};vertical-align:top;">'
        f"{crumb_html}"
        f'<a href="{task["url"]}" style="color:#111827;text-decoration:none;'
        f'font-size:13px;font-weight:600;">{task["name"]}</a>'
        f"{chips_html}{desc_html}{slip_html}"
        f"</td>"
        f'<td style="font-size:12px;color:#6b7280;white-space:nowrap;'
        f'padding:10px 12px;vertical-align:top;width:90px;">{due}</td>'
        f"</tr>"
    )


def _project_header(name: str, url: str, count: int, color: str) -> str:
    label = (
        f'<a href="{url}" style="color:#111827;text-decoration:none;font-weight:700;">{name}</a>'
        if url
        else f'<span style="font-weight:700;">{name}</span>'
    )
    return (
        f"<tr>"
        f'<td colspan="2" style="padding:8px 14px;background:#f3f4f6;'
        f'border-left:4px solid {color};border-bottom:1px solid #e5e7eb;">'
        f'<span style="font-size:13px;color:#111827;">{label}</span>'
        f"&nbsp;&nbsp;{_badge(str(count), color)}"
        f"</td>"
        f"</tr>"
    )


def _parent_header(name: str, url: str, count: int) -> str:
    label = (
        f'<a href="{url}" style="color:#374151;text-decoration:none;">{name}</a>'
        if url
        else f'<span style="color:#374151;">{name}</span>'
    )
    return (
        f"<tr>"
        f'<td colspan="2" style="padding:6px 12px 6px 28px;background:#fafafa;'
        f'border-bottom:1px solid #f3f4f6;">'
        f'<span style="font-size:12px;font-weight:600;">&#8627; {label}</span>'
        f'&nbsp;<span style="font-size:11px;color:#9ca3af;">({count} tasks)</span>'
        f"</td>"
        f"</tr>"
    )


def _section_card(
    key: str, label: str, color: str, desc: str, tasks: list[dict]
) -> str:
    if not tasks:
        rows = (
            '<tr><td colspan="2" style="padding:16px 12px;color:#9ca3af;font-size:13px;">'
            "&#10003;&nbsp;All clear</td></tr>"
        )
    else:
        show_slip = key == "max_slippage"
        row_parts = []
        for proj_name, proj_url, direct, parent_groups in _group_tasks(tasks):
            total = len(direct) + sum(len(pg[2]) for pg in parent_groups)
            row_parts.append(_project_header(proj_name, proj_url, total, color))
            for t in direct:
                row_parts.append(
                    _task_row(t, color, show_slippage=show_slip, indent=False)
                )
            for par_name, par_url, par_tasks in parent_groups:
                row_parts.append(_parent_header(par_name, par_url, len(par_tasks)))
                for t in par_tasks:
                    row_parts.append(
                        _task_row(t, color, show_slippage=show_slip, indent=True)
                    )
        rows = "\n".join(row_parts)

    count_color = color if tasks else "#10b981"
    return f"""
<table width="100%" cellpadding="0" cellspacing="0"
       style="margin-bottom:24px;border-radius:8px;border:1px solid #e5e7eb;overflow:hidden;">
  <tr style="background:#f9fafb;">
    <td style="padding:12px 16px;">
      <span style="font-size:14px;font-weight:700;color:#111827;">{label}</span>
      <span style="font-size:12px;color:#9ca3af;margin-left:8px;">{desc}</span>
    </td>
    <td style="padding:12px 16px;text-align:right;white-space:nowrap;width:60px;">
      {_badge(str(len(tasks)), count_color)}
    </td>
  </tr>
  <tr><td colspan="2" style="padding:0;">
    <table width="100%" cellpadding="0" cellspacing="0">{rows}</table>
  </td></tr>
</table>"""


# ---------------------------------------------------------------------------
# Email assembly
# ---------------------------------------------------------------------------


def _build_html(report: dict) -> str:
    today_str = date.today().strftime("%B %d, %Y")
    db_name = report.get("db_name", "Notion")
    total = report["total_open"]

    active = [
        (key, label, color, desc)
        for key, label, color, desc in _ALL_SECTIONS
        if report.get(key) is not None
    ]

    stats = "".join(
        f'<td style="text-align:center;padding:16px 8px;">'
        f'<div style="font-size:28px;font-weight:800;color:{color};">{len(report[key])}</div>'
        f'<div style="font-size:11px;color:#9ca3af;margin-top:2px;">{label}</div>'
        f"</td>"
        for key, label, color, _ in active
    )

    sections_html = "".join(
        _section_card(key, label, color, desc, report[key])
        for key, label, color, desc in active
    )

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="background:#f3f4f6;padding:24px 0;">
<tr><td align="center" style="padding:0 12px;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="max-width:1200px;background:#fff;border-radius:12px;overflow:hidden;
              box-shadow:0 1px 4px rgba(0,0,0,.08);">

  <tr>
    <td style="background:#111827;padding:28px 32px;">
      <div style="font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;
                  letter-spacing:.08em;margin-bottom:6px;">Task Health Report</div>
      <div style="font-size:22px;font-weight:800;color:#fff;">{db_name}</div>
      <div style="font-size:13px;color:#9ca3af;margin-top:6px;">
        {today_str}&nbsp;&nbsp;&#183;&nbsp;&nbsp;{total} open tasks
      </div>
    </td>
  </tr>

  <tr>
    <td style="padding:0;border-bottom:2px solid #e5e7eb;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>{stats}</tr></table>
    </td>
  </tr>

  <tr><td style="padding:24px 24px 8px;">{sections_html}</td></tr>

  <tr>
    <td style="padding:16px 24px 28px;text-align:center;font-size:11px;color:#9ca3af;">
      Auto-generated from Notion &mdash; excludes tasks marked <em>Done</em>
      and blank/untitled entries.
    </td>
  </tr>

</table>
</td></tr>
</table>
</body></html>"""


# ---------------------------------------------------------------------------
# Owner daily digest (Actionable Messages)
# ---------------------------------------------------------------------------

_PRIORITY_BORDER = {"High": "#ef4444", "Medium": "#f59e0b", "Low": "#10b981"}


def _digest_task_row(task: dict) -> str:
    """HTML fallback row for one in-progress task."""
    due = task.get("due_date") or "—"
    pri = task.get("priority", "")
    border = _PRIORITY_BORDER.get(pri, "#6b7280")
    pri_badge = (
        (
            f'<span style="display:inline-block;padding:1px 7px;border-radius:9px;'
            f"background:{_PRIORITY_COLORS.get(pri, '#9ca3af')};color:#fff;font-size:11px;"
            f'margin-left:6px;">{pri}</span>'
        )
        if pri
        else ""
    )
    db_badge = (
        f'<span style="display:inline-block;padding:1px 7px;border-radius:9px;'
        f'background:#e5e7eb;color:#374151;font-size:11px;margin-left:4px;">'
        f"{task['db_name']}</span>"
    )
    return (
        f'<tr style="border-bottom:1px solid #f3f4f6;">'
        f'<td style="padding:10px 14px;border-left:3px solid {border};vertical-align:middle;">'
        f'<a href="{task["url"]}" style="color:#111827;font-size:13px;font-weight:600;'
        f'text-decoration:none;">{task["name"]}</a>'
        f"{pri_badge}{db_badge}"
        f"</td>"
        f'<td style="padding:10px 14px;font-size:12px;color:#6b7280;white-space:nowrap;">{due}</td>'
        f'<td style="padding:10px 14px;font-size:12px;color:#9ca3af;font-style:italic;'
        f'white-space:nowrap;">Log via Outlook card</td>'
        f"</tr>"
    )


def _build_digest_html(owner_name: str, tasks: list[dict], today_str: str) -> str:
    """HTML fallback body for non-Outlook clients."""
    rows = "\n".join(_digest_task_row(t) for t in tasks)
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:24px 0;">
<tr><td align="center" style="padding:0 12px;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="max-width:700px;background:#fff;border-radius:12px;overflow:hidden;
              box-shadow:0 1px 4px rgba(0,0,0,.08);">
  <tr>
    <td style="background:#111827;padding:24px 28px;">
      <div style="font-size:11px;font-weight:600;color:#6b7280;text-transform:uppercase;
                  letter-spacing:.08em;margin-bottom:4px;">Daily Hours Log</div>
      <div style="font-size:20px;font-weight:800;color:#fff;">Hi {owner_name},</div>
      <div style="font-size:13px;color:#9ca3af;margin-top:4px;">
        {today_str} &nbsp;&#183;&nbsp; {len(tasks)} task(s) in progress
      </div>
    </td>
  </tr>
  <tr><td style="padding:20px 24px;">
    <p style="font-size:13px;color:#374151;margin:0 0 16px;">
      Please log your hours for each task below. Open this email in
      <strong>Outlook</strong> to log hours interactively, or visit each
      task in Notion directly.
    </p>
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
      <tr style="background:#f9fafb;">
        <th style="padding:10px 14px;text-align:left;font-size:12px;color:#6b7280;font-weight:600;">Task</th>
        <th style="padding:10px 14px;text-align:left;font-size:12px;color:#6b7280;font-weight:600;width:90px;">Due</th>
        <th style="padding:10px 14px;text-align:left;font-size:12px;color:#6b7280;font-weight:600;width:130px;">Hours</th>
      </tr>
      {rows}
    </table>
  </td></tr>
  <tr>
    <td style="padding:12px 24px 24px;text-align:center;font-size:11px;color:#9ca3af;">
      Auto-generated daily digest &mdash; open in Outlook to log hours inline.
    </td>
  </tr>
</table>
</td></tr>
</table>
</body></html>"""


def _task_facts(task: dict, due: str, pri: str) -> list[dict]:
    """Build the FactSet rows for one task, including project/parent chain."""
    facts = []
    # Breadcrumb: Project › Parent Task
    crumbs = []
    if task.get("project"):
        crumbs.append(task["project"]["name"])
    if task.get("parent_task"):
        crumbs.append(task["parent_task"]["name"])
    if crumbs:
        facts.append({"title": "Path", "value": " › ".join(crumbs)})
    facts += [
        {"title": "Due", "value": due},
        {"title": "Priority", "value": pri},
        {"title": "DB", "value": task["db_name"]},
    ]
    return facts


def _build_digest_card(
    owner_name: str,
    tasks: list[dict],
    app_base_url: str,
    today_str: str,
    originator_id: str = "",
) -> dict:
    """
    Build an Adaptive Card for the daily digest.

    Uses Action.ShowCard so each task has its own expandable Log Hours button
    with an independent number input — identical UX to the old MessageCard
    ActionCard pattern but in the format Exchange Online actually recognises.
    """
    # Top-level body: one row per task showing name + metadata
    body_items = [
        {
            "type": "TextBlock",
            "size": "Medium",
            "weight": "Bolder",
            "text": f"Hi {owner_name} — {today_str}",
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": f"You have **{len(tasks)}** in-progress task(s). "
            "Tap **Log Hours** next to each task to record your time.",
            "wrap": True,
            "isSubtle": True,
            "spacing": "Small",
        },
        {"type": "Separator"},
    ]

    # Build one Input.Number per task (all visible — no ShowCard expansion needed)
    # Input ID uses the task UUID with dashes so the server can reconstruct the
    # task_id by stripping the "h_" prefix.  Task names travel as "n_UUID" static
    # fields so the server can log them without a separate Notion lookup.
    hours_refs = {}  # {input_id: template_placeholder}
    name_refs = {}  # {name_key:  task_name}

    for task in tasks:
        due = task.get("due_date") or "No due date"
        pri = task.get("priority") or "—"
        inp_id = f"h_{task['id']}"
        name_key = f"n_{task['id']}"
        hours_refs[inp_id] = f"{{{{{inp_id}.value}}}}"
        name_refs[name_key] = task["name"]

        body_items.append(
            {
                "type": "Container",
                "spacing": "Small",
                "separator": True,
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "weight": "Bolder",
                                        "text": task["name"],
                                        "wrap": True,
                                    },
                                    {
                                        "type": "FactSet",
                                        "spacing": "Small",
                                        "facts": _task_facts(task, due, pri),
                                    },
                                ],
                            },
                            {
                                "type": "Column",
                                "width": "auto",
                                "verticalContentAlignment": "Center",
                                "items": [
                                    {
                                        "type": "Input.Number",
                                        "id": inp_id,
                                        "label": "hrs",
                                        "placeholder": "0",
                                        "value": "0",
                                        "min": 0,
                                        "max": 12,
                                        "isRequired": True,
                                    }
                                ],
                            },
                        ],
                    },
                ],
            }
        )

    # Single Submit All button at the bottom — body carries all input values
    # plus static task-name lookup so the server needs no extra API calls
    post_body = json.dumps(
        {"date": today_str, **hours_refs, **name_refs},
        ensure_ascii=True,
    )

    card: dict = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "hideOriginalBody": True,
        "body": body_items,
        "actions": [
            {
                "type": "Action.Http",
                "title": "Submit All Hours",
                "method": "POST",
                "url": f"{app_base_url}/log-hours-action",
                "headers": [{"name": "Content-Type", "value": "application/json"}],
                "body": post_body,
            }
        ],
    }
    if originator_id:
        card["originator"] = originator_id

    return card


def build_owner_digest_email(
    owner_name: str,
    tasks: list[dict],
    app_base_url: str,
    today_str: str,
    originator_id: str = "",
) -> dict:
    """
    Returns {"subject": str, "html": str, "card": dict} for one owner's digest.

    The Adaptive Card is embedded in the HTML <head> as:
        <script type="application/adaptivecard+json">...</script>
    and sent via the Graph API JSON endpoint (not raw MIME).
    Exchange Online preserves this script tag and renders the interactive
    card in Outlook; other clients see the HTML table fallback.
    """
    card = _build_digest_card(owner_name, tasks, app_base_url, today_str, originator_id)
    html_body = _build_digest_html(owner_name, tasks, today_str)

    # Embed card in <head> using the correct script type for Adaptive Cards.
    # Must be application/adaptivecard+json (NOT application/ld+json which is
    # legacy MessageCard only). Graph API JSON endpoint preserves this tag.
    card_script = (
        '<script type="application/adaptivecard+json">'
        + json.dumps(card, ensure_ascii=True)
        + "</script>"
    )
    html_with_card = html_body.replace("</head>", card_script + "\n</head>", 1)

    return {
        "subject": f"Daily Hours Log — {today_str} ({len(tasks)} task(s))",
        "html": html_with_card,
        "card": card,
    }


def send_owner_daily_digests(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sender_email: str,
    owner_map: dict[str, dict],
    app_base_url: str,
    originator_id: str = "",
) -> list[str]:
    """
    Send one digest email per owner in owner_map.
    owner_map: {email: {"owner_name": str, "tasks": [...]}}
    Returns list of error strings (empty = all sent successfully).
    """
    token = _acquire_token(tenant_id, client_id, client_secret)
    today_str = date.today().strftime("%d %b %Y")
    errors: list[str] = []

    for owner_email, info in owner_map.items():
        digest = build_owner_digest_email(
            owner_name=info["owner_name"],
            tasks=info["tasks"],
            app_base_url=app_base_url,
            today_str=today_str,
            originator_id=originator_id,
        )
        payload = {
            "message": {
                "subject": digest["subject"],
                "body": {"contentType": "HTML", "content": digest["html"]},
                "toRecipients": [{"emailAddress": {"address": owner_email}}],
            }
        }
        try:
            _graph_send(sender_email, token, payload)
        except Exception as exc:
            errors.append(f"{owner_email}: {exc}")

    return errors


# ---------------------------------------------------------------------------
# Hours submission confirmation email
# ---------------------------------------------------------------------------


def send_hours_confirmation(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sender_email: str,
    recipient_email: str,
    outcome: str,  # "ok" | "overwrite" | "cap_exceeded"
    entries: list[dict],  # [{task_name, hours}]
    daily_total: float,
    daily_cap: float,
    date_str: str,
    cap_error: str = "",
) -> None:
    """Send a brief confirmation email after a hours submission."""
    if outcome == "cap_exceeded":
        subject = f"⚠ Hours Not Logged — {date_str}"
        color = "#ef4444"
        headline = "Submission rejected — daily cap exceeded"
        body_lines = [
            f'<p style="color:#374151;font-size:13px;margin:0 0 12px;">{cap_error}</p>',
            f'<p style="color:#6b7280;font-size:12px;margin:0;">',
            f"Please re-submit with a lower total.",
            f"</p>",
        ]
    else:
        overwrote = outcome == "overwrite"
        subject = f"✓ Hours {'Updated' if overwrote else 'Logged'} — {date_str}"
        color = "#f59e0b" if overwrote else "#10b981"
        headline = (
            "Previous entries updated"
            if overwrote
            else f"{sum(e['hours'] for e in entries):.1f} h logged across {len(entries)} task(s)"
        )
        rows = "".join(
            f'<tr style="border-bottom:1px solid #f3f4f6;">'
            f'<td style="padding:7px 12px;font-size:13px;color:#111827;">{e["task_name"]}</td>'
            f'<td style="padding:7px 12px;font-size:13px;color:#374151;text-align:right;'
            f'white-space:nowrap;">{e["hours"]:.1f} h</td>'
            f"</tr>"
            for e in entries
        )
        cap_pct = min(int(daily_total / daily_cap * 100), 100)
        bar_color = "#10b981" if daily_total <= daily_cap * 0.75 else "#f59e0b"
        body_lines = [
            f'<table width="100%" cellpadding="0" cellspacing="0"'
            f' style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;margin-bottom:14px;">',
            f'<tr style="background:#f9fafb;"><th style="padding:8px 12px;text-align:left;'
            f'font-size:11px;color:#6b7280;font-weight:600;">Task</th>'
            f'<th style="padding:8px 12px;text-align:right;font-size:11px;color:#6b7280;'
            f'font-weight:600;width:60px;">Hours</th></tr>',
            rows,
            f"</table>",
            f'<p style="font-size:12px;color:#6b7280;margin:0 0 6px;">',
            f"Daily total: <strong>{daily_total:.1f} h</strong> / {daily_cap:.0f} h cap</p>",
            f'<div style="background:#f3f4f6;border-radius:4px;height:6px;overflow:hidden;">',
            f'<div style="background:{bar_color};width:{cap_pct}%;height:6px;"></div></div>',
        ]

    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:20px 0;">
<tr><td align="center" style="padding:0 12px;">
<table width="100%" cellpadding="0" cellspacing="0"
       style="max-width:480px;background:#fff;border-radius:10px;overflow:hidden;
              box-shadow:0 1px 4px rgba(0,0,0,.08);">
  <tr>
    <td style="background:{color};padding:16px 20px;">
      <div style="font-size:11px;font-weight:600;color:rgba(255,255,255,.75);
                  text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px;">
        Hours Log Confirmation</div>
      <div style="font-size:16px;font-weight:700;color:#fff;">{headline}</div>
      <div style="font-size:12px;color:rgba(255,255,255,.8);margin-top:3px;">{date_str}</div>
    </td>
  </tr>
  <tr><td style="padding:16px 20px;">
    {"".join(body_lines)}
  </td></tr>
</table>
</td></tr>
</table></body></html>"""

    token = _acquire_token(tenant_id, client_id, client_secret)
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": recipient_email}}],
        }
    }
    _graph_send(sender_email, token, payload)


def send_reminder_email(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sender_email: str,
    recipient_emails: list[str],
    report: dict,
) -> None:
    token = _acquire_token(tenant_id, client_id, client_secret)
    active_keys = [k for k, *_ in _ALL_SECTIONS if report.get(k) is not None]
    total_issues = sum(len(report[k]) for k in active_keys)
    db_name = report.get("db_name", "Notion")
    payload = {
        "message": {
            "subject": (
                f"[{db_name}] Task Health Report — "
                f"{total_issues} item(s) need attention "
                f"({date.today().strftime('%d %b %Y')})"
            ),
            "body": {"contentType": "HTML", "content": _build_html(report)},
            "toRecipients": [
                {"emailAddress": {"address": r}} for r in recipient_emails
            ],
        }
    }
    _graph_send(sender_email, token, payload)


# ---------------------------------------------------------------------------
# MIS generic send helper
# ---------------------------------------------------------------------------


def send_mis_email(
    token: str,
    sender_email: str,
    recipient_emails: list,
    subject: str,
    html: str,
) -> None:
    """
    Send a single HTML email using a pre-acquired Graph API token.
    The token is acquired once per mis_runner.py run and reused for all sends.
    """
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [
                {"emailAddress": {"address": r}} for r in recipient_emails
            ],
        }
    }
    _graph_send(sender_email, token, payload)


# ============================================================================
# MIS NOTIFICATION TEMPLATES
# Req 1 — Task Assigned/Reassigned  →  _build_task_assigned_html
# Req 2 — Daily owner digest        →  _build_digest_html  (already exists)
# Req 3 — Slippage digest           →  _build_slippage_digest_html
# Req 4 — Project/Task started      →  _build_project_started_html
# Req 5 — Due-in-2-days reminder    →  _build_due_soon_html
#
# Design tokens (match existing Task Health Report):
#   page bg         #f3f4f6
#   card            white, border-radius:12px, box-shadow:0 1px 4px rgba(0,0,0,.08)
#   header bg       #111827
#   stat numbers    font-size:22px, font-weight:800
#   section cards   border:1px solid #e5e7eb, border-radius:8px
#   task-row accent border-left:4px solid {accent}
#   badges          _badge() / _PRIORITY_COLORS  (reused from above)
# ============================================================================


# ---------------------------------------------------------------------------
# Shared inner helpers  (used only by the MIS templates below)
# ---------------------------------------------------------------------------


def _mis_dark_header(
    eyebrow: str,
    eyebrow_color: str,
    label_text: str,
    label_color: str,
    title: str,
    subtitle: str,
) -> str:
    """Dark #111827 header block with an accent pill label above the title."""
    return (
        f"<tr>"
        f'<td style="background:#111827;padding:24px 28px;">'
        f'<div style="display:inline-block;padding:2px 10px;border-radius:10px;'
        f"background:{label_color};color:#fff;font-size:11px;font-weight:700;"
        f'text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">{label_text}</div>'
        f'<div style="font-size:11px;font-weight:600;color:{eyebrow_color};'
        f'text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">{eyebrow}</div>'
        f'<div style="font-size:22px;font-weight:800;color:#fff;line-height:1.3;">{title}</div>'
        f'<div style="font-size:12px;color:#9ca3af;margin-top:6px;">{subtitle}</div>'
        f"</td></tr>"
    )


def _mis_stat_strip(cells: list) -> str:
    """
    Stat strip row with [(value, label, color), ...] tuples.
    Matches the big-number strip in the Task Health Report.
    Long text values use a smaller font so they don't overflow.
    """
    import html as _html

    inner = ""
    n = len(cells)
    for i, (value, label, color) in enumerate(cells):
        border = "border-right:1px solid #f3f4f6;" if i < n - 1 else ""
        # Shrink font for text values (dates, names) vs short numbers/codes
        is_long = len(str(value)) > 6
        fsize = "14px" if is_long else "22px"
        inner += (
            f'<td style="text-align:center;padding:14px 8px;{border}width:{100 // n}%;">'
            f'<div style="font-size:{fsize};font-weight:800;color:{color};'
            f'word-break:break-word;line-height:1.2;">'
            f"{_html.escape(str(value))}</div>"
            f'<div style="font-size:11px;color:#9ca3af;margin-top:3px;">'
            f"{_html.escape(label)}</div>"
            f"</td>"
        )
    return (
        f'<tr><td style="border-bottom:2px solid #e5e7eb;padding:0;">'
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>{inner}</tr></table>'
        f"</td></tr>"
    )


def _mis_section_card(title: str, badge_val: str, accent: str, inner_html: str) -> str:
    """Section card matching _section_card() style (header row + content)."""
    badge_html = f"&nbsp;{_badge(badge_val, accent)}" if badge_val else ""
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="margin-bottom:20px;border-radius:8px;border:1px solid #e5e7eb;overflow:hidden;">'
        f'<tr style="background:#f9fafb;">'
        f'<td style="padding:12px 16px;border-bottom:1px solid #e5e7eb;">'
        f'<span style="font-size:13px;font-weight:700;color:#111827;">{title}</span>'
        f"{badge_html}"
        f"</td></tr>"
        f'<tr><td style="padding:14px 16px;">{inner_html}</td></tr>'
        f"</table>"
    )


def _mis_crumb(task: dict) -> str:
    """Project › Parent Task breadcrumb line (dim, matches existing _task_row crumbs)."""
    crumbs = []
    if task.get("project"):
        p = task["project"]
        crumbs.append(
            f'<a href="{p["url"]}" style="color:#9ca3af;text-decoration:none;">{p["name"]}</a>'
            if p.get("url")
            else f'<span style="color:#9ca3af;">{p["name"]}</span>'
        )
    if task.get("parent_task"):
        pt = task["parent_task"]
        crumbs.append(
            f'<a href="{pt["url"]}" style="color:#9ca3af;text-decoration:none;">{pt["name"]}</a>'
            if pt.get("url")
            else f'<span style="color:#9ca3af;">{pt["name"]}</span>'
        )
    return (
        (
            f'<div style="font-size:10px;margin-bottom:4px;">'
            + " &rsaquo; ".join(crumbs)
            + "</div>"
        )
        if crumbs
        else ""
    )


def _mis_chips(task: dict) -> str:
    """Status + Priority + Team badges, matching _task_row chip style."""
    chips = []
    if task.get("status"):
        chips.append(_badge(task["status"], "#6b7280", small=True))
    if task.get("priority"):
        chips.append(
            _badge(
                task["priority"],
                _PRIORITY_COLORS.get(task["priority"], "#9ca3af"),
                small=True,
            )
        )
    for team in (task.get("teams") or [])[:3]:
        chips.append(_badge(team, "#0ea5e9", small=True))
    return (
        (
            '<div style="margin-top:6px;">'
            + "".join(f'<span style="margin-right:4px;">{c}</span>' for c in chips)
            + "</div>"
        )
        if chips
        else ""
    )


def _mis_cta(label: str, url: str, color: str) -> str:
    return (
        f'<a href="{url}" style="display:inline-block;margin-top:14px;padding:10px 22px;'
        f"background:{color};color:#fff;border-radius:6px;text-decoration:none;"
        f'font-size:13px;font-weight:700;">{label} &rarr;</a>'
    )


def _mis_email_shell(
    header: str, stat_strip: str, body: str, footer: str, max_width: str = "680px"
) -> str:
    """Outer wrapper that exactly mirrors _build_html()'s page + card chrome."""
    return (
        f'<!DOCTYPE html><html><head><meta charset="UTF-8">'
        f'<meta name="viewport" content="width=device-width,initial-scale=1">'
        f"</head>"
        f'<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,sans-serif;">'
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="background:#f3f4f6;padding:24px 0;">'
        f'<tr><td align="center" style="padding:0 12px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="max-width:{max_width};background:#fff;border-radius:12px;overflow:hidden;'
        f'box-shadow:0 1px 4px rgba(0,0,0,.08);">'
        f"{header}"
        f"{stat_strip}"
        f'<tr><td style="padding:20px 24px 8px;">{body}</td></tr>'
        f'<tr><td style="padding:10px 24px 24px;text-align:center;'
        f'font-size:11px;color:#9ca3af;">{footer}</td></tr>'
        f"</table></td></tr></table></body></html>"
    )


# ---------------------------------------------------------------------------
# T1 — Task Assigned / Reassigned
#
# Sent to:  Assignee  +  Admins  +  Project PM
# Trigger:  New page in DB  OR  Assignee field changed
#
# task dict must include:
#   name, url, due_date, status, priority, teams (list), description,
#   project ({name,url}|None), parent_task ({name,url}|None),
#   db_name, assignees ([{name, email}])
# reason: "created" | "reassigned"
# ---------------------------------------------------------------------------


def _build_task_assigned_html(task: dict, reason: str) -> str:
    accent = _MIS_ACCENT.get(reason, _MIS_ACCENT["created"])
    label_text = "New Task" if reason == "created" else "Reassigned"
    eyebrow = "Task Assignment"
    today_str = date.today().strftime("%d %b %Y")

    assignees = task.get("assignees") or []
    assignee_str = ", ".join(a.get("name", "") for a in assignees) or "—"
    due = task.get("due_date") or "—"
    priority = task.get("priority") or "—"
    pri_color = _PRIORITY_COLORS.get(priority, "#9ca3af")
    db_name = task.get("db_name", "Notion")

    # Description block (indented, matches existing style)
    desc_raw = task.get("description") or ""
    desc_html = (
        (
            f'<div style="font-size:12px;color:#6b7280;margin-top:10px;padding:10px 12px;'
            f'background:#f9fafb;border-radius:6px;border-left:3px solid #e5e7eb;">'
            f"{desc_raw}{'&hellip;' if len(desc_raw) >= 120 else ''}</div>"
        )
        if desc_raw
        else ""
    )

    header = _mis_dark_header(
        eyebrow=eyebrow,
        eyebrow_color="#9ca3af",
        label_text=label_text,
        label_color=accent,
        title=task["name"],
        subtitle=f"{db_name}&nbsp;&nbsp;&#183;&nbsp;&nbsp;{today_str}",
    )

    stat_strip = _mis_stat_strip(
        [
            (due, "Due Date", accent),
            (priority, "Priority", pri_color if priority != "—" else "#9ca3af"),
            (assignee_str, "Assigned To", "#111827"),
        ]
    )

    task_card = (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="margin-bottom:16px;border-radius:8px;border:1px solid #e5e7eb;overflow:hidden;">'
        f'<tr style="background:#f9fafb;">'
        f'<td style="padding:12px 16px;border-bottom:1px solid #e5e7eb;">'
        f'<span style="font-size:13px;font-weight:700;color:#111827;">Task Details</span>'
        f"</td></tr>"
        f'<tr><td style="padding:14px 16px;border-left:4px solid {accent};">'
        f"{_mis_crumb(task)}"
        f'<a href="{task["url"]}" style="color:#111827;font-size:14px;font-weight:700;'
        f'text-decoration:none;">{task["name"]}</a>'
        f"{_mis_chips(task)}"
        f"{desc_html}"
        f"<div>{_mis_cta('Open in Notion', task['url'], accent)}</div>"
        f"</td></tr></table>"
    )

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=task_card,
        footer="Sent via MIS &mdash; Task tracking notification",
        max_width="640px",
    )


# ---------------------------------------------------------------------------
# T3 — Slippage Digest
#
# Sent to:  Project Managers  +  Leaders  +  Executives
# Trigger:  Daily scheduled scan
#
# slipped_tasks: sorted by slippage_days desc, each dict has:
#   name, url, due_date, original_date, slippage_days,
#   all_dates ([str] newest-first), n_reschedules (int),
#   overdue (bool), days_overdue (int),
#   status, priority, teams ([str]),
#   owner_name (str), reviewer_names (str),
#   blocking ([{name,url}]|[]),
#   project ({name,url}|None), db_name
# ---------------------------------------------------------------------------


def _slip_trail_html(task: dict) -> str:
    """Date-hop trail box shown inside each slippage task card."""
    import html as _html

    all_dates = task.get("all_dates") or []
    slippage_days = task.get("slippage_days", 0)
    days_overdue = task.get("days_overdue", 0)
    n_resc = max(len(all_dates) - 1, 0)
    is_overdue = bool(task.get("overdue") or days_overdue > 0)

    accent = "#dc2626" if is_overdue else "#f59e0b"
    bg = "#fff1f2" if is_overdue else "#fffbeb"
    label_color = "#991b1b" if is_overdue else "#92400e"
    pill_border = "#fca5a5" if is_overdue else "#fcd34d"

    # Oldest-first, deduplicated
    seen: set = set()
    display_dates: list = []
    for d in reversed(all_dates):
        if d not in seen:
            seen.add(d)
            display_dates.append(d)

    if not display_dates:
        orig = task.get("original_date") or "?"
        curr = task.get("due_date") or "?"
        display_dates = list(dict.fromkeys([orig, curr]))

    pills = []
    for i, d in enumerate(display_dates):
        is_last = i == len(display_dates) - 1
        if i == 0:
            pills.append(
                f'<span style="background:#fff;border:1px solid #e5e7eb;border-radius:4px;'
                f'padding:2px 8px;font-size:11px;white-space:nowrap;">'
                f"Start:&nbsp;{_html.escape(str(d))}</span>"
            )
        elif is_last:
            overdue_note = (
                f"&nbsp;&#9888;&nbsp;Overdue&nbsp;+{days_overdue}d"
                if is_overdue
                else ""
            )
            pills.append(
                f'<span style="background:{bg};border:1px solid {accent};border-radius:4px;'
                f'padding:2px 8px;font-size:11px;font-weight:700;color:{label_color};white-space:nowrap;">'
                f"Now:&nbsp;{_html.escape(str(d))}&nbsp;(+{slippage_days}d){overdue_note}</span>"
            )
        else:
            pills.append(
                f'<span style="background:#fff;border:1px solid {pill_border};border-radius:4px;'
                f'padding:2px 8px;font-size:11px;white-space:nowrap;">'
                f"{_html.escape(str(d))}</span>"
            )

    arrow = '<span style="color:#9ca3af;font-size:11px;padding:0 2px;">&rarr;</span>'
    pills_html = arrow.join(pills)
    resc_label = (
        f"Slippage trail &mdash; {n_resc} reschedule{'s' if n_resc != 1 else ''}"
    )

    return (
        f'<div style="margin-top:10px;background:{bg};border-radius:6px;'
        f'padding:10px 12px;border-left:3px solid {accent};">'
        f'<div style="font-size:11px;font-weight:700;color:{label_color};margin-bottom:6px;">'
        f"{resc_label}</div>"
        f'<div style="font-size:12px;color:#374151;display:flex;align-items:center;'
        f'flex-wrap:wrap;gap:4px;">{pills_html}</div>'
        f"</div>"
    )


def _slip_task_card_html(task: dict) -> str:
    """One task card for the slippage digest — card-per-task layout."""
    import html as _html

    slippage_days = task.get("slippage_days", 0)
    is_overdue = bool(task.get("overdue") or task.get("days_overdue", 0) > 0)
    days_overdue = task.get("days_overdue", 0)
    n_resc = max(len(task.get("all_dates") or []) - 1, 0)
    blocking = task.get("blocking") or []
    accent = "#dc2626" if is_overdue else "#f59e0b"
    card_border = "#fecaca" if is_overdue else "#e5e7eb"

    slip_badge = (
        f'<span style="display:inline-block;padding:3px 10px;border-radius:10px;'
        f'background:{accent};color:#fff;font-size:13px;font-weight:700;">+{slippage_days}d</span>'
    )
    extra = ""
    if is_overdue:
        extra += (
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#fee2e2;color:#991b1b;font-size:11px;font-weight:700;margin-left:4px;">'
            f"OVERDUE +{days_overdue}d</span>"
        )
    if blocking:
        extra += (
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;margin-left:4px;">'
            f"Blocking {len(blocking)} task(s)</span>"
        )
    if n_resc > 1:
        extra += (
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;margin-left:4px;">'
            f"Rescheduled {n_resc}&times;</span>"
        )

    chips = []
    if task.get("status"):
        chips.append(_badge(task["status"], "#6b7280", small=True))
    if task.get("priority"):
        chips.append(
            _badge(
                task["priority"],
                _PRIORITY_COLORS.get(task["priority"], "#9ca3af"),
                small=True,
            )
        )
    for team in (task.get("teams") or [])[:2]:
        chips.append(_badge(team, "#0ea5e9", small=True))
    chips_html = (
        (
            "<div style='margin-top:8px;'>"
            + "".join(f'<span style="margin-right:4px;">{c}</span>' for c in chips)
            + "</div>"
        )
        if chips
        else ""
    )

    owner = _html.escape(task.get("owner_name") or "—")
    reviewer = _html.escape(task.get("reviewer_names") or "—")
    people_html = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;font-size:12px;">'
        f"<tr>"
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Owner</td>'
        f'<td style="padding:2px 0;font-weight:600;color:#374151;">{owner}</td>'
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Reviewer</td>'
        f'<td style="padding:2px 0;color:#374151;">{reviewer}</td>'
        f"</tr></table>"
    )

    downstream_html = ""
    if blocking:
        rows = "".join(
            f'<li style="margin:3px 0;">'
            f'<a href="{b["url"]}" style="color:#6d28d9;text-decoration:none;font-weight:600;">'
            f"{_html.escape(b['name'])}</a></li>"
            if b.get("url")
            else f'<li style="margin:3px 0;color:#6d28d9;font-weight:600;">'
            f"{_html.escape(b['name'])}</li>"
            for b in blocking
        )
        downstream_html = (
            f'<div style="margin-top:10px;background:#faf5ff;border-radius:6px;'
            f'padding:10px 14px;border-left:3px solid #8b5cf6;">'
            f'<div style="font-size:11px;font-weight:700;color:#5b21b6;margin-bottom:6px;">'
            f"&#9656; Affected Tasks</div>"
            f'<ul style="margin:0;padding-left:16px;font-size:12px;color:#6d28d9;line-height:1.6;">'
            f"{rows}</ul></div>"
        )

    return (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="margin-bottom:12px;border:1px solid {card_border};'
        f'border-radius:8px;overflow:hidden;">'
        f'<tr><td style="padding:14px 16px;border-left:4px solid {accent};">'
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        f'<td style="vertical-align:top;">'
        f'<a href="{task["url"]}" style="color:#111827;font-size:14px;font-weight:700;'
        f'text-decoration:none;">{_html.escape(task["name"])}</a>'
        f"</td>"
        f'<td style="vertical-align:top;text-align:right;white-space:nowrap;padding-left:12px;">'
        f"{slip_badge}"
        f'<div style="margin-top:4px;">{extra}</div>'
        f"</td></tr></table>"
        f"{chips_html}"
        f"{people_html}"
        f"{_slip_trail_html(task)}"
        f"{downstream_html}"
        f"</td></tr></table>"
    )


def _stalled_section_html(stalled_tasks: list) -> str:
    """Shared 'On Hold / Blocked Tasks' section for slippage and overdue digests."""
    import html as _html

    if not stalled_tasks:
        return ""

    rows = ""
    for t in stalled_tasks:
        status = t.get("status") or "—"
        is_blocked = status.lower() == "blocked"
        badge_bg = "#fee2e2" if is_blocked else "#fef3c7"
        badge_color = "#991b1b" if is_blocked else "#92400e"
        due = t.get("due_date") or "—"
        owner = _html.escape(t.get("owner_name") or "—")
        proj = (t.get("project") or {}).get("name") or "—"
        blocking = t.get("blocking") or []
        affected = ""
        if blocking:
            names = ", ".join(
                f'<a href="{b["url"]}" style="color:#6d28d9;text-decoration:none;">'
                f"{_html.escape(b['name'])}</a>"
                if b.get("url")
                else f'<span style="color:#6d28d9;">{_html.escape(b["name"])}</span>'
                for b in blocking
            )
            affected = (
                f'<div style="margin-top:4px;font-size:11px;color:#6d28d9;">'
                f"&#9656; Blocking: {names}</div>"
            )
        rows += (
            f'<tr style="border-bottom:1px solid #f3f4f6;">'
            f'<td style="padding:9px 12px;vertical-align:top;">'
            f'<a href="{t["url"]}" style="color:#111827;font-size:13px;font-weight:600;'
            f'text-decoration:none;">{_html.escape(t["name"])}</a>'
            f"{affected}</td>"
            f'<td style="padding:9px 12px;vertical-align:top;white-space:nowrap;">'
            f'<span style="display:inline-block;padding:2px 8px;border-radius:8px;'
            f'background:{badge_bg};color:{badge_color};font-size:11px;font-weight:700;">'
            f"{_html.escape(status)}</span></td>"
            f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;vertical-align:top;">'
            f"{_html.escape(proj)}</td>"
            f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;vertical-align:top;'
            f'white-space:nowrap;">{_html.escape(owner)}</td>'
            f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;vertical-align:top;'
            f'white-space:nowrap;">{_html.escape(due)}</td>'
            f"</tr>"
        )

    return (
        f'<div style="margin:24px 0 8px;">'
        f'<div style="font-size:13px;font-weight:700;color:#374151;'
        f"padding:10px 14px;background:#f9fafb;border-radius:8px 8px 0 0;"
        f'border:1px solid #e5e7eb;border-bottom:none;">'
        f"&#9646; On Hold / Blocked Tasks"
        f'<span style="margin-left:8px;font-size:11px;font-weight:600;'
        f'background:#e5e7eb;color:#6b7280;padding:2px 8px;border-radius:10px;">'
        f"{len(stalled_tasks)}</span></div>"
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="border:1px solid #e5e7eb;border-radius:0 0 8px 8px;'
        f'border-top:none;overflow:hidden;">'
        f'<tr style="background:#f9fafb;">'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;color:#9ca3af;'
        f'font-weight:600;">Task</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;color:#9ca3af;'
        f'font-weight:600;">Status</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;color:#9ca3af;'
        f'font-weight:600;">Project</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;color:#9ca3af;'
        f'font-weight:600;">Owner</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;color:#9ca3af;'
        f'font-weight:600;">Due Date</th>'
        f"</tr>"
        f"{rows}"
        f"</table></div>"
    )


def _build_slippage_digest_html(slipped_tasks: list, stalled_tasks: list = None) -> str:
    import html as _html

    today_str = date.today().strftime("%d %b %Y")
    total = len(slipped_tasks)
    max_slip = max((t.get("slippage_days", 0) for t in slipped_tasks), default=0)
    also_overdue = sum(
        1 for t in slipped_tasks if t.get("overdue") or t.get("days_overdue", 0) > 0
    )
    downstream_cnt = sum(1 for t in slipped_tasks if t.get("blocking"))
    proj_names = {
        (t.get("project") or {}).get("name") or "No Project" for t in slipped_tasks
    }
    proj_count = len(proj_names)

    header = _mis_dark_header(
        eyebrow="Due Date Slippage Report",
        eyebrow_color="#9ca3af",
        label_text="Slippage Digest",
        label_color="#dc2626",
        title="Due Date Slippage Report",
        subtitle=f"{today_str}&nbsp;&nbsp;&#183;&nbsp;&nbsp;Tasks with shifted due dates",
    )

    stat_strip = _mis_stat_strip(
        [
            (str(total), "Tasks Slipped", "#dc2626"),
            (f"+{max_slip}d", "Biggest Slip", "#f59e0b"),
            (str(also_overdue), "Also Overdue", "#ef4444"),
            (str(downstream_cnt), "Downstream at Risk", "#8b5cf6"),
            (str(proj_count), "Projects Affected", "#6b7280"),
        ]
    )

    # Group by project (insertion order preserved)
    groups: dict = {}
    for t in slipped_tasks:
        proj = t.get("project") or {}
        pk = proj.get("name") or "No Project"
        pu = proj.get("url") or ""
        if pk not in groups:
            groups[pk] = {"url": pu, "tasks": []}
        groups[pk]["tasks"].append(t)

    body_html = ""
    for proj_name, grp in groups.items():
        proj_link = (
            f'<a href="{grp["url"]}" style="color:#991b1b;text-decoration:none;font-weight:700;">'
            f"{_html.escape(proj_name)}</a>"
            if grp["url"]
            else f'<span style="font-weight:700;">{_html.escape(proj_name)}</span>'
        )
        body_html += (
            f'<div style="display:inline-block;background:#fff5f5;border-radius:4px;'
            f"padding:4px 10px;font-size:11px;font-weight:700;color:#991b1b;"
            f'text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;">'
            f"{proj_link}</div><br>"
        )
        for t in grp["tasks"]:
            body_html += _slip_task_card_html(t)
        body_html += '<div style="height:6px;"></div>'

    body_html += _stalled_section_html(stalled_tasks or [])

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=body_html,
        footer=(
            "Auto-generated slippage digest &mdash; based on Due Date History in Notion. "
            "Downstream impact derived from the <em>Blocking</em> relation."
        ),
        max_width="900px",
    )


# ---------------------------------------------------------------------------
# T6 — Overdue Digest
#
# Sent to:  executives + leaders + all PMs
# Trigger:  --overdue flag; cron daily 9:15 AM IST
#
# task dict fields used:
#   name, url, due_date, days_overdue, overdue (bool),
#   slippage_days, status, priority, teams ([str]),
#   owner_name (str), project ({name,url}|None), db_name
# ---------------------------------------------------------------------------


def _overdue_task_card_html(task: dict) -> str:
    """One card per overdue task for the overdue digest."""
    import html as _html

    days_overdue = task.get("days_overdue", 0)
    slippage_days = task.get("slippage_days", 0)
    due_date = task.get("due_date") or "—"
    blocking = task.get("blocking") or []

    overdue_badge = (
        f'<span style="display:inline-block;padding:3px 10px;border-radius:10px;'
        f'background:#dc2626;color:#fff;font-size:13px;font-weight:700;">'
        f"OVERDUE +{days_overdue}d</span>"
    )

    extra = ""
    if slippage_days > 0:
        extra += (
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;margin-left:4px;">'
            f"Also slipped +{slippage_days}d</span>"
        )
    if blocking:
        extra += (
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;margin-left:4px;">'
            f"Blocking {len(blocking)} task(s)</span>"
        )

    chips = []
    if task.get("status"):
        chips.append(_badge(task["status"], "#6b7280", small=True))
    if task.get("priority"):
        chips.append(
            _badge(
                task["priority"],
                _PRIORITY_COLORS.get(task["priority"], "#9ca3af"),
                small=True,
            )
        )
    for team in (task.get("teams") or [])[:2]:
        chips.append(_badge(team, "#0ea5e9", small=True))
    chips_html = (
        (
            "<div style='margin-top:8px;'>"
            + "".join(f'<span style="margin-right:4px;">{c}</span>' for c in chips)
            + "</div>"
        )
        if chips
        else ""
    )

    owner = _html.escape(task.get("owner_name") or "—")
    people_html = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;font-size:12px;">'
        f"<tr>"
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Owner</td>'
        f'<td style="padding:2px 0;font-weight:600;color:#374151;">{owner}</td>'
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Was due</td>'
        f'<td style="padding:2px 0;font-weight:600;color:#dc2626;">{_html.escape(due_date)}</td>'
        f"</tr></table>"
    )

    downstream_html = ""
    if blocking:
        rows = "".join(
            f'<li style="margin:3px 0;">'
            f'<a href="{b["url"]}" style="color:#6d28d9;text-decoration:none;font-weight:600;">'
            f"{_html.escape(b['name'])}</a></li>"
            if b.get("url")
            else f'<li style="margin:3px 0;color:#6d28d9;font-weight:600;">'
            f"{_html.escape(b['name'])}</li>"
            for b in blocking
        )
        downstream_html = (
            f'<div style="margin-top:10px;background:#faf5ff;border-radius:6px;'
            f'padding:10px 14px;border-left:3px solid #8b5cf6;">'
            f'<div style="font-size:11px;font-weight:700;color:#5b21b6;margin-bottom:6px;">'
            f"&#9656; Affected Tasks</div>"
            f'<ul style="margin:0;padding-left:16px;font-size:12px;color:#6d28d9;line-height:1.6;">'
            f"{rows}</ul></div>"
        )

    return (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="margin-bottom:12px;border:1px solid #fecaca;'
        f'border-radius:8px;overflow:hidden;">'
        f'<tr><td style="padding:14px 16px;border-left:4px solid #dc2626;">'
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        f'<td style="vertical-align:top;">'
        f'<a href="{task["url"]}" style="color:#111827;font-size:14px;font-weight:700;'
        f'text-decoration:none;">{_html.escape(task["name"])}</a>'
        f"</td>"
        f'<td style="vertical-align:top;text-align:right;white-space:nowrap;padding-left:12px;">'
        f"{overdue_badge}"
        f'<div style="margin-top:4px;">{extra}</div>'
        f"</td></tr></table>"
        f"{chips_html}"
        f"{people_html}"
        f"{downstream_html}"
        f"</td></tr></table>"
    )


def _build_overdue_digest_html(overdue_tasks: list, stalled_tasks: list = None) -> str:
    import html as _html

    today_str = date.today().strftime("%d %b %Y")
    total = len(overdue_tasks)
    worst = max((t.get("days_overdue", 0) for t in overdue_tasks), default=0)
    high_pri = sum(
        1
        for t in overdue_tasks
        if (t.get("priority") or "").lower() in ("high", "critical")
    )
    downstream_cnt = sum(1 for t in overdue_tasks if t.get("blocking"))
    proj_names = {
        (t.get("project") or {}).get("name") or "No Project" for t in overdue_tasks
    }
    proj_count = len(proj_names)

    header = _mis_dark_header(
        eyebrow="Overdue Tasks Report",
        eyebrow_color="#9ca3af",
        label_text="Overdue Alert",
        label_color="#dc2626",
        title="Overdue Tasks Report",
        subtitle=f"{today_str}&nbsp;&nbsp;&#183;&nbsp;&nbsp;Tasks past their due date",
    )

    stat_strip = _mis_stat_strip(
        [
            (str(total), "Tasks Overdue", "#dc2626"),
            (f"+{worst}d", "Most Overdue", "#ef4444"),
            (str(high_pri), "High Priority", "#f59e0b"),
            (str(downstream_cnt), "Downstream at Risk", "#8b5cf6"),
            (str(proj_count), "Projects Affected", "#6b7280"),
        ]
    )

    groups: dict = {}
    for t in overdue_tasks:
        proj = t.get("project") or {}
        pk = proj.get("name") or "No Project"
        pu = proj.get("url") or ""
        if pk not in groups:
            groups[pk] = {"url": pu, "tasks": []}
        groups[pk]["tasks"].append(t)

    body_html = ""
    for proj_name, grp in groups.items():
        proj_link = (
            f'<a href="{grp["url"]}" style="color:#991b1b;text-decoration:none;font-weight:700;">'
            f"{_html.escape(proj_name)}</a>"
            if grp["url"]
            else f'<span style="font-weight:700;">{_html.escape(proj_name)}</span>'
        )
        body_html += (
            f'<div style="display:inline-block;background:#fff5f5;border-radius:4px;'
            f"padding:4px 10px;font-size:11px;font-weight:700;color:#991b1b;"
            f'text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;">'
            f"{proj_link}</div><br>"
        )
        for t in grp["tasks"]:
            body_html += _overdue_task_card_html(t)
        body_html += '<div style="height:6px;"></div>'

    body_html += _stalled_section_html(stalled_tasks or [])

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=body_html,
        footer="Auto-generated overdue digest &mdash; based on the Overdue formula field in Notion.",
        max_width="900px",
    )


# ---------------------------------------------------------------------------
# T3+T6 Combined — Slippage & Overdue Digest
#
# Replaces the separate T3 (slippage) and T6 (overdue) digests.
# On Hold tasks are excluded from the main body and shown only in a
# dedicated bottom section. Blocked tasks remain in the main body.
# Tasks that appear in both slipped_tasks and overdue_tasks are deduplicated.
#
# Parameters:
#   slipped_tasks  — from detect_slippage() (slippage_days >= min_days)
#   overdue_tasks  — from detect_overdue()  (days_overdue  >= min_days)
#   stalled_tasks  — On Hold / Blocked tasks (collected by either detector)
# ---------------------------------------------------------------------------


def _is_on_hold(task: dict) -> bool:
    return (task.get("status") or "").lower() == "on hold"


def _trim_number(n: float) -> str:
    n = round(n, 1)
    return str(int(n)) if n == int(n) else str(n)


def _fmt_days(days) -> str:
    """
    Format a day count, collapsing to weeks past 8 days and to months past
    4 weeks so the digest doesn't show unwieldy day/week counts.
    """
    days = days or 0
    if days <= 8:
        return f"{days}d"
    weeks = days / 7
    if weeks <= 4:
        return f"{_trim_number(weeks)}w"
    return f"{_trim_number(days / 30)}mo"


def _updated_ago_label(days) -> str:
    """'Updated today' / 'Updated {n} ago', using _fmt_days for the count."""
    if not days:
        return "Updated today"
    return f"Updated {_fmt_days(days)} ago"


def _combined_project_badge_html(name: str, url: str) -> str:
    """Project name badge shown above a group of tasks in the combined digest."""
    import html as _html

    label = (
        f'<a href="{url}" style="color:#991b1b;text-decoration:none;font-weight:700;">'
        f"{_html.escape(name)}</a>"
        if url
        else f'<span style="font-weight:700;">{_html.escape(name)}</span>'
    )
    return (
        f'<div style="display:inline-block;background:#fff5f5;border-radius:4px;'
        f"padding:4px 10px;font-size:11px;font-weight:700;color:#991b1b;"
        f'text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;">'
        f"{label}</div><br>"
    )


def _combined_slip_trail_html(task: dict) -> str:
    """
    Slippage trail box for combined cards.
    Overdue-only tasks (slippage_days == 0) get a simplified "Was due → Overdue" display.
    """
    import html as _html

    all_dates = task.get("all_dates") or []
    slippage_days = task.get("slippage_days", 0)
    days_overdue = task.get("days_overdue", 0)
    is_overdue = bool(task.get("overdue") or days_overdue > 0)
    n_resc = max(len(all_dates) - 1, 0)

    accent = "#dc2626" if is_overdue else "#f59e0b"
    bg = "#fff1f2" if is_overdue else "#fffbeb"
    label_color = "#991b1b" if is_overdue else "#92400e"
    pill_border = "#fca5a5" if is_overdue else "#fcd34d"

    if slippage_days == 0 and is_overdue:
        due_date = _html.escape(str(task.get("due_date") or "?"))
        return (
            f'<div style="margin-top:10px;background:{bg};border-radius:6px;'
            f'padding:10px 12px;border-left:3px solid {accent};">'
            f'<div style="font-size:11px;font-weight:700;color:{label_color};margin-bottom:6px;">'
            f"No reschedule &mdash; past original due date</div>"
            f'<div style="font-size:12px;color:#374151;display:flex;align-items:center;'
            f'flex-wrap:wrap;gap:4px;">'
            f'<span style="background:#fff;border:1px solid #e5e7eb;border-radius:4px;'
            f'padding:2px 8px;font-size:11px;white-space:nowrap;">Was due:&nbsp;{due_date}</span>'
            f'<span style="color:#9ca3af;font-size:11px;padding:0 2px;">&rarr;</span>'
            f'<span style="background:{bg};border:1px solid {accent};border-radius:4px;'
            f'padding:2px 8px;font-size:11px;font-weight:700;color:{label_color};white-space:nowrap;">'
            f"Overdue&nbsp;+{_fmt_days(days_overdue)}&nbsp;&#9888;</span>"
            f"</div></div>"
        )

    seen: set = set()
    display_dates: list = []
    for d in reversed(all_dates):
        if d not in seen:
            seen.add(d)
            display_dates.append(d)

    if not display_dates:
        orig = task.get("original_date") or "?"
        curr = task.get("due_date") or "?"
        display_dates = list(dict.fromkeys([orig, curr]))

    pills = []
    for i, d in enumerate(display_dates):
        is_last = i == len(display_dates) - 1
        if i == 0:
            pills.append(
                f'<span style="background:#fff;border:1px solid #e5e7eb;border-radius:4px;'
                f'padding:2px 8px;font-size:11px;white-space:nowrap;">'
                f"Start:&nbsp;{_html.escape(str(d))}</span>"
            )
        elif is_last:
            overdue_note = (
                f"&nbsp;&#9888;&nbsp;Overdue&nbsp;+{_fmt_days(days_overdue)}"
                if is_overdue
                else ""
            )
            pills.append(
                f'<span style="background:{bg};border:1px solid {accent};border-radius:4px;'
                f'padding:2px 8px;font-size:11px;font-weight:700;color:{label_color};white-space:nowrap;">'
                f"Now:&nbsp;{_html.escape(str(d))}&nbsp;(+{_fmt_days(slippage_days)}){overdue_note}</span>"
            )
        else:
            pills.append(
                f'<span style="background:#fff;border:1px solid {pill_border};border-radius:4px;'
                f'padding:2px 8px;font-size:11px;white-space:nowrap;">'
                f"{_html.escape(str(d))}</span>"
            )

    arrow = '<span style="color:#9ca3af;font-size:11px;padding:0 2px;">&rarr;</span>'
    pills_html = arrow.join(pills)
    resc_label = (
        f"Slippage trail &mdash; {n_resc} reschedule{'s' if n_resc != 1 else ''}"
    )

    return (
        f'<div style="margin-top:10px;background:{bg};border-radius:6px;'
        f'padding:10px 12px;border-left:3px solid {accent};">'
        f'<div style="font-size:11px;font-weight:700;color:{label_color};margin-bottom:6px;">'
        f"{resc_label}</div>"
        f'<div style="font-size:12px;color:#374151;display:flex;align-items:center;'
        f'flex-wrap:wrap;gap:4px;">{pills_html}</div>'
        f"</div>"
    )


def _combined_task_card_html(task: dict) -> str:
    """Unified task card for both slipped and overdue-only tasks."""
    import html as _html

    slippage_days = task.get("slippage_days", 0)
    is_overdue = bool(task.get("overdue") or task.get("days_overdue", 0) > 0)
    days_overdue = task.get("days_overdue", 0)
    n_resc = max(len(task.get("all_dates") or []) - 1, 0)
    blocking = task.get("blocking") or []

    accent = "#dc2626" if is_overdue else "#f59e0b"
    card_border = "#fecaca" if is_overdue else "#e5e7eb"

    main_badge = (
        (
            f'<span style="display:inline-block;padding:3px 10px;border-radius:10px;'
            f'background:{accent};color:#fff;font-size:13px;font-weight:700;">+{_fmt_days(slippage_days)}</span>'
        )
        if slippage_days > 0
        else ""
    )

    extra_parts = []
    if is_overdue:
        extra_parts.append(
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#fee2e2;color:#991b1b;font-size:11px;font-weight:700;">'
            f"OVERDUE +{_fmt_days(days_overdue)}</span>"
        )
    if blocking:
        extra_parts.append(
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#ede9fe;color:#6d28d9;font-size:11px;font-weight:700;">'
            f"Blocking {len(blocking)} task(s)</span>"
        )
    if n_resc > 1:
        extra_parts.append(
            f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;'
            f'background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;">'
            f"Rescheduled {n_resc}&times;</span>"
        )
    extra_html = (
        (
            '<div style="margin-top:4px;">'
            + "".join(
                f'<span style="margin-right:4px;">{p}</span>' for p in extra_parts
            )
            + "</div>"
        )
        if extra_parts
        else ""
    )

    chips = []
    if task.get("status"):
        chips.append(_badge(task["status"], "#6b7280", small=True))
    if task.get("priority"):
        chips.append(
            _badge(
                task["priority"],
                _PRIORITY_COLORS.get(task["priority"], "#9ca3af"),
                small=True,
            )
        )
    for team in (task.get("teams") or [])[:2]:
        chips.append(_badge(team, "#0ea5e9", small=True))
    if task.get("days_since_update") is not None:
        chips.append(
            _badge(_updated_ago_label(task["days_since_update"]), "#6366f1", small=True)
        )
    chips_html = (
        (
            "<div style='margin-top:8px;'>"
            + "".join(f'<span style="margin-right:4px;">{c}</span>' for c in chips)
            + "</div>"
        )
        if chips
        else ""
    )

    owner = _html.escape(task.get("owner_name") or "—")
    reviewer = _html.escape(task.get("reviewer_names") or "—")
    people_html = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;font-size:12px;">'
        f"<tr>"
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Owner</td>'
        f'<td style="padding:2px 0;font-weight:600;color:#374151;">{owner}</td>'
        f'<td style="width:80px;color:#9ca3af;padding:2px 0;">Reviewer</td>'
        f'<td style="padding:2px 0;color:#374151;">{reviewer}</td>'
        f"</tr></table>"
    )

    downstream_html = ""
    if blocking:
        rows = "".join(
            f'<li style="margin:3px 0;">'
            f'<a href="{b["url"]}" style="color:#6d28d9;text-decoration:none;font-weight:600;">'
            f"{_html.escape(b['name'])}</a></li>"
            if b.get("url")
            else f'<li style="margin:3px 0;color:#6d28d9;font-weight:600;">'
            f"{_html.escape(b['name'])}</li>"
            for b in blocking
        )
        downstream_html = (
            f'<div style="margin-top:10px;background:#faf5ff;border-radius:6px;'
            f'padding:10px 14px;border-left:3px solid #8b5cf6;">'
            f'<div style="font-size:11px;font-weight:700;color:#5b21b6;margin-bottom:6px;">'
            f"&#9656; Affected Tasks</div>"
            f'<ul style="margin:0;padding-left:16px;font-size:12px;color:#6d28d9;line-height:1.6;">'
            f"{rows}</ul></div>"
        )

    return (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="margin-bottom:12px;border:1px solid {card_border};'
        f'border-radius:8px;overflow:hidden;">'
        f'<tr><td style="padding:14px 16px;border-left:4px solid {accent};">'
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        f'<td style="vertical-align:top;">'
        f'<a href="{task["url"]}" style="color:#111827;font-size:14px;font-weight:700;'
        f'text-decoration:none;">{_html.escape(task["name"])}</a>'
        f"</td>"
        f'<td style="vertical-align:top;text-align:right;white-space:nowrap;padding-left:12px;">'
        f"{main_badge}{extra_html}"
        f"</td>"
        f"</tr></table>"
        f"{chips_html}"
        f"{people_html}"
        f"{_combined_slip_trail_html(task)}"
        f"{downstream_html}"
        f"</td></tr></table>"
    )


def _stalled_project_header_row(name: str, url: str) -> str:
    """Colspan sub-header row grouping the stalled table by project."""
    import html as _html

    label = (
        f'<a href="{url}" style="color:#991b1b;text-decoration:none;font-weight:700;">'
        f"{_html.escape(name)}</a>"
        if url
        else f'<span style="font-weight:700;">{_html.escape(name)}</span>'
    )
    return (
        f"<tr>"
        f'<td colspan="6" style="padding:8px 12px;background:#fff5f5;'
        f'border-left:4px solid #991b1b;border-bottom:1px solid #e5e7eb;">'
        f'<span style="font-size:11px;font-weight:700;color:#991b1b;'
        f'text-transform:uppercase;letter-spacing:.04em;">{label}</span>'
        f"</td>"
        f"</tr>"
    )


def _stalled_section_v2_html(stalled_tasks: list) -> str:
    """
    On Hold / Blocked section with an extra Slippage/Overdue column.
    Rows that have active slippage or overdue are highlighted with a cyan accent (#0891b2).
    Grouped by project, matching the main report; the Project column is dropped
    in favor of a project header above each group's rows.
    """
    import html as _html

    if not stalled_tasks:
        return ""

    groups: dict = {}
    for t in stalled_tasks:
        proj = t.get("project") or {}
        pk = proj.get("name") or "No Project"
        pu = proj.get("url") or ""
        if pk not in groups:
            groups[pk] = {"url": pu, "tasks": []}
        groups[pk]["tasks"].append(t)

    rows = ""
    for proj_name, grp in groups.items():
        rows += _stalled_project_header_row(proj_name, grp["url"])
        for t in grp["tasks"]:
            status = t.get("status") or "—"
            is_blocked = status.lower() == "blocked"
            badge_bg = "#fee2e2" if is_blocked else "#fef3c7"
            badge_color = "#991b1b" if is_blocked else "#92400e"
            due = t.get("due_date") or "—"
            owner = _html.escape(t.get("owner_name") or "—")
            blocking = t.get("blocking") or []
            slip_days = t.get("slippage_days", 0)
            overdue_days = t.get("days_overdue", 0)
            is_overdue = bool(t.get("overdue") or overdue_days > 0)
            has_issue = slip_days > 0 or is_overdue

            row_style = "background:#f0f9ff;" if has_issue else ""
            cell_style = "border-left:3px solid #0891b2;" if has_issue else ""

            if has_issue:
                badges = []
                if slip_days > 0:
                    badges.append(
                        f'<span style="display:inline-block;padding:1px 7px;border-radius:6px;'
                        f"background:#fff0f0;border:1px solid #fca5a5;color:#b91c1c;"
                        f'font-size:10px;font-weight:700;white-space:nowrap;">+{_fmt_days(slip_days)} slip</span>'
                    )
                if is_overdue:
                    badges.append(
                        f'<span style="display:inline-block;padding:1px 7px;border-radius:6px;'
                        f"background:#fff0f0;border:1px solid #fca5a5;color:#b91c1c;"
                        f'font-size:10px;font-weight:700;white-space:nowrap;">OVD +{_fmt_days(overdue_days)}</span>'
                    )
                slip_ov_cell = (
                    '<div style="display:flex;flex-direction:column;gap:3px;">'
                    + "".join(badges)
                    + "</div>"
                )
            else:
                slip_ov_cell = '<span style="color:#9ca3af;font-size:11px;">—</span>'

            days_since_update = t.get("days_since_update")
            last_updated_cell = (
                _fmt_days(days_since_update)
                if days_since_update is not None
                else "—"
            )

            affected = ""
            if blocking:
                names = ", ".join(
                    f'<a href="{b["url"]}" style="color:#6d28d9;text-decoration:none;">'
                    f"{_html.escape(b['name'])}</a>"
                    if b.get("url")
                    else f'<span style="color:#6d28d9;">{_html.escape(b["name"])}</span>'
                    for b in blocking
                )
                affected = (
                    f'<div style="margin-top:4px;font-size:11px;color:#6d28d9;">'
                    f"&#9656; Blocking: {names}</div>"
                )

            rows += (
                f'<tr style="border-bottom:1px solid #f3f4f6;{row_style}">'
                f'<td style="padding:9px 12px;vertical-align:top;{cell_style}">'
                f'<a href="{t["url"]}" style="color:#111827;font-size:13px;font-weight:600;'
                f'text-decoration:none;">{_html.escape(t["name"])}</a>'
                f"{affected}</td>"
                f'<td style="padding:9px 12px;vertical-align:top;white-space:nowrap;">'
                f'<span style="display:inline-block;padding:2px 8px;border-radius:8px;'
                f'background:{badge_bg};color:{badge_color};font-size:11px;font-weight:700;">'
                f"{_html.escape(status)}</span></td>"
                f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;vertical-align:top;">'
                f"{_html.escape(owner)}</td>"
                f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;'
                f'vertical-align:top;white-space:nowrap;">{_html.escape(due)}</td>'
                f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;'
                f'vertical-align:top;white-space:nowrap;">{last_updated_cell}</td>'
                f'<td style="padding:9px 12px;vertical-align:top;">{slip_ov_cell}</td>'
                f"</tr>"
            )

    has_any_issue = any(
        t.get("slippage_days", 0) > 0
        or t.get("overdue")
        or t.get("days_overdue", 0) > 0
        for t in stalled_tasks
    )
    note = (
        (
            f'<span style="margin-left:12px;font-size:11px;color:#0891b2;font-weight:500;">'
            f"&#9432; Highlighted rows have active slippage or overdue</span>"
        )
        if has_any_issue
        else ""
    )

    return (
        f'<div style="margin:24px 0 8px;">'
        f'<div style="font-size:13px;font-weight:700;color:#374151;'
        f"padding:10px 14px;background:#f9fafb;border-radius:8px 8px 0 0;"
        f'border:1px solid #e5e7eb;border-bottom:none;">'
        f"&#9646; On Hold / Blocked Tasks"
        f'<span style="margin-left:8px;font-size:11px;font-weight:600;'
        f'background:#e5e7eb;color:#6b7280;padding:2px 8px;border-radius:10px;">'
        f"{len(stalled_tasks)}</span>"
        f"{note}</div>"
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="border:1px solid #e5e7eb;border-radius:0 0 8px 8px;'
        f'border-top:none;overflow:hidden;">'
        f'<tr style="background:#f9fafb;">'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Task</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Status</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Owner</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Due Date</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Last Updated</th>'
        f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
        f'color:#9ca3af;font-weight:600;">Slippage / Overdue</th>'
        f"</tr>"
        f"{rows}"
        f"</table></div>"
    )


def _build_combined_digest_html(
    slipped_tasks: list,
    overdue_tasks: list,
    stalled_tasks: list = None,
) -> str:
    """
    Combined Slippage & Overdue digest (replaces separate T3 + T6).

    - On Hold tasks are excluded from the main body; shown in the bottom section only.
    - Blocked tasks remain in the main body (active downstream impact).
    - Tasks in both slipped_tasks and overdue_tasks are deduplicated (slipped takes precedence).
    - Overdue-only tasks (slippage_days == 0) use a simplified "Was due → Overdue" trail.
    - stalled_tasks: On Hold / Blocked collected by detect_slippage() or detect_overdue().
    """
    import html as _html

    today_str = date.today().strftime("%d %b %Y")
    stalled = stalled_tasks or []

    # Exclude On Hold from main body; Blocked stays in main
    active_slipped = [t for t in slipped_tasks if not _is_on_hold(t)]
    active_overdue = [t for t in overdue_tasks if not _is_on_hold(t)]

    # Deduplicate: slipped tasks take the full card; pure-overdue (no slip) get simpler trail
    slipped_ids = {t["id"] for t in active_slipped}
    overdue_only = [t for t in active_overdue if t["id"] not in slipped_ids]
    all_active = active_slipped + overdue_only

    # Stats
    total_slipped = len(active_slipped)
    total_overdue = sum(
        1 for t in all_active if t.get("overdue") or t.get("days_overdue", 0) > 0
    )
    max_slip = max((t.get("slippage_days", 0) for t in active_slipped), default=0)
    downstream_cnt = sum(1 for t in all_active if t.get("blocking"))
    proj_names = {
        (t.get("project") or {}).get("name") or "No Project" for t in all_active
    }
    proj_count = len(proj_names)

    header = _mis_dark_header(
        eyebrow="Schedule Status Report",
        eyebrow_color="#9ca3af",
        label_text="Project Timeline Digest",
        label_color="#dc2626",
        title="Slippage & Overdue Report",
        subtitle=f"{today_str}&nbsp;&nbsp;&#183;&nbsp;&nbsp;Active slippage, overdue and dependency impacts",
    )

    stat_strip = _mis_stat_strip(
        [
            (str(total_slipped), "Slipped Tasks", "#f59e0b"),
            (str(total_overdue), "Overdue Tasks", "#dc2626"),
            (f"+{_fmt_days(max_slip)}", "Biggest Slip", "#ef4444"),
            (str(downstream_cnt), "Downstream at Risk", "#8b5cf6"),
            (str(proj_count), "Projects Affected", "#6b7280"),
        ]
    )

    # Group by project
    groups: dict = {}
    for t in all_active:
        proj = t.get("project") or {}
        pk = proj.get("name") or "No Project"
        pu = proj.get("url") or ""
        if pk not in groups:
            groups[pk] = {"url": pu, "tasks": []}
        groups[pk]["tasks"].append(t)

    body_html = ""
    for proj_name, grp in groups.items():
        proj_link = (
            f'<a href="{grp["url"]}" style="color:#991b1b;text-decoration:none;font-weight:700;">'
            f"{_html.escape(proj_name)}</a>"
            if grp["url"]
            else f'<span style="font-weight:700;">{_html.escape(proj_name)}</span>'
        )
        body_html += (
            f'<div style="display:inline-block;background:#fff5f5;border-radius:4px;'
            f"padding:4px 10px;font-size:11px;font-weight:700;color:#991b1b;"
            f'text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;">'
            f"{proj_link}</div><br>"
        )
        for t in grp["tasks"]:
            body_html += _combined_task_card_html(t)
        body_html += '<div style="height:6px;"></div>'

    body_html += _stalled_section_v2_html(stalled)

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=body_html,
        footer="Auto-generated schedule digest by MIS.",
        max_width="900px",
    )


# ---------------------------------------------------------------------------
# T4 — Project / Task Started  (detailed kickoff briefing)
#
# Sent to:  Project Owner  (CC: Assignee)
# Trigger:  Status transitions to "In progress" (first time only)
#
# task dict must include:
#   name, url, due_date, status, priority, teams (list), description,
#   project ({name,url}|None), parent_task ({name,url}|None), db_name,
#   assignees ([{name, email}]),
#   team_members ([{name}])     — everyone on the task/project
# sub_tasks: [{name, url, status, due_date, owner_name}]
# ---------------------------------------------------------------------------


def _build_project_started_html(task: dict, sub_tasks: list) -> str:
    accent = _MIS_ACCENT["started"]
    today_str = date.today().strftime("%d %b %Y")
    due = task.get("due_date") or "—"
    priority = task.get("priority") or "—"
    pri_color = _PRIORITY_COLORS.get(priority, "#9ca3af")
    db_name = task.get("db_name", "Notion")

    assignees = task.get("assignees") or []
    assignee_str = ", ".join(a.get("name", "") for a in assignees) or "—"
    team_members = task.get("team_members") or []

    desc_raw = task.get("description") or ""
    desc_html = (
        (
            f'<div style="font-size:13px;color:#374151;line-height:1.65;'
            f"padding:12px 14px;background:#f9fafb;border-radius:6px;"
            f'border-left:3px solid {accent};">{desc_raw}</div>'
        )
        if desc_raw
        else (
            f'<span style="font-size:12px;color:#9ca3af;">No description provided.</span>'
        )
    )

    team_pills = (
        "".join(
            f'<span style="display:inline-block;padding:4px 10px;border-radius:12px;'
            f'background:#f3f4f6;color:#374151;font-size:12px;margin:3px 4px 3px 0;">'
            f"&#128100; {m.get('name', '')}</span>"
            for m in team_members
        )
        if team_members
        else '<span style="font-size:12px;color:#9ca3af;">&mdash;</span>'
    )

    sub_rows = ""
    for st in sub_tasks:
        st_done = (st.get("status") or "").lower() in ("done", "complete", "completed")
        st_color = "#10b981" if st_done else "#6b7280"
        sub_rows += (
            f'<tr style="border-bottom:1px solid #f3f4f6;">'
            f'<td style="padding:9px 12px;vertical-align:middle;">'
            f'<a href="{st["url"]}" style="color:#111827;font-size:13px;font-weight:600;'
            f'text-decoration:none;">{st["name"]}</a></td>'
            f'<td style="padding:9px 12px;vertical-align:middle;">'
            f"{_badge(st.get('status') or '&mdash;', st_color, small=True)}</td>"
            f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;'
            f'white-space:nowrap;vertical-align:middle;">'
            f"{st.get('due_date') or '&mdash;'}</td>"
            f'<td style="padding:9px 12px;font-size:12px;color:#6b7280;'
            f'white-space:nowrap;vertical-align:middle;">'
            f"{st.get('owner_name') or '&mdash;'}</td>"
            f"</tr>"
        )

    sub_section = ""
    if sub_tasks:
        sub_table = (
            f'<table width="100%" cellpadding="0" cellspacing="0"'
            f' style="border-collapse:collapse;">'
            f'<tr style="background:#f9fafb;">'
            f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
            f'color:#6b7280;font-weight:600;">Sub-task</th>'
            f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
            f'color:#6b7280;font-weight:600;width:100px;">Status</th>'
            f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
            f'color:#6b7280;font-weight:600;width:90px;">Due</th>'
            f'<th style="padding:8px 12px;text-align:left;font-size:11px;'
            f'color:#6b7280;font-weight:600;width:130px;">Owner</th>'
            f"</tr>{sub_rows}</table>"
        )
        sub_section = _mis_section_card(
            title="Sub-tasks",
            badge_val=str(len(sub_tasks)),
            accent=accent,
            inner_html=sub_table,
        )

    body = (
        _mis_section_card(
            title="Description &amp; Scope",
            badge_val="",
            accent=accent,
            inner_html=_mis_crumb(task) + desc_html,
        )
        + _mis_section_card(
            title="Team",
            badge_val=str(len(team_members)) if team_members else "—",
            accent=accent,
            inner_html=team_pills,
        )
        + sub_section
        + f'<div style="text-align:center;padding:6px 0 14px;">'
        f"{_mis_cta('Open Task in Notion', task['url'], accent)}</div>"
    )

    header = _mis_dark_header(
        eyebrow="Project Kickoff",
        eyebrow_color="#9ca3af",
        label_text="Now In Progress",
        label_color=accent,
        title=task["name"],
        subtitle=(
            f"{db_name}&nbsp;&nbsp;&#183;&nbsp;&nbsp;{today_str}"
            f"&nbsp;&nbsp;&#183;&nbsp;&nbsp;Assignee: {assignee_str}"
        ),
    )

    stat_strip = _mis_stat_strip(
        [
            (due, "Due Date", accent),
            (priority, "Priority", pri_color if priority != "—" else "#9ca3af"),
            (task.get("status") or "In Progress", "Status", accent),
            (str(len(sub_tasks)), "Sub-tasks", "#6b7280"),
        ]
    )

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=body,
        footer="Sent via MIS &mdash; Project kickoff notification for Project Owners",
        max_width="720px",
    )


# ---------------------------------------------------------------------------
# T5 — Due-in-2-Days Reminder  (short, urgent)
#
# Sent to:  Project Owner
# Trigger:  Daily scan — tasks where due_date == today+2 AND status != Done
#
# task dict must include:
#   name, url, due_date, status, priority,
#   project ({name,url}|None), parent_task ({name,url}|None),
#   db_name, assignees ([{name, email}])
# days_until_due: 1 or 2  (adjusts header label; template layout is the same)
# ---------------------------------------------------------------------------


def _build_due_soon_html(task: dict, days_until_due: int = 2) -> str:
    accent = _MIS_ACCENT["due_soon"]
    today_str = date.today().strftime("%d %b %Y")
    due = task.get("due_date") or "—"
    priority = task.get("priority") or "—"
    pri_color = _PRIORITY_COLORS.get(priority, "#9ca3af")
    db_name = task.get("db_name", "Notion")

    assignees = task.get("assignees") or []
    assignee_str = ", ".join(a.get("name", "") for a in assignees) or "—"
    day_label = (
        "Due Tomorrow" if days_until_due == 1 else f"Due in {days_until_due} Days"
    )

    task_block = (
        f'<table width="100%" cellpadding="0" cellspacing="0"'
        f' style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;margin-bottom:16px;">'
        f'<tr><td style="padding:16px 18px;border-left:4px solid {accent};">'
        f"{_mis_crumb(task)}"
        f'<a href="{task["url"]}" style="color:#111827;font-size:15px;font-weight:700;'
        f'text-decoration:none;">{task["name"]}</a>'
        f"{_mis_chips(task)}"
        f"<div>{_mis_cta('Review &amp; Action', task['url'], accent)}</div>"
        f"</td></tr></table>"
    )

    assignee_note = (
        f'<p style="margin:0 0 14px;font-size:12px;color:#6b7280;text-align:center;">'
        f'Assignee: <strong style="color:#374151;">{assignee_str}</strong>'
        f"&nbsp;&nbsp;&#183;&nbsp;&nbsp;"
        f'Database: <strong style="color:#374151;">{db_name}</strong>'
        f"</p>"
    )

    header = _mis_dark_header(
        eyebrow=day_label,
        eyebrow_color=accent,
        label_text=day_label,
        label_color=accent,
        title=task["name"],
        subtitle=f"{db_name}&nbsp;&nbsp;&#183;&nbsp;&nbsp;{today_str}",
    )

    stat_strip = _mis_stat_strip(
        [
            (due, "Due Date", accent),
            (priority, "Priority", pri_color if priority != "—" else "#9ca3af"),
            (assignee_str, "Assignee", "#111827"),
        ]
    )

    return _mis_email_shell(
        header=header,
        stat_strip=stat_strip,
        body=task_block + assignee_note,
        footer="Auto-reminder sent via MIS &mdash; 2 days before due date",
        max_width="560px",
    )
