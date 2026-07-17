"""
vt_mailer.py — Supervisory Digest email template + Graph API sender.

Rebuilds the approved preview_cards/vt_supervisory_digest_v4.html design as
nested-table, inline-style HTML — the flexbox/grid/backdrop-filter/3D-tilt
CSS in that preview file doesn't render in real mail clients (Outlook in
particular), so this is a from-scratch, email-safe reconstruction of the
same colors, sections, and copy.

Dry run is the default and only becomes a real send when explicitly told
to: VT_MAILER_DRY_RUN defaults to "1" (safe), and writes the rendered HTML
to VT_MAILER_DRY_RUN_DIR (default: violation_tester/test_emails/) instead of
calling Graph. Mirrors the parent project's mailer.py dry-run pattern
exactly, but fully isolated — own env vars, own output folder, no imports
from or shared state with mailer.py.
"""
import base64
import logging
import os
import re
from datetime import date, datetime
from pathlib import Path

import msal
import requests

log = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent
_GRAPH_SEND_MAIL = "https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
_SCOPES = ["https://graph.microsoft.com/.default"]

# (ribbon/accent color, table-header tint, tint text color)
_STATUS_COLORS = {
    "Data Error": ("#5b6472", "#eef1f4", "#475569"),
    "Violated":   ("#c0392b", "#fceceb", "#96281b"),
    "Blocked":    ("#4c51bf", "#ececfc", "#363a91"),
    "At Risk":    ("#d68910", "#fdf1de", "#96650c"),
    "On Track":   ("#1f9d78", "#e8f7f2", "#146856"),
}
_RIBBON_TEXT = {
    "Data Error": "Data fixes required",
    "Violated": "Immediate action required",
    "Blocked": "Dependencies awaiting resolution",
    "At Risk": "Needs attention",
    "On Track": "Stable operations",
}
_DETAIL_LABEL = {
    "Data Error": "Issue",
    "Violated": "Slip",
    "Blocked": "Waiting on",
    "At Risk": "Slip",
    "On Track": "Margin",
}
_GROUP_ORDER = ["Data Error", "Violated", "Blocked", "At Risk", "On Track"]


# ---------------------------------------------------------------------------
# Dry run / Graph send — mirrors mailer.py's pattern, fully isolated
# ---------------------------------------------------------------------------

def _dry_run_enabled() -> bool:
    return os.environ.get("VT_MAILER_DRY_RUN", "1").lower() in ("1", "true", "yes")


def _dry_run_dir() -> Path:
    return Path(os.environ.get("VT_MAILER_DRY_RUN_DIR", str(BASE_DIR / "test_emails")))


_DRY_RUN_COUNTER = {"n": 0}


def _safe_slug(s: str, max_len: int = 60) -> str:
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", (s or "digest").strip())
    return s[:max_len].strip("_") or "digest"


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
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description', result.get('error'))}")
    return result["access_token"]


def _graph_send(sender_email: str, token: str, payload: dict) -> Path | None:
    """
    Send a Graph sendMail payload, or — when VT_MAILER_DRY_RUN is enabled
    (the default) — write the rendered HTML to VT_MAILER_DRY_RUN_DIR and skip
    the network call entirely. Returns the written Path in dry-run mode, or
    None after a real send.
    """
    if _dry_run_enabled():
        out_dir = _dry_run_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        _DRY_RUN_COUNTER["n"] += 1
        msg = payload.get("message", {})
        subject = msg.get("subject", "no-subject")
        recipients = [r.get("emailAddress", {}).get("address", "") for r in msg.get("toRecipients", [])]
        ts = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        fname = f"{ts}_{_DRY_RUN_COUNTER['n']:03d}_{_safe_slug(subject)}.html"
        out_path = out_dir / fname

        meta_banner = (
            f"<!-- DRY RUN — would have sent via {sender_email} -->\n"
            f'<div style="background:#fdf6d8;border:1px solid #d68910;'
            f'padding:8px 12px;font-family:monospace;font-size:12px;color:#6b4a05;">'
            f"<strong>DRY RUN</strong> &middot; subject: {subject} &middot; "
            f"to: {', '.join(recipients) or '—'} &middot; from: {sender_email}"
            f"</div>\n"
        )
        body_html = msg.get("body", {}).get("content", "")
        out_path.write_text(meta_banner + body_html, encoding="utf-8")
        log.info("[DRY RUN] wrote %s (to=%s, subject=%r)", out_path, ",".join(recipients), subject)
        return out_path

    resp = requests.post(
        _GRAPH_SEND_MAIL.format(sender=sender_email),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
    )
    resp.raise_for_status()
    return None


# ---------------------------------------------------------------------------
# Small formatting helpers
# ---------------------------------------------------------------------------

def _fmt_date(d) -> str:
    return d.strftime("%d %b %Y") if d else "—"


def _join_names(names: list[str]) -> str:
    if not names:
        return ""
    if len(names) == 1:
        return names[0]
    if len(names) == 2:
        return f"{names[0]} and {names[1]}"
    return ", ".join(names[:-1]) + f", and {names[-1]}"


def _detail_text(r: dict) -> str:
    status = r["planning_status"]
    if status == "Data Error":
        return r["data_error_reason"] or "—"
    if status == "Blocked":
        return _join_names(r["waiting_on"]) or "—"
    if status in ("Violated", "At Risk"):
        slip = r["projected_slip_days"]
        if slip is None:
            return "—"
        if slip > 0:
            return f'+{slip}d (allowed {r["allowed_slippage"] or 0}d)'
        if r["due_date"] and date.today() > r["due_date"]:
            return "Already past due"
        return "On schedule"
    # On Track
    slip = r["projected_slip_days"]
    if slip is None:
        return "—"
    return "On schedule" if slip == 0 else f"{-slip}d margin"


def _root_cause_text(r: dict) -> tuple[str, str] | None:
    """Returns (kind, text) for the row directly under the task row, or None."""
    b = r.get("bottleneck")
    if b:
        days = b["delay_days"]
        unit = "day" if days == 1 else "days"
        return "critical", f'<b>Root cause:</b> Delayed by {b["predecessor"]} (Owned by {b["owner"]}) by {days} {unit}.'
    if r["planning_status"] == "Blocked" and r["waiting_on"]:
        names = _join_names(r["waiting_on"])
        return "waiting", f"<b>Waiting on:</b> {names} to reach Complete — no schedule impact yet."
    return None


def _suggestion_text(r: dict) -> str | None:
    if r["planning_status"] not in ("Blocked", "Violated"):
        return None
    s = r.get("reassignment_suggestion")
    if s is None:
        return (
            "<b>Suggested Resolution:</b> No available bandwidth detected elsewhere in the project "
            "right now — extending the due date or adding resourcing may be the only option."
        )
    if s["basis"] == "owner":
        return (
            f'<b>Suggested Resolution:</b> Reassign bandwidth from {s["candidate_owner"]} on '
            f'&quot;{s["candidate_task"]}&quot; ({s["slack_days"]}d margin) to help this task.'
        )
    if s["basis"] == "type":
        return (
            f'<b>Suggested Resolution:</b> &quot;{s["candidate_task"]}&quot; ({s["candidate_type"]}) has '
            f'{s["slack_days"]}d of schedule margin — consider reallocating its resourcing to help this task.'
        )
    return (
        f'<b>Suggested Resolution:</b> No same-type slack found; &quot;{s["candidate_task"]}&quot; '
        f'({s["candidate_type"]}) is the only task with spare margin ({s["slack_days"]}d) project-wide '
        f"— consider reallocating from there."
    )


# ---------------------------------------------------------------------------
# HTML assembly — nested tables, inline styles only
# ---------------------------------------------------------------------------

def _blame_row_html(kind: str, text: str) -> str:
    bg, color = {
        "critical": ("#fceceb", "#7a2016"),
        "waiting": ("#f4f5fb", "#4a4f6b"),
        "suggestion": ("#fff4c2", "#6b4a05"),
    }[kind]
    return (
        f'<tr><td colspan="5" style="padding:8px 14px 12px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0">'
        f'<tr><td style="background:{bg};color:{color};border-radius:8px;padding:9px 13px;'
        f'font-size:12px;line-height:1.5;">{text}</td></tr>'
        f"</table></td></tr>"
    )


def _task_row_html(r: dict) -> str:
    _, tint_bg, tint_text = _STATUS_COLORS[r["planning_status"]]
    return (
        f"<tr>"
        f'<td style="padding:12px 14px;border-top:1px solid #eef2f6;font-size:12.3px;color:#33404d;vertical-align:top;">'
        f'<div style="font-weight:600;color:#17324f;">{r["title"]}</div>'
        f'<div style="font-size:10.6px;color:#8b98a6;margin-top:3px;">Type: {r["task_type"]} | Priority: {r["priority"]}</div>'
        f'<div style="font-size:10.6px;color:#8b98a6;margin-top:1px;">Owner: {r["owner"]}</div>'
        f'</td>'
        f'<td style="padding:12px 14px;border-top:1px solid #eef2f6;font-size:12.3px;color:#33404d;vertical-align:top;">{_fmt_date(r["due_date"])}</td>'
        f'<td style="padding:12px 14px;border-top:1px solid #eef2f6;font-size:12.3px;color:#33404d;vertical-align:top;">{_fmt_date(r["ees"])}</td>'
        f'<td style="padding:12px 14px;border-top:1px solid #eef2f6;font-size:12.3px;color:#33404d;vertical-align:top;">{_detail_text(r)}</td>'
        f'<td style="padding:12px 14px;border-top:1px solid #eef2f6;vertical-align:top;">'
        f'<span style="display:inline-block;background:{tint_bg};color:{tint_text};border-radius:20px;'
        f'padding:3px 11px;font-size:10.3px;font-weight:600;white-space:nowrap;">{r["planning_status"]}</span>'
        f"</td></tr>"
    )


def _status_group_html(status: str, group: list[dict]) -> str:
    if not group:
        return ""
    accent, tint_bg, tint_text = _STATUS_COLORS[status]
    count_label = "task" if len(group) == 1 else "tasks"

    rows = []
    for r in group:
        rows.append(_task_row_html(r))
        cause = _root_cause_text(r)
        if cause:
            rows.append(_blame_row_html(cause[0], cause[1]))
        suggestion = _suggestion_text(r)
        if suggestion:
            rows.append(_blame_row_html("suggestion", suggestion))

    detail_label = _DETAIL_LABEL[status]
    header_row = (
        f'<tr style="background:{tint_bg};color:{tint_text};">'
        f'<th style="text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;">Task</th>'
        f'<th style="text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;">Due</th>'
        f'<th style="text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;">EES</th>'
        f'<th style="text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;">{detail_label}</th>'
        f'<th style="text-align:left;font-size:9.5px;letter-spacing:.8px;text-transform:uppercase;padding:10px 14px;">Status</th>'
        f"</tr>"
    )

    return (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="background:{accent};border-radius:10px;">'
        f'<tr><td style="padding:12px 18px;color:#fff;font-size:13.5px;font-weight:700;">{_RIBBON_TEXT[status]}</td>'
        f'<td align="right" style="padding:12px 18px;">'
        f'<span style="background:rgba(255,255,255,.22);color:#fff;border-radius:20px;padding:3px 11px;'
        f'font-size:11px;font-weight:600;">{len(group)} {count_label}</span>'
        f"</td></tr></table>"
        f'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:0 0 12px 12px;">'
        f"{header_row}{''.join(rows)}"
        f"</table>"
    )


def _overview_html(counts: dict, total: int) -> str:
    bubbles = "".join(
        f'<td width="20%" align="center" style="padding:25px 6px 11px 6px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border:1px solid #d7e4ef;border-radius:20px;">'
        f'<tr><td align="center" style="padding:11px 6px;">'
        f'<span style="font-weight:800;font-size:15px;color:{_STATUS_COLORS[s][0]};">{counts.get(s, 0)}</span>'
        f'<span style="font-size:11.5px;font-weight:600;color:#33404d;"> {s}</span>'
        f"</td></tr></table>"
        f"</td>"
        for s in _GROUP_ORDER
    )
    bar = "".join(
        f'<td width="{round(100 * counts.get(s, 0) / total) if total else 0}%" '
        f'style="background:{_STATUS_COLORS[s][0]};font-size:1px;line-height:10px;">&nbsp;</td>'
        for s in _GROUP_ORDER
        if counts.get(s, 0)
    )
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="height:10px;">'
        f"<tr>{bar}</tr></table>"
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>{bubbles}</tr></table>'
        f'<div style="font-size:11px;color:#7c8a99;text-align:left;margin-top:10px;">EES refers to Effective Earliest Start</div>'
    )


def _glossary_html() -> str:
    items = [
        ("#5b6472", "Data Error", "the record has missing or conflicting dates, or two tasks depend on each other in a loop, so a status can't be calculated until it's corrected."),
        ("#c0392b", "Violated", "the projected finish is already past the due date, or slips beyond the allowed buffer for its priority."),
        ("#4c51bf", "Blocked", "a task this one depends on hasn't finished yet, so it can't start on schedule."),
        ("#d68910", "At Risk", "the projected finish slips into the allowed buffer, or a resource conflict was detected — not late yet, but worth watching."),
        ("#1f9d78", "On Track", "projected to finish on or before the due date, with no conflicts or unresolved dependencies."),
    ]
    rows = "".join(
        f'<tr><td style="padding:7px 0;font-size:12px;color:#445468;line-height:1.5;">'
        f'<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{color};margin-right:8px;"></span>'
        f'<b style="color:#17324f;">{name}</b> — {desc}</td></tr>'
        for color, name, desc in items
    )
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="background:#fff;border:1px solid #d7e4ef;border-radius:14px;">'
        f'<tr><td style="padding:20px 24px;">'
        f'<div style="font-weight:700;font-size:12.5px;color:#17324f;margin-bottom:12px;">Status Criteria Reference</div>'
        f'<table width="100%" cellpadding="0" cellspacing="0">{rows}</table>'
        f"</td></tr></table>"
    )


def render_digest_html(results: list[dict], logo_src: str, generated_at: date | None = None) -> str:
    """
    Assembles the full Supervisory Digest email — nested tables, inline
    styles only, no CSS that depends on a stylesheet or modern browser
    engine. `logo_src` is a data: URI for local dry-run preview, or a
    "cid:..." reference matched to an inline Graph attachment for real sends.
    """
    generated_at = generated_at or date.today()
    counts = {s: 0 for s in _GROUP_ORDER}
    for r in results:
        counts[r["planning_status"]] = counts.get(r["planning_status"], 0) + 1
    total = len(results)

    groups_html = "".join(
        f'<tr><td style="padding:26px 28px 0;">{group_html}</td></tr>'
        for status in _GROUP_ORDER
        if (group_html := _status_group_html(status, [r for r in results if r["planning_status"] == status]))
    )

    banner = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="background:#17324f;border-radius:22px 22px 0 0;">'
        f'<tr><td style="padding:26px 32px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        f'<td><table cellpadding="0" cellspacing="0"><tr>'
        f'<td style="width:46px;height:46px;border-radius:12px;overflow:hidden;background:rgba(255,255,255,.08);">'
        f'<img src="{logo_src}" width="46" height="46" alt="Company logo" style="display:block;border-radius:12px;">'
        f"</td>"
        f'<td style="padding-left:14px;">'
        f'<div style="font-family:Arial,sans-serif;font-weight:700;font-size:18px;color:#fff;">Violation Report</div>'
        f'<div style="font-size:11px;color:#9db4c9;margin-top:1px;">Dynamic Operations Monitor</div>'
        f"</td></tr></table></td>"
        f'<td align="right">'
        f'<div style="font-weight:700;font-size:15px;color:#fff;">Supervisory Schedule Intelligence</div>'
        f'<div style="font-size:11.5px;color:#9db4c9;margin-top:3px;">{generated_at:%d %b %Y} &nbsp;&middot;&nbsp; {total} tasks monitored</div>'
        f"</td></tr></table>"
        f"</td></tr></table>"
    )

    footer = (
        f'<tr><td style="padding:26px 28px 26px;border-top:1px solid #d7e4ef;text-align:center;">'
        f'<div style="font-size:11px;color:#7c8a99;">Generated automatically by the Supervisory Architecture</div>'
        f"</td></tr>"
    )

    return (
        f'<!DOCTYPE html><html><head><meta charset="UTF-8">'
        f'<meta name="viewport" content="width=device-width,initial-scale=1"></head>'
        f'<body style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;">'
        f'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f6;padding:32px 0;">'
        f'<tr><td align="center" style="padding:0 12px;">'
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="max-width:880px;background:#eaf2f9;border-radius:22px;overflow:hidden;">'
        f"<tr><td>{banner}</td></tr>"
        f'<tr><td style="padding:26px 28px 0;">{_overview_html(counts, total)}</td></tr>'
        f"{groups_html}"
        f'<tr><td style="padding:30px 28px 0;">{_glossary_html()}</td></tr>'
        f"{footer}"
        f"</table></td></tr></table></body></html>"
    )


# ---------------------------------------------------------------------------
# Entry point used by vt_runner.py
# ---------------------------------------------------------------------------

def build_and_send_digest(results: list[dict], sender_email: str, recipients: list[str]) -> Path | None:
    """
    Renders the digest and either writes it to disk (dry run, the default)
    or sends it via Graph (only when VT_MAILER_DRY_RUN is explicitly off).
    Returns the written Path in dry-run mode, None after a real send.
    """
    logo_path = BASE_DIR / "bedpl.jpg"
    logo_bytes = logo_path.read_bytes() if logo_path.exists() else b""
    b64 = base64.b64encode(logo_bytes).decode() if logo_bytes else ""

    logo_src = f"data:image/jpeg;base64,{b64}" if _dry_run_enabled() else "cid:vtlogo"
    html = render_digest_html(results, logo_src)

    counts = {s: sum(1 for r in results if r["planning_status"] == s) for s in _GROUP_ORDER}
    subject = (
        f"Supervisory Digest — {counts['Violated']} violated, {counts['Blocked']} blocked, "
        f"{counts['Data Error']} data errors — {date.today():%d %b %Y}"
    )

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": r}} for r in recipients],
        }
    }
    if not _dry_run_enabled() and logo_bytes:
        payload["message"]["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "logo.jpg",
            "contentBytes": b64,
            "contentType": "image/jpeg",
            "isInline": True,
            "contentId": "vtlogo",
        }]

    token = _acquire_token(
        os.environ.get("VT_AZURE_TENANT_ID", ""),
        os.environ.get("VT_AZURE_CLIENT_ID", ""),
        os.environ.get("VT_AZURE_CLIENT_SECRET", ""),
    )
    return _graph_send(sender_email, token, payload)
