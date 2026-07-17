"""
vt_notion_client.py — read-only Notion API access for the Violation Tester.

Fetches every task row from the Violation Tester database, paginating through
`has_more`/`next_cursor` and retrying on rate-limit/server responses. Contains
no update/write endpoint calls anywhere — this module can only ever read.
"""
import re
import time

import requests

NOTION_API_VERSION = "2022-06-28"
_BASE_URL = "https://api.notion.com/v1"
_DATE_RE = re.compile(r"Change to\s*(\d{4}-\d{2}-\d{2})")
_NUMBER_RE = re.compile(r"(\d+)")
_TAG_RE = re.compile(r"(Effort|Lead)", re.IGNORECASE)


def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }


def _request_with_retry(method: str, url: str, headers: dict, json_body: dict, max_retries: int = 4) -> dict:
    """Issue a request with exponential backoff on 429 / 5xx responses."""
    resp = None
    for attempt in range(max_retries):
        resp = requests.request(method, url, headers=headers, json=json_body, timeout=15)
        if resp.status_code == 429 or resp.status_code >= 500:
            wait = float(resp.headers.get("Retry-After", 1.5 * (attempt + 1)))
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()
    return resp.json()


def _rich_text(prop: dict) -> str:
    return "".join(t.get("plain_text", "") for t in prop.get("rich_text", []))


def _parse_duration(raw_text: str) -> tuple[int | None, str | None]:
    """Extract (days, tag) from strings like '10 (Lead)', '6 (Effort)', or the
    typo'd '14 Effort)'. Returns (None, None) if no number is present."""
    num_match = _NUMBER_RE.search(raw_text or "")
    days = int(num_match.group(1)) if num_match else None
    tag_match = _TAG_RE.search(raw_text or "")
    tag = tag_match.group(1).title() if tag_match else None
    return days, tag


def _slip_history(history_text: str) -> list[str]:
    """All 'Change to <date>' entries in the order they appear (newest-first,
    matching how Due Date History is appended in this DB)."""
    return _DATE_RE.findall(history_text or "")


def _extract(page: dict, fields: dict) -> dict | None:
    props = page.get("properties", {})

    title = "".join(
        t.get("plain_text", "") for t in props.get(fields["title"], {}).get("title", [])
    ).strip()
    if not title:
        return None  # blank/malformed row — excluded here, not silently guessed downstream

    def date_val(key: str) -> str | None:
        d = props.get(fields[key], {}).get("date")
        return d.get("start") if d else None

    status_name = (props.get(fields["status"], {}).get("status") or {}).get("name", "")
    priority_name = (props.get(fields["priority"], {}).get("select") or {}).get("name", "")
    task_type_name = (props.get(fields["task_type"], {}).get("select") or {}).get("name", "")

    duration_days, duration_tag = _parse_duration(_rich_text(props.get(fields["effort_lead"], {})))

    allowed_slippage = None
    slip_formula = props.get(fields["allowed_slippage"], {}).get("formula", {})
    if slip_formula.get("type") == "number":
        allowed_slippage = slip_formula.get("number")
    elif slip_formula.get("type") == "string":
        # This formula uses format(), which renders the number as a string
        raw = (slip_formula.get("string") or "").strip()
        if raw.lstrip("-").isdigit():
            allowed_slippage = int(raw)

    dependencies_met = None
    deps_formula = props.get(fields["dependencies_met"], {}).get("formula", {})
    if deps_formula.get("type") == "boolean":
        dependencies_met = deps_formula.get("boolean")

    return {
        "page_id": page["id"],
        "url": page.get("url", ""),
        "title": title,
        "status": status_name,
        "priority": priority_name,
        "task_type": task_type_name,
        "start_date": date_val("start_date"),
        "due_date": date_val("due_date"),
        "due_date_history": _slip_history(_rich_text(props.get(fields["history"], {}))),
        "duration_days": duration_days,
        "duration_tag": duration_tag,
        "assignees": [p.get("name", "") for p in props.get(fields["assignee"], {}).get("people", [])],
        "blocked_by": [r["id"] for r in props.get(fields["blocked_by"], {}).get("relation", [])],
        "blocks": [r["id"] for r in props.get(fields["blocks"], {}).get("relation", [])],
        "dependencies_met": dependencies_met,
        "allowed_slippage": allowed_slippage,
        "created_time": page.get("created_time"),
    }


def fetch_all_tasks(
    token: str,
    database_id: str,
    fields: dict,
    project_filter: dict | None = None,
) -> list[dict]:
    """
    Fetch every row of the Violation Tester database (paginated), normalized
    into internal task records. Read-only — only ever issues a query POST.

    project_filter, if given, is {"field": <Project Link property name>,
    "page_id": <project page id>} — this DB holds multiple unrelated projects
    (e.g. Italy Visit, Domain Switchover) alongside the Violation Tester set,
    so the filter is applied server-side to scope the query to just one
    project's rows rather than fetching and discarding everything else.

    Rows with no title (blank/malformed records) are silently excluded here;
    callers that need to know how many were skipped should compare the
    returned length against a raw page count if that matters.
    """
    headers = _headers(token)
    url = f"{_BASE_URL}/databases/{database_id}/query"
    payload: dict = {"page_size": 100}
    if project_filter:
        payload["filter"] = {
            "property": project_filter["field"],
            "relation": {"contains": project_filter["page_id"]},
        }
    pages: list[dict] = []
    while True:
        data = _request_with_retry("POST", url, headers, payload)
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    tasks: list[dict] = []
    for page in pages:
        record = _extract(page, fields)
        if record is not None:
            tasks.append(record)
    return tasks
