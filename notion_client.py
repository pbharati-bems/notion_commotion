import re
import time
import requests
import concurrent.futures
from datetime import date, datetime

NOTION_API_VERSION = "2022-06-28"
_BASE_URL = "https://api.notion.com/v1"
_DATE_RE = re.compile(r"Change to(\d{4}-\d{2}-\d{2})")

# Field config keys:
#   title       – title property name
#   due_date    – date property name
#   status      – status property name
#   done_value  – exact "done" option string
#   owner       – assignee people property (or None to skip section)
#   reviewer    – reviewer people property (or None to skip section)
#   overdue     – formula boolean property (or None to skip)
#   history     – due-date-history rich_text property (or None to skip)
#   description – rich_text description property (or None to skip)
#   priority    – select priority property (or None to skip)
#   team        – multi_select team property (or None to skip)
#   project     – relation property linking to a project page (or None to skip)
#   parent_task – relation property linking to a parent task page (or None to skip)

# Field configs have moved to databases.json.
# get_task_report() accepts a fields dict loaded from there.


def get_users(token: str) -> list[dict]:
    """Return all human workspace members, sorted by name. Bots excluded."""
    headers = {"Authorization": f"Bearer {token}", "Notion-Version": NOTION_API_VERSION}
    users, params = [], {"page_size": 100}
    while True:
        resp = requests.get(f"{_BASE_URL}/users", headers=headers, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        for u in data.get("results", []):
            if u.get("type") == "person":
                users.append({"id": u["id"], "name": u.get("name", "(Unknown)")})
        if not data.get("has_more"):
            break
        params["start_cursor"] = data["next_cursor"]
    return sorted(users, key=lambda u: u["name"].lower())


def _page_title(page: dict) -> str:
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", [])).strip()
    return "".join(t.get("plain_text", "") for t in page.get("title", [])).strip() or "(Untitled)"


def _matches(user_obj: dict, user_id: str, user_name: str) -> bool:
    """Match a Notion user object by ID (exact) or by name (case-insensitive substring)."""
    if user_id:
        return user_obj.get("id") == user_id
    if user_name:
        return user_name.lower() in user_obj.get("name", "").lower()
    return False


def _user_in_props(props: dict, user_id: str, user_name: str) -> list[str]:
    """Return names of people-type properties that contain the target user."""
    return [
        name for name, prop in props.items()
        if prop.get("type") == "people"
        and any(_matches(p, user_id, user_name) for p in prop.get("people", []))
    ]


def _user_in_blocks(blocks: list, user_id: str, user_name: str) -> bool:
    """True if the target user is @mentioned in any block's rich text."""
    for block in blocks:
        content = block.get(block.get("type", ""), {})
        if not isinstance(content, dict):
            continue
        for rt in content.get("rich_text", []):
            if (rt.get("type") == "mention"
                    and rt.get("mention", {}).get("type") == "user"
                    and _matches(rt["mention"]["user"], user_id, user_name)):
                return True
    return False


def _fmt_last_edited(ts: str) -> str:
    if not ts:
        return ""
    try:
        return datetime.fromisoformat(ts.replace("Z", "+00:00")).strftime("%d %b %Y, %I:%M %p UTC")
    except ValueError:
        return ts


def scan_user_mentions(
    token: str,
    user_id: str = "",
    user_name: str = "",
    check_blocks: bool = False,
    max_workers: int = 8,
):
    """
    Generator: yields total / progress / result / done events.

    check_blocks=False  — fast, property-check only (data already in search response)
    check_blocks=True   — also scans @mention blocks, concurrent fetches via thread pool
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }

    # Collect all accessible pages
    pages: list[dict] = []
    payload: dict = {"query": "", "page_size": 100}
    while True:
        resp = requests.post(f"{_BASE_URL}/search", headers=headers, json=payload, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    yield {"type": "total", "total": len(pages)}

    if not check_blocks:
        # Fast path: all data is already in the search response
        for i, page in enumerate(pages):
            title = _page_title(page)
            yield {"type": "progress", "current": i + 1, "title": title}
            prop_matches = _user_in_props(page.get("properties", {}), user_id, user_name)
            if prop_matches:
                yield {
                    "type":         "result",
                    "title":        title,
                    "url":          page.get("url", ""),
                    "obj_type":     page.get("object", "page"),
                    "prop_matches": prop_matches,
                    "block_match":  False,
                    "last_edited":  _fmt_last_edited(page.get("last_edited_time", "")),
                }
        yield {"type": "done"}
        return

    # Deep path: also fetch blocks, but concurrently
    def _fetch_blocks(page: dict) -> list:
        for attempt in range(3):
            try:
                r = requests.get(
                    f"{_BASE_URL}/blocks/{page['id']}/children",
                    headers=headers, timeout=10,
                )
                if r.status_code == 429:
                    time.sleep(1.5 * (attempt + 1))
                    continue
                return r.json().get("results", []) if r.ok else []
            except Exception:
                return []
        return []

    completed = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_map = {pool.submit(_fetch_blocks, p): p for p in pages}
        for future in concurrent.futures.as_completed(future_map):
            page   = future_map[future]
            title  = _page_title(page)
            blocks = future.result()
            completed += 1
            yield {"type": "progress", "current": completed, "title": title}

            prop_matches = _user_in_props(page.get("properties", {}), user_id, user_name)
            block_match  = _user_in_blocks(blocks, user_id, user_name)

            if prop_matches or block_match:
                yield {
                    "type":         "result",
                    "title":        title,
                    "url":          page.get("url", ""),
                    "obj_type":     page.get("object", "page"),
                    "prop_matches": prop_matches,
                    "block_match":  block_match,
                    "last_edited":  _fmt_last_edited(page.get("last_edited_time", "")),
                }

    yield {"type": "done"}


def _slippage(history_text: str) -> tuple[int, str | None, str | None]:
    """(days_slipped, original_date, latest_date). History entries are newest-first."""
    dates = _DATE_RE.findall(history_text)
    if len(dates) < 2:
        return 0, dates[0] if dates else None, dates[0] if dates else None
    try:
        delta = (
            datetime.strptime(dates[0], "%Y-%m-%d").date()
            - datetime.strptime(dates[-1], "%Y-%m-%d").date()
        )
        return delta.days, dates[-1], dates[0]
    except ValueError:
        return 0, dates[-1], dates[0]


def _rich_text(prop: dict) -> str:
    return "".join(t.get("plain_text", "") for t in prop.get("rich_text", []))


def _fetch_page_refs(ids: set, headers: dict) -> dict:
    """Fetch {page_id: {name, url}} for a set of page IDs. Silently skips failures."""
    refs: dict = {}
    for pid in ids:
        try:
            resp = requests.get(f"{_BASE_URL}/pages/{pid}", headers=headers, timeout=10)
            if not resp.ok:
                continue
            page = resp.json()
            props = page.get("properties", {})
            title = ""
            for prop in props.values():
                if prop.get("type") == "title":
                    title = "".join(t.get("plain_text", "") for t in prop.get("title", []))
                    break
            refs[pid] = {"name": title.strip() or "(Untitled)", "url": page.get("url", "")}
        except Exception:
            pass
    return refs


def _extract(page: dict, f: dict) -> dict | None:
    props = page.get("properties", {})

    title = "".join(
        t.get("plain_text", "") for t in props.get(f["title"], {}).get("title", [])
    ).strip()
    if not title or title == "(Untitled)":
        return None

    due_obj = props.get(f["due_date"], {}).get("date")
    due_date = due_obj.get("start") if due_obj else None

    status_name = props.get(f["status"], {}).get("status", {}).get("name", "") or ""

    overdue = False
    if f.get("overdue"):
        overdue = props.get(f["overdue"], {}).get("formula", {}).get("boolean", False)

    slippage_days, original_date = 0, None
    if f.get("history"):
        slippage_days, original_date, _ = _slippage(_rich_text(props.get(f["history"], {})))

    description = ""
    if f.get("description"):
        description = _rich_text(props.get(f["description"], {}))[:120]

    priority = ""
    if f.get("priority"):
        priority = (props.get(f["priority"], {}).get("select") or {}).get("name", "")

    teams = []
    if f.get("team"):
        teams = [o.get("name", "") for o in props.get(f["team"], {}).get("multi_select", [])]

    has_owner = False
    if f.get("owner"):
        has_owner = bool(props.get(f["owner"], {}).get("people"))

    has_reviewer = False
    if f.get("reviewer"):
        has_reviewer = bool(props.get(f["reviewer"], {}).get("people"))

    # Store raw IDs for later enrichment — take first relation only
    project_rel = props.get(f.get("project", "") or "", {}).get("relation", [])
    parent_rel  = props.get(f.get("parent_task", "") or "", {}).get("relation", [])

    return {
        "name":           title,
        "url":            page.get("url", ""),
        "due_date":       due_date,
        "status":         status_name,
        "description":    description,
        "priority":       priority,
        "teams":          teams,
        "has_owner":      has_owner,
        "has_reviewer":   has_reviewer,
        "overdue":        overdue,
        "slippage_days":  slippage_days,
        "original_date":  original_date,
        # Temporary IDs — replaced with {name, url} dicts after ref fetch
        "_project_id":    project_rel[0]["id"] if project_rel else None,
        "_parent_id":     parent_rel[0]["id"]  if parent_rel  else None,
    }


def get_in_progress_tasks_by_owner(
    token: str, database_id: str, fields: dict, db_name: str
) -> dict[str, dict]:
    """
    Query tasks whose status == fields["in_progress_value"] and group them by owner email.

    Returns {owner_email: {"owner_name": str, "tasks": [task_dict, ...]}}
    where each task_dict contains: id, name, url, due_date, status, priority, db_name.

    Requires the Notion integration to have "Read user information including email
    addresses" permission; owners without a resolvable email are skipped.
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }
    in_progress = fields.get("in_progress_value", "In progress")
    payload: dict = {
        "filter": {
            "property": fields["status"],
            "status": {"equals": in_progress},
        }
    }
    url = f"{_BASE_URL}/databases/{database_id}/query"
    pages: list[dict] = []
    while True:
        resp = requests.post(url, headers=headers, json=payload)
        resp.raise_for_status()
        data = resp.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    # First pass — extract tasks and collect relation IDs for batch fetch
    raw_tasks: list[dict] = []
    relation_ids: set[str] = set()

    for page in pages:
        props = page.get("properties", {})

        title = "".join(
            t.get("plain_text", "") for t in props.get(fields["title"], {}).get("title", [])
        ).strip()
        if not title or title == "(Untitled)":
            continue

        due_obj = props.get(fields["due_date"], {}).get("date")
        due_date = due_obj.get("start") if due_obj else None
        status_name = props.get(fields["status"], {}).get("status", {}).get("name", "") or ""
        priority = ""
        if fields.get("priority"):
            priority = (props.get(fields["priority"], {}).get("select") or {}).get("name", "")

        proj_rel   = props.get(fields.get("project", "") or "", {}).get("relation", [])
        parent_rel = props.get(fields.get("parent_task", "") or "", {}).get("relation", [])
        proj_id    = proj_rel[0]["id"]   if proj_rel   else None
        parent_id  = parent_rel[0]["id"] if parent_rel else None
        if proj_id:   relation_ids.add(proj_id)
        if parent_id: relation_ids.add(parent_id)

        owner_field = fields.get("owner")
        owners = []
        if owner_field:
            for person in props.get(owner_field, {}).get("people", []):
                email = (person.get("person") or {}).get("email", "")
                if email:
                    owners.append({"email": email, "name": person.get("name", email)})

        raw_tasks.append({
            "id":        page["id"],
            "name":      title,
            "url":       page.get("url", ""),
            "due_date":  due_date,
            "status":    status_name,
            "priority":  priority,
            "db_name":   db_name,
            "owners":    owners,
            "_proj_id":  proj_id,
            "_par_id":   parent_id,
        })

    # Batch fetch relation names
    refs = _fetch_page_refs(relation_ids, headers) if relation_ids else {}

    # Second pass — enrich and group by owner
    by_owner: dict[str, dict] = {}
    for raw in raw_tasks:
        task = {
            "id":          raw["id"],
            "name":        raw["name"],
            "url":         raw["url"],
            "due_date":    raw["due_date"],
            "status":      raw["status"],
            "priority":    raw["priority"],
            "db_name":     raw["db_name"],
            "project":     refs.get(raw["_proj_id"])  if raw["_proj_id"]  else None,
            "parent_task": refs.get(raw["_par_id"])   if raw["_par_id"]   else None,
        }
        for owner in raw["owners"]:
            email = owner["email"]
            if email not in by_owner:
                by_owner[email] = {"owner_name": owner["name"], "tasks": []}
            by_owner[email]["tasks"].append(task)

    return by_owner


def get_task_report(token: str, database_id: str, fields: dict) -> dict:
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }

    # Fetch DB title
    db_resp = requests.get(f"{_BASE_URL}/databases/{database_id}", headers=headers)
    db_resp.raise_for_status()
    db_name = "".join(
        t.get("plain_text", "") for t in db_resp.json().get("title", [])
    ) or "Notion Database"

    # Query all non-Done tasks (paginated)
    payload: dict = {
        "filter": {
            "property": fields["status"],
            "status": {"does_not_equal": fields["done_value"]},
        }
    }
    url = f"{_BASE_URL}/databases/{database_id}/query"
    tasks: list[dict] = []

    while True:
        resp = requests.post(url, headers=headers, json=payload)
        resp.raise_for_status()
        data = resp.json()
        for page in data.get("results", []):
            task = _extract(page, fields)
            if task:
                tasks.append(task)
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    # Fetch project and parent task names/URLs (deduplicated)
    relation_ids = {
        t[k] for t in tasks for k in ("_project_id", "_parent_id") if t.get(k)
    }
    refs = _fetch_page_refs(relation_ids, headers) if relation_ids else {}

    # Enrich tasks and remove temporary ID keys
    for t in tasks:
        pid = t.pop("_project_id", None)
        prid = t.pop("_parent_id", None)
        t["project"]     = refs.get(pid)     if pid  else None
        t["parent_task"] = refs.get(prid)    if prid else None

    return {
        "db_name":    db_name,
        "total_open": len(tasks),
        "no_due_date":  [t for t in tasks if not t["due_date"]],
        "overdue":      [t for t in tasks if t["overdue"]],
        "max_slippage": sorted(
            [t for t in tasks if t["slippage_days"] > 0],
            key=lambda t: t["slippage_days"],
            reverse=True,
        )[:10],
        # None = field absent from this DB → section omitted in email
        "no_owner":    [t for t in tasks if not t["has_owner"]]    if fields.get("owner")    else None,
        "no_reviewer": [t for t in tasks if not t["has_reviewer"]] if fields.get("reviewer") else None,
    }


# ---------------------------------------------------------------------------
# MIS helpers
# ---------------------------------------------------------------------------

def get_all_tasks_for_mis(
    token: str, database_id: str, fields: dict, db_name: str
) -> list[dict]:
    """
    Fetch all non-done tasks with full field extraction, including assignee
    email addresses required by the MIS event detector.

    Returns a list of task dicts, each containing:
      id, name, url, due_date, status, priority, teams (list), description,
      slippage_days, original_date, db_name,
      assignees:    [{name, email}]   — owner/assignee people field
      team_members: [{name, email}]   — owner + reviewer combined, deduplicated
      project:      {name, url} | None
      parent_task:  {name, url} | None
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }

    payload: dict = {
        "filter": {
            "property": fields["status"],
            "status": {"does_not_equal": fields["done_value"]},
        }
    }
    url   = f"{_BASE_URL}/databases/{database_id}/query"
    pages: list[dict] = []
    while True:
        resp = requests.post(url, headers=headers, json=payload)
        resp.raise_for_status()
        data = resp.json()
        pages.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        payload["start_cursor"] = data["next_cursor"]

    raw_tasks: list[dict] = []
    relation_ids: set[str] = set()

    for page in pages:
        props = page.get("properties", {})

        title = "".join(
            t.get("plain_text", "") for t in props.get(fields["title"], {}).get("title", [])
        ).strip()
        if not title or title == "(Untitled)":
            continue

        due_obj  = props.get(fields["due_date"], {}).get("date")
        due_date = due_obj.get("start") if due_obj else None

        status_name = props.get(fields["status"], {}).get("status", {}).get("name", "") or ""

        slippage_days, original_date = 0, None
        all_dates: list = []
        if fields.get("history"):
            _hist = _rich_text(props.get(fields["history"], {}))
            slippage_days, original_date, _ = _slippage(_hist)
            all_dates = _DATE_RE.findall(_hist)  # newest-first

        description = ""
        if fields.get("description"):
            description = _rich_text(props.get(fields["description"], {}))

        priority = ""
        if fields.get("priority"):
            priority = (props.get(fields["priority"], {}).get("select") or {}).get("name", "")

        teams: list[str] = []
        if fields.get("team"):
            teams = [o.get("name", "") for o in props.get(fields["team"], {}).get("multi_select", [])]

        # Assignees — owner field, with email (requires user-email permission in integration)
        assignees: list[dict] = []
        owner_field = fields.get("owner")
        if owner_field:
            for person in props.get(owner_field, {}).get("people", []):
                email = (person.get("person") or {}).get("email", "")
                assignees.append({
                    "name":  person.get("name", email or "Unknown"),
                    "email": email,
                })

        # Team members = owner(s) + reviewer(s), deduplicated by name
        team_members: list[dict] = list(assignees)
        reviewer_field = fields.get("reviewer")
        reviewer_names = ""
        if reviewer_field:
            existing_names = {m["name"] for m in team_members}
            _rev_people = props.get(reviewer_field, {}).get("people", [])
            reviewer_names = ", ".join(p.get("name", "") for p in _rev_people if p.get("name"))
            for person in _rev_people:
                email = (person.get("person") or {}).get("email", "")
                name  = person.get("name", email or "Unknown")
                if name not in existing_names:
                    team_members.append({"name": name, "email": email})
                    existing_names.add(name)

        # Overdue + days overdue
        overdue_val = False
        if fields.get("overdue"):
            overdue_val = props.get(fields["overdue"], {}).get("formula", {}).get("boolean", False)
        days_overdue = 0
        if overdue_val and due_date:
            try:
                days_overdue = max(0, (date.today() - date.fromisoformat(due_date)).days)
            except ValueError:
                pass

        proj_rel     = props.get(fields.get("project",     "") or "", {}).get("relation", [])
        par_rel      = props.get(fields.get("parent_task", "") or "", {}).get("relation", [])
        blocking_rel = props.get(fields.get("blocking",    "") or "", {}).get("relation", [])
        proj_id      = proj_rel[0]["id"]     if proj_rel     else None
        parent_id    = par_rel[0]["id"]      if par_rel      else None
        blocking_ids = [r["id"] for r in blocking_rel]
        if proj_id:   relation_ids.add(proj_id)
        if parent_id: relation_ids.add(parent_id)
        for bid in blocking_ids:
            relation_ids.add(bid)

        raw_tasks.append({
            "id":             page["id"],
            "name":           title,
            "url":            page.get("url", ""),
            "due_date":       due_date,
            "status":         status_name,
            "priority":       priority,
            "teams":          teams,
            "description":    description,
            "slippage_days":  slippage_days,
            "original_date":  original_date,
            "all_dates":      all_dates,
            "n_reschedules":  max(len(all_dates) - 1, 0),
            "overdue":        overdue_val,
            "days_overdue":   days_overdue,
            "reviewer_names": reviewer_names,
            "db_name":        db_name,
            "assignees":      assignees,
            "team_members":   team_members,
            "_project_id":    proj_id,
            "_parent_id":     parent_id,
            "_blocking_ids":  blocking_ids,
        })

    refs = _fetch_page_refs(relation_ids, headers) if relation_ids else {}

    tasks: list[dict] = []
    for raw in raw_tasks:
        pid   = raw.pop("_project_id",   None)
        prid  = raw.pop("_parent_id",    None)
        bids  = raw.pop("_blocking_ids", [])
        raw["project"]     = refs.get(pid)  if pid  else None
        raw["parent_task"] = refs.get(prid) if prid else None
        raw["blocking"]    = [refs[b] for b in bids if b in refs]
        tasks.append(raw)

    return tasks


def get_subtasks(
    token: str, database_id: str, fields: dict,
    parent_page_id: str, db_name: str,
) -> list[dict]:
    """
    Return tasks in this database whose parent_task relation contains parent_page_id.
    Used to populate the sub-task section of the "project started" email.
    Returns [] if the parent_task field is not configured or the query fails.
    """
    parent_field = fields.get("parent_task")
    if not parent_field:
        return []

    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json",
    }
    payload = {
        "filter": {
            "property": parent_field,
            "relation": {"contains": parent_page_id},
        }
    }
    try:
        resp = requests.post(
            f"{_BASE_URL}/databases/{database_id}/query",
            headers=headers, json=payload, timeout=15,
        )
        resp.raise_for_status()
        pages = resp.json().get("results", [])
    except Exception:
        return []

    sub_tasks: list[dict] = []
    for page in pages:
        props = page.get("properties", {})
        title = "".join(
            t.get("plain_text", "") for t in props.get(fields["title"], {}).get("title", [])
        ).strip()
        if not title:
            continue

        status_name = props.get(fields["status"], {}).get("status", {}).get("name", "") or ""
        due_obj     = props.get(fields["due_date"], {}).get("date")
        due_date    = due_obj.get("start") if due_obj else None

        owner_field = fields.get("owner")
        owner_name  = ""
        if owner_field:
            people = props.get(owner_field, {}).get("people", [])
            if people:
                owner_name = people[0].get("name", "")

        sub_tasks.append({
            "name":       title,
            "url":        page.get("url", ""),
            "status":     status_name,
            "due_date":   due_date,
            "owner_name": owner_name,
        })

    return sub_tasks
