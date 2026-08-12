#!/usr/bin/env python3
"""
Shared Microsoft Planner / Graph business logic.

This module holds everything both the CLI scripts (planner_import.py,
planner_export.py) and the Flask web app use: the Graph HTTP client, task
matching, the upsert logic, checklist diffing, priority/date parsing, the
local match-state file, and the receipt writer.

It deliberately knows *nothing* about how a token is obtained. Callers pass
in a token provider:

    - the CLI passes one backed by app-only client credentials
    - the web app passes one backed by delegated auth code flow (MSAL
      acquire_token_silent)

...and nothing about how output is presented: every progress message goes
through an `emit` callable (print for the CLI, a stream writer for the web
app) rather than to stdout directly.

Delegated vs. application permissions
------------------------------------
Under *application* permissions the app can reach any plan in the tenant.
Under *delegated* permissions, Planner authorizes against the signed-in
user's membership of the plan's backing Microsoft 365 group — being a tenant
admin does not by itself grant Planner data access. A plan you can see in the
Planner UI but are not a member of will return 403.

Resolving assignee emails (GET /users/{email}) needs delegated
`User.ReadBasic.All` (or User.Read.All). Plain `User.Read` only grants /me,
so every `assigned_to` value would silently fail to resolve. That case is
detected explicitly and aborts the run — see GraphPermissionError.
"""

import csv
import json
import os
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Callable, Iterable, List, Optional
from urllib.parse import quote

import requests

GRAPH = "https://graph.microsoft.com/v1.0"

DEFAULT_STATE_DIR = "planner_state"

# Delegated scopes the web app requests. Tasks.ReadWrite covers plans,
# buckets, tasks and task details; User.ReadBasic.All is required to look up
# assignees by email. Group.Read.All is deliberately NOT requested: bucket and
# task listing authorize off group membership, not a group scope. If a bucket
# call ever 403s, read the Graph error body surfaced by check() before adding
# scopes speculatively.
DELEGATED_SCOPES = ["Tasks.ReadWrite", "User.ReadBasic.All"]

PRIORITY_MAP = {
    "urgent": 1,
    "important": 3,
    "medium": 5,
    "low": 9,
}

# Microsoft's documented priority buckets (Planner UI only ever *sets* 1/3/5/9,
# but reads back correctly across the full 0-10 range):
#   0-1 Urgent, 2-4 Important, 5-7 Medium, 8-10 Low
PRIORITY_BUCKETS = [
    (0, 1, "Urgent"),
    (2, 4, "Important"),
    (5, 7, "Medium"),
    (8, 10, "Low"),
]

FIELDNAMES = [
    "id", "title", "bucket", "start_date", "due_date", "priority",
    "assigned_to", "description", "percent_complete", "checklist",
]

RECEIPT_FIELDNAMES = [
    "row", "title", "bucket", "action", "task_id", "task_url", "timestamp",
]


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

class PlannerError(Exception):
    """Base class for errors this module raises deliberately."""


class GraphPermissionError(PlannerError):
    """
    The signed-in identity lacks a permission the run needs.

    Distinct from "this particular user/plan doesn't exist": a 403 means the
    app registration is misconfigured, and continuing would silently produce
    wrong results (e.g. an import that drops every assignment). Runs abort on
    this rather than degrading.
    """


class DryRunViolation(PlannerError):
    """A write was attempted during a dry run. Should be unreachable."""


# ---------------------------------------------------------------------------
# Graph client
# ---------------------------------------------------------------------------

TokenProvider = Callable[..., str]


class GraphClient:
    """
    Thin Graph wrapper with throttling backoff, token refresh, and paging.

    `token_provider` is normally a callable taking an optional
    `force_refresh` keyword and returning a bearer token string. A plain
    string is also accepted (handy in tests), in which case no refresh is
    possible.

    Taking a provider rather than a token matters for real imports: they make
    one to three Graph calls per row with throttling backoff, which can
    comfortably outlive a single access token. A mid-run 401 would otherwise
    abandon the run with Planner already partially written.
    """

    MAX_THROTTLE_RETRIES = 5

    def __init__(self, token_provider: TokenProvider, emit: Callable[[str], None] = print, session=None):
        if callable(token_provider):
            self._token_provider = token_provider
        else:
            _static = token_provider
            self._token_provider = lambda force_refresh=False: _static
        self.emit = emit
        self.session = session or requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {self._token_provider()}", "Content-Type": "application/json"}
        )

    def _refresh_token(self) -> bool:
        try:
            token = self._token_provider(force_refresh=True)
        except TypeError:
            token = self._token_provider()
        except Exception as exc:  # noqa: BLE001 - surfaced to the user, not swallowed
            self.emit(f"  Token refresh failed: {exc}")
            return False
        if not token:
            return False
        self.session.headers["Authorization"] = f"Bearer {token}"
        return True

    def request(self, method, url, **kwargs):
        resp = None
        refreshed = False
        attempt = 0
        while attempt < self.MAX_THROTTLE_RETRIES:
            resp = self.session.request(method, url, **kwargs)

            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                self.emit(f"  Rate limited, waiting {wait}s...")
                time.sleep(wait)
                attempt += 1
                continue

            if resp.status_code == 401 and not refreshed:
                # Long imports can outlive the access token; take one shot at
                # a silent refresh before giving up on the run.
                refreshed = True
                self.emit("  Access token expired, refreshing...")
                if self._refresh_token():
                    continue

            return resp
        return resp

    def get(self, url, **kw):
        return self.request("GET", url, **kw)

    def post(self, url, **kw):
        return self.request("POST", url, **kw)

    def patch(self, url, **kw):
        return self.request("PATCH", url, **kw)

    def paginate(self, url, **kw) -> List[dict]:
        """
        Follow @odata.nextLink and return every item.

        Not optional: a partially-read task list makes the bucket+title
        fallback match miss, which makes the importer create duplicates of
        tasks that already exist.
        """
        items: List[dict] = []
        while url:
            resp = self.get(url, **kw)
            check(resp, self.emit)
            data = resp.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            kw = {}  # nextLink already carries the query string
        return items


class ReadOnlyGraphClient(GraphClient):
    """
    Dry-run client: refuses anything that isn't a GET.

    Dry-run safety is enforced here, structurally, rather than by a branch in
    the run loop — so no future reshuffling of that loop can quietly make a
    dry run write to Planner. The corresponding test asserts zero non-GET
    calls across a full dry-run import.
    """

    def request(self, method, url, **kwargs):
        if method.upper() != "GET":
            raise DryRunViolation(f"Dry run attempted a {method.upper()} to {url}")
        return super().request(method, url, **kwargs)


def check(resp, emit: Callable[[str], None] = print):
    """Raise with Graph's actual error body surfaced, instead of a bare status code."""
    if not resp.ok:
        emit(f"  Graph error {resp.status_code} for {resp.request.method} {resp.url}")
        emit(f"  Response body: {resp.text}")
    resp.raise_for_status()


# ---------------------------------------------------------------------------
# Buckets and tasks: read existing plan state
# ---------------------------------------------------------------------------

def get_buckets(client: GraphClient, plan_id):
    """Returns {bucket_name: bucket_id}."""
    buckets = client.paginate(f"{GRAPH}/planner/plans/{plan_id}/buckets")
    return {b["name"]: b["id"] for b in buckets}


def get_buckets_by_id(client: GraphClient, plan_id):
    """Returns {bucket_id: bucket_name}."""
    buckets = client.paginate(f"{GRAPH}/planner/plans/{plan_id}/buckets")
    return {b["id"]: b["name"] for b in buckets}


def create_bucket(client: GraphClient, plan_id, name, order_hint=" !"):
    resp = client.post(
        f"{GRAPH}/planner/buckets",
        json={"name": name, "planId": plan_id, "orderHint": order_hint},
    )
    check(resp, client.emit)
    return resp.json()["id"]


def get_existing_tasks(client: GraphClient, plan_id):
    return client.paginate(f"{GRAPH}/planner/plans/{plan_id}/tasks")


def get_task_details(client: GraphClient, task_id):
    resp = client.get(f"{GRAPH}/planner/tasks/{task_id}/details")
    check(resp, client.emit)
    return resp.json()


# ---------------------------------------------------------------------------
# Matching: local state file + live bucket/title fallback
# ---------------------------------------------------------------------------

def normalize_key(bucket, title):
    return f"{(bucket or '').strip().lower()}::{(title or '').strip().lower()}"


def row_key(row):
    explicit = (row.get("id") or "").strip()
    if explicit:
        return f"id:{explicit}"
    return f"bt:{normalize_key(row.get('bucket') or 'To do', row['title'])}"


class StateStore:
    """
    The local match-state file (planner_state/<plan-id>.json).

    Saved incrementally after every row, atomically, so a run that dies
    part-way still records what it actually wrote. The bucket+title fallback
    covers a lost state file, but there's no reason to lean on it.

    The state directory is caller-supplied rather than relative to the process
    working directory, because the web app's cwd is not the project root. If
    this ever serves more than one user, state needs to be keyed per user as
    well as per plan.
    """

    def __init__(self, plan_id, state_dir=DEFAULT_STATE_DIR):
        self.plan_id = plan_id
        self.state_dir = state_dir
        self.path = os.path.join(state_dir, f"{plan_id}.json")
        self._data = {}
        os.makedirs(state_dir, exist_ok=True)
        if os.path.exists(self.path):
            with open(self.path) as f:
                self._data = json.load(f)

    def get(self, key):
        return self._data.get(key)

    def set(self, key, task_id, bucket_id):
        self._data[key] = {"task_id": task_id, "bucket_id": bucket_id}

    def save(self):
        tmp = f"{self.path}.tmp"
        with open(tmp, "w") as f:
            json.dump(self._data, f, indent=2, sort_keys=True)
        os.replace(tmp, self.path)


# ---------------------------------------------------------------------------
# Users / assignments
# ---------------------------------------------------------------------------

def resolve_user(client: GraphClient, email, cache):
    """
    Resolve an email to a user id.

    404 means that specific person isn't in the directory: warn, skip the
    assignment, carry on (the historical behavior). 401/403 means the app
    registration can't read users at all, which would silently strip every
    assignment from the import — that aborts the run instead.
    """
    if email in cache:
        return cache[email]
    # safe="@" keeps the wire format identical to a plain email while still
    # escaping anything unexpected in the value.
    resp = client.get(f"{GRAPH}/users/{quote(email, safe='@')}", params={"$select": "id"})
    if resp.status_code == 200:
        uid = resp.json()["id"]
        cache[email] = uid
        return uid
    if resp.status_code in (401, 403):
        raise GraphPermissionError(
            f"Not permitted to look up users in the directory (HTTP {resp.status_code} "
            f"for GET /users/{email}). Delegated 'User.ReadBasic.All' is required to "
            f"resolve assignees by email; plain 'User.Read' only grants access to your "
            f"own profile. Graph said: {resp.text}"
        )
    client.emit(f"  Warning: could not resolve user '{email}' ({resp.status_code}); skipping assignment.")
    cache[email] = None
    return None


def split_emails(emails_str):
    return [e.strip() for e in (emails_str or "").split(";") if e.strip()]


def build_assignments(client: GraphClient, emails_str, cache):
    if not emails_str:
        return {}
    assignments = {}
    for email in split_emails(emails_str):
        uid = resolve_user(client, email, cache)
        if uid:
            assignments[uid] = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !",
            }
    return assignments


def preflight_user_resolution(client: GraphClient, rows, cache):
    """
    Resolve one assignee before the run starts writing.

    Turns a missing User.ReadBasic.All into an error raised before the first
    task is created, rather than an import that "succeeds" with every
    assignment dropped.
    """
    for row in rows:
        emails = split_emails(row.get("assigned_to"))
        if emails:
            resolve_user(client, emails[0], cache)
            return


# ---------------------------------------------------------------------------
# Field parsing helpers
# ---------------------------------------------------------------------------

def parse_priority(value):
    if value is None or value == "":
        return None
    value = str(value).strip()
    if value.lower() in PRIORITY_MAP:
        return PRIORITY_MAP[value.lower()]
    try:
        return max(0, min(10, int(value)))
    except ValueError:
        return None


def priority_to_label(value):
    if value is None:
        return ""
    try:
        value = int(value)
    except (TypeError, ValueError):
        return ""
    for low, high, label in PRIORITY_BUCKETS:
        if low <= value <= high:
            return label
    return str(value)


def to_graph_date(value):
    if not value:
        return None
    value = str(value).strip()
    try:
        dt = datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None
    return dt.strftime("%Y-%m-%dT00:00:00Z")


def date_only(iso_value):
    if not iso_value:
        return ""
    return iso_value.split("T")[0]


def task_url(plan_id, task_id):
    return f"https://planner.cloud.microsoft/webui/plan/{plan_id}/view/board/task/{task_id}"


# ---------------------------------------------------------------------------
# Checklist diffing (pure function so it's easy to test in isolation)
# ---------------------------------------------------------------------------

def build_checklist_patch(desired_items, existing_checklist):
    """
    desired_items: list of item title strings, in the order they should appear.
    existing_checklist: the current plannerTaskDetails "checklist" dict
      (id -> {"title": ..., "isChecked": ...}), or None/{} if there isn't one.

    Returns a patch dict suitable for the checklist property (matched items
    keep their id and isChecked state; new items get a fresh id; removed
    items are set to None), or None if the desired state already matches.
    """
    existing_by_title = {
        v["title"].strip().lower(): (k, v.get("isChecked", False))
        for k, v in (existing_checklist or {}).items()
    }
    desired_titles_norm = {d.strip().lower() for d in desired_items}
    existing_titles_norm = set(existing_by_title.keys())

    if desired_titles_norm == existing_titles_norm:
        return None

    patch = {}
    for item in desired_items:
        norm = item.strip().lower()
        if norm in existing_by_title:
            guid, checked = existing_by_title[norm]
            patch[guid] = {
                "@odata.type": "#microsoft.graph.plannerChecklistItem",
                "title": item.strip(),
                "isChecked": checked,
            }
        else:
            patch[str(uuid.uuid4())] = {
                "@odata.type": "#microsoft.graph.plannerChecklistItem",
                "title": item.strip(),
                "isChecked": False,
            }
    for title_norm, (guid, _checked) in existing_by_title.items():
        if title_norm not in desired_titles_norm:
            patch[guid] = None
    return patch


# ---------------------------------------------------------------------------
# Task details (description / checklist)
# ---------------------------------------------------------------------------

def set_task_details(client: GraphClient, task_id, description, checklist):
    """Returns True if anything was actually changed."""
    description = (description or "").strip()
    checklist = (checklist or "").strip()
    if not description and not checklist:
        return False

    resp = client.get(f"{GRAPH}/planner/tasks/{task_id}/details")
    check(resp, client.emit)
    etag = resp.headers.get("ETag")
    current = resp.json()

    body = {}
    if description and current.get("description") != description:
        body["description"] = description

    if checklist:
        desired_items = [c.strip() for c in checklist.split(";") if c.strip()]
        patch = build_checklist_patch(desired_items, current.get("checklist"))
        if patch is not None:
            body["checklist"] = patch

    if not body:
        return False

    r = client.patch(
        f"{GRAPH}/planner/tasks/{task_id}/details",
        json=body,
        headers={"If-Match": etag},
    )
    if r.status_code not in (200, 204):
        client.emit(f"  Warning: failed to set details for task {task_id}: {r.status_code} {r.text}")
        return False
    return True


# ---------------------------------------------------------------------------
# Create / update
# ---------------------------------------------------------------------------

def create_task(client: GraphClient, plan_id, bucket_id, row, user_cache):
    body = {"planId": plan_id, "bucketId": bucket_id, "title": row["title"]}

    due = to_graph_date(row.get("due_date"))
    if due:
        body["dueDateTime"] = due
    start = to_graph_date(row.get("start_date"))
    if start:
        body["startDateTime"] = start

    priority = parse_priority(row.get("priority"))
    if priority is not None:
        body["priority"] = priority

    pct = row.get("percent_complete")
    if pct not in (None, ""):
        try:
            body["percentComplete"] = int(pct)
        except ValueError:
            pass

    assignments = build_assignments(client, row.get("assigned_to", ""), user_cache)
    if assignments:
        body["assignments"] = assignments

    resp = client.post(f"{GRAPH}/planner/tasks", json=body)
    if resp.status_code != 201:
        client.emit(f"  Failed to create '{row['title']}': {resp.status_code} {resp.text}")
        return None
    task = resp.json()

    set_task_details(client, task["id"], row.get("description"), row.get("checklist"))
    return task


def update_task(client: GraphClient, task, bucket_id, row, user_cache):
    """Returns 'updated', 'unchanged', or 'failed'."""
    etag = task.get("@odata.etag")
    body = {}

    if task.get("bucketId") != bucket_id:
        body["bucketId"] = bucket_id
    if task.get("title") != row["title"]:
        body["title"] = row["title"]

    due = to_graph_date(row.get("due_date"))
    if due and task.get("dueDateTime") != due:
        body["dueDateTime"] = due
    start = to_graph_date(row.get("start_date"))
    if start and task.get("startDateTime") != start:
        body["startDateTime"] = start

    priority = parse_priority(row.get("priority"))
    if priority is not None and task.get("priority") != priority:
        body["priority"] = priority

    # percentComplete is intentionally never touched here — it tracks live
    # progress in Planner, which the source file doesn't know about.

    assigned_to_raw = (row.get("assigned_to") or "").strip()
    if assigned_to_raw:
        desired = build_assignments(client, assigned_to_raw, user_cache)
        existing_ids = set((task.get("assignments") or {}).keys())
        if not desired and existing_ids:
            # The row named assignees but none of them resolved (a typo, or
            # someone who has left). Syncing an empty set here would silently
            # unassign everyone on the strength of a bad email, so leave the
            # existing assignments alone instead.
            client.emit(
                f"  Warning: no assignee on '{row['title']}' could be resolved; "
                f"leaving existing assignments unchanged."
            )
        else:
            assignment_patch = {}
            for uid, obj in desired.items():
                if uid not in existing_ids:
                    assignment_patch[uid] = obj
            for uid in existing_ids - set(desired.keys()):
                assignment_patch[uid] = None
            if assignment_patch:
                body["assignments"] = assignment_patch

    task_body_changed = bool(body)
    if body:
        resp = client.patch(
            f"{GRAPH}/planner/tasks/{task['id']}", json=body, headers={"If-Match": etag}
        )
        if resp.status_code not in (200, 204):
            client.emit(f"  Failed to update '{row['title']}': {resp.status_code} {resp.text}")
            return "failed"

    details_changed = set_task_details(client, task["id"], row.get("description"), row.get("checklist"))

    return "updated" if (task_body_changed or details_changed) else "unchanged"


# ---------------------------------------------------------------------------
# Input file loading
# ---------------------------------------------------------------------------

def load_rows(source, filename: Optional[str] = None) -> List[dict]:
    """
    Load task rows from a path, a binary stream, or bytes.

    The web app hands over an uploaded file object; the CLI hands over a path.
    Format is chosen by extension where available but confirmed by sniffing
    the content, since an uploaded filename isn't trustworthy.
    """
    if hasattr(source, "read"):
        raw = source.read()
        filename = filename or getattr(source, "filename", None) or getattr(source, "name", "")
    elif isinstance(source, (bytes, bytearray)):
        raw = bytes(source)
        filename = filename or ""
    else:
        filename = filename or str(source)
        with open(source, "rb") as f:
            raw = f.read()

    if isinstance(raw, str):
        text = raw
    else:
        text = raw.decode("utf-8-sig", errors="replace")
    text = text.lstrip("﻿")

    looks_json = text.lstrip()[:1] in ("{", "[")
    if looks_json or (filename or "").lower().endswith(".json"):
        if not looks_json:
            raise PlannerError("File is named .json but does not contain JSON.")
        data = json.loads(text)
        rows = data if isinstance(data, list) else data.get("tasks", [])
        return [dict(r) for r in rows]

    return list(csv.DictReader(text.splitlines()))


def normalize_row(raw_row: dict) -> dict:
    return {str(k).strip().lower(): v for k, v in raw_row.items()}


# ---------------------------------------------------------------------------
# Receipt log
# ---------------------------------------------------------------------------

def write_receipt(receipt_rows, out_path):
    parent = os.path.dirname(out_path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=RECEIPT_FIELDNAMES)
        writer.writeheader()
        writer.writerows(receipt_rows)
    return out_path


def default_receipt_path(input_path):
    stem = os.path.splitext(os.path.basename(input_path))[0]
    return f"{stem}_receipt.csv"


# ---------------------------------------------------------------------------
# Import driver
# ---------------------------------------------------------------------------

@dataclass
class ImportSummary:
    dry_run: bool = False
    created: int = 0
    updated: int = 0
    unchanged: int = 0
    failed: int = 0
    would_create: int = 0
    would_update: int = 0
    skipped: int = 0
    receipt_path: Optional[str] = None
    state_path: Optional[str] = None
    receipt_rows: List[dict] = field(default_factory=list)
    aborted: Optional[str] = None


def run_import(
    client: GraphClient,
    plan_id: str,
    rows: Iterable[dict],
    *,
    dry_run: bool = False,
    state_dir: str = DEFAULT_STATE_DIR,
    receipt_path: Optional[str] = None,
    emit: Callable[[str], None] = print,
) -> ImportSummary:
    """
    Create or update every row in `rows` against `plan_id`.

    Safety properties, all of which have tests:
      - dry_run makes no writes at all (enforced by ReadOnlyGraphClient, which
        the caller is expected to pass; the branch below is belt and braces)
      - matching falls back to bucket+title when neither the state file nor an
        explicit id resolves
      - percent_complete is applied on create only
      - blank fields never clear an existing value
    """
    rows = list(rows)
    summary = ImportSummary(dry_run=dry_run)

    emit(f"Loading existing buckets for plan {plan_id}...")
    buckets = get_buckets(client, plan_id)
    bucket_name_by_id = {v: k for k, v in buckets.items()}

    emit("Loading existing tasks...")
    existing_tasks = get_existing_tasks(client, plan_id)
    tasks_by_id = {t["id"]: t for t in existing_tasks}
    tasks_by_bucket_title = {}
    for t in existing_tasks:
        bname = bucket_name_by_id.get(t["bucketId"], "")
        tasks_by_bucket_title[normalize_key(bname, t["title"])] = t

    state = StateStore(plan_id, state_dir)
    summary.state_path = state.path
    user_cache = {}

    # Fail fast on a missing directory-read permission, before anything is
    # written, rather than silently importing every task with no assignees.
    preflight_user_resolution(client, [normalize_row(r) for r in rows], user_cache)

    buckets_previewed = set()

    try:
        for i, raw_row in enumerate(rows, 1):
            row = normalize_row(raw_row)
            title = (row.get("title") or "").strip()
            if not title:
                emit(f"[{i}] Skipping row with no title.")
                summary.skipped += 1
                continue
            row["title"] = title
            bucket_name = (row.get("bucket") or "To do").strip()

            key = row_key(row)
            explicit_id = (row.get("id") or "").strip()
            state_entry = state.get(key)
            if state_entry and state_entry.get("task_id") in tasks_by_id:
                match = tasks_by_id[state_entry["task_id"]]
            elif explicit_id and explicit_id in tasks_by_id:
                # Covers a freshly exported file being re-imported before the
                # local state file has ever seen this key.
                match = tasks_by_id[explicit_id]
            else:
                match = tasks_by_bucket_title.get(normalize_key(bucket_name, title))

            if dry_run:
                if bucket_name not in buckets and bucket_name not in buckets_previewed:
                    emit(f"  Would create bucket '{bucket_name}'")
                    buckets_previewed.add(bucket_name)
                if match:
                    emit(f"[{i}] Would update '{title}' in bucket '{bucket_name}'")
                    summary.would_update += 1
                else:
                    emit(f"[{i}] Would create '{title}' in bucket '{bucket_name}'")
                    summary.would_create += 1
                continue

            if bucket_name not in buckets:
                emit(f"  Creating bucket '{bucket_name}'")
                new_id = create_bucket(client, plan_id, bucket_name)
                buckets[bucket_name] = new_id
                bucket_name_by_id[new_id] = bucket_name
            bucket_id = buckets[bucket_name]

            if match:
                emit(f"[{i}] Updating '{title}'...")
                action = update_task(client, match, bucket_id, row, user_cache)
                task_id = match["id"] if action != "failed" else ""
            else:
                emit(f"[{i}] Creating '{title}'...")
                task = create_task(client, plan_id, bucket_id, row, user_cache)
                action = "created" if task else "failed"
                task_id = task["id"] if task else ""
                if task:
                    # Keep the in-memory indexes current so a duplicate row
                    # later in the same file matches instead of creating twice.
                    tasks_by_id[task["id"]] = task
                    tasks_by_bucket_title[normalize_key(bucket_name, title)] = task

            if task_id:
                state.set(key, task_id, bucket_id)
                state.save()  # incremental: a crash mid-run keeps what we wrote

            summary.receipt_rows.append({
                "row": i,
                "title": title,
                "bucket": bucket_name,
                "action": action,
                "task_id": task_id,
                "task_url": task_url(plan_id, task_id) if task_id else "",
                "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            })

            if action == "created":
                summary.created += 1
            elif action == "updated":
                summary.updated += 1
            elif action == "unchanged":
                summary.unchanged += 1
            else:
                summary.failed += 1
    finally:
        # Write the receipt even if the run blew up part-way, so there's always
        # a record of what actually reached Planner.
        if not dry_run and summary.receipt_rows and receipt_path:
            summary.receipt_path = write_receipt(summary.receipt_rows, receipt_path)

    return summary


# ---------------------------------------------------------------------------
# Export driver
# ---------------------------------------------------------------------------

def resolve_user_label(client: GraphClient, user_id, cache):
    if user_id in cache:
        return cache[user_id]
    resp = client.get(f"{GRAPH}/users/{user_id}", params={"$select": "mail,userPrincipalName"})
    if resp.status_code in (401, 403):
        raise GraphPermissionError(
            f"Not permitted to look up users in the directory (HTTP {resp.status_code} "
            f"for GET /users/{user_id}). Delegated 'User.ReadBasic.All' is required to "
            f"write assignee emails into the export instead of raw object IDs. "
            f"Graph said: {resp.text}"
        )
    if resp.status_code != 200:
        cache[user_id] = user_id  # fall back to the raw ID rather than dropping it silently
        return cache[user_id]
    data = resp.json()
    label = data.get("mail") or data.get("userPrincipalName") or user_id
    cache[user_id] = label
    return label


def run_export(client: GraphClient, plan_id: str, *, emit: Callable[[str], None] = print) -> List[dict]:
    emit(f"Fetching buckets for plan {plan_id}...")
    buckets = get_buckets_by_id(client, plan_id)

    emit("Fetching tasks...")
    tasks = get_existing_tasks(client, plan_id)

    user_cache = {}
    rows = []
    for t in tasks:
        emit(f"  Reading details for '{t['title']}'...")
        details = get_task_details(client, t["id"])
        checklist_items = sorted(v["title"] for v in (details.get("checklist") or {}).values())
        assignees = [
            resolve_user_label(client, uid, user_cache)
            for uid in (t.get("assignments") or {}).keys()
        ]
        rows.append({
            "id": t["id"],
            "title": t["title"],
            "bucket": buckets.get(t["bucketId"], ""),
            "start_date": date_only(t.get("startDateTime")),
            "due_date": date_only(t.get("dueDateTime")),
            "priority": priority_to_label(t.get("priority")),
            "assigned_to": ";".join(assignees),
            "description": (details.get("description") or ""),
            "percent_complete": t.get("percentComplete", 0),
            "checklist": ";".join(checklist_items),
        })
    return rows


def write_export_csv(rows, out_path):
    parent = os.path.dirname(out_path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)
    return out_path
