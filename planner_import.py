#!/usr/bin/env python3
"""
Microsoft Planner bulk task importer — create + upsert.

Reads tasks from a CSV or JSON file and creates or updates them in a
Microsoft Planner plan via the Microsoft Graph API, using app-only (client
credentials) auth. Built to be run repeatedly against an evolving plan:
existing tasks are matched (by an explicit "id" column if present, otherwise
by bucket + title) so re-running with an edited file updates tasks in place
instead of creating duplicates.

Setup (see README.md for full details):
  1. Register an app in Entra ID (Azure AD) admin center.
  2. Add these APPLICATION (not delegated) API permissions for Microsoft
     Graph, then grant admin consent:
       - Tasks.ReadWrite.All
       - User.Read.All
  3. Create a client secret for the app.
  4. Copy config.sample.json to config.json and fill in tenant_id,
     client_id, client_secret (or set them as environment variables
     PLANNER_TENANT_ID / PLANNER_CLIENT_ID / PLANNER_CLIENT_SECRET).
  5. pip install -r requirements.txt
  6. python planner_import.py tasks.csv --plan-id <your-plan-id> --dry-run
     python planner_import.py tasks.csv --plan-id <your-plan-id>

Input file columns / JSON keys (only "title" is required):
  id, title, bucket, due_date, start_date, priority, assigned_to,
  description, percent_complete, checklist

  - id: optional stable key for matching this row to a Planner task across
    runs, even if you later rename the title. If omitted, matching falls
    back to (bucket, title) — fine as long as you don't rename tasks.
  - due_date / start_date: YYYY-MM-DD
  - priority: Urgent | Important | Medium | Low  (or a raw 0-10 number)
  - assigned_to: semicolon-separated emails, e.g. "a@co.com;b@co.com"
  - percent_complete: 0, 50, or 100 — only applied when a task is first
    created. Re-running never overwrites progress you've since tracked
    directly in Planner.
  - checklist: semicolon-separated item titles
  - bucket: created automatically if it doesn't already exist in the plan

Re-running behavior (upsert):
  - Matched tasks are updated in place: title, bucket, dates, priority, and
    — only if the column is non-blank on that row — assignments,
    description, and checklist.
  - Leaving a field blank on a later run does NOT clear a previously-set
    value; it just leaves that field untouched. There's currently no way to
    explicitly clear a field via the input file — do that in Planner
    directly.
  - Matching uses a local state file (planner_state/<plan-id>.json), with a
    live bucket+title lookup as a fallback if that file is missing or out
    of sync, so you won't get duplicates even if the state file is lost.
  - Each real (non-dry-run) run writes a receipt file
    (<input-file>_receipt.csv) recording what happened to every row: the
    resulting task ID, a direct link, and whether it was created, updated,
    or left unchanged.
"""

import argparse
import csv
import json
import os
import sys
import time
import uuid
from datetime import datetime, timezone

import requests
import msal

GRAPH = "https://graph.microsoft.com/v1.0"
STATE_DIR = "planner_state"

PRIORITY_MAP = {
    "urgent": 1,
    "important": 3,
    "medium": 5,
    "low": 9,
}


def load_config():
    cfg = {}
    if os.path.exists("config.json"):
        with open("config.json") as f:
            cfg = json.load(f)
    cfg.setdefault("tenant_id", os.environ.get("PLANNER_TENANT_ID"))
    cfg.setdefault("client_id", os.environ.get("PLANNER_CLIENT_ID"))
    cfg.setdefault("client_secret", os.environ.get("PLANNER_CLIENT_SECRET"))
    missing = [k for k in ("tenant_id", "client_id", "client_secret") if not cfg.get(k)]
    if missing:
        sys.exit(
            f"Missing config values: {', '.join(missing)}. "
            f"Set them in config.json (see config.sample.json) or as environment variables."
        )
    return cfg


def get_token(cfg):
    app = msal.ConfidentialClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        client_credential=cfg["client_secret"],
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        sys.exit(f"Auth failed: {result.get('error_description', result)}")
    return result["access_token"]


class GraphClient:
    def __init__(self, token):
        self.session = requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        )

    def request(self, method, url, **kwargs):
        resp = None
        for attempt in range(5):
            resp = self.session.request(method, url, **kwargs)
            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                print(f"  Rate limited, waiting {wait}s...")
                time.sleep(wait)
                continue
            return resp
        return resp

    def get(self, url, **kw):
        return self.request("GET", url, **kw)

    def post(self, url, **kw):
        return self.request("POST", url, **kw)

    def patch(self, url, **kw):
        return self.request("PATCH", url, **kw)


def check(resp):
    """Raise with Graph's actual error body surfaced, instead of a bare status code."""
    if not resp.ok:
        print(f"  Graph error {resp.status_code} for {resp.request.method} {resp.url}")
        print(f"  Response body: {resp.text}")
    resp.raise_for_status()


# ---------------------------------------------------------------------------
# Buckets and tasks: read existing plan state
# ---------------------------------------------------------------------------

def get_buckets(client, plan_id):
    resp = client.get(f"{GRAPH}/planner/plans/{plan_id}/buckets")
    check(resp)
    return {b["name"]: b["id"] for b in resp.json()["value"]}


def create_bucket(client, plan_id, name, order_hint=" !"):
    resp = client.post(
        f"{GRAPH}/planner/buckets",
        json={"name": name, "planId": plan_id, "orderHint": order_hint},
    )
    check(resp)
    return resp.json()["id"]


def get_existing_tasks(client, plan_id):
    resp = client.get(f"{GRAPH}/planner/plans/{plan_id}/tasks")
    check(resp)
    return resp.json()["value"]


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


def state_path(plan_id):
    os.makedirs(STATE_DIR, exist_ok=True)
    return os.path.join(STATE_DIR, f"{plan_id}.json")


def load_state(plan_id):
    path = state_path(plan_id)
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return {}


def save_state(plan_id, state):
    with open(state_path(plan_id), "w") as f:
        json.dump(state, f, indent=2, sort_keys=True)


# ---------------------------------------------------------------------------
# Users / assignments
# ---------------------------------------------------------------------------

def resolve_user(client, email, cache):
    if email in cache:
        return cache[email]
    resp = client.get(f"{GRAPH}/users/{email}", params={"$select": "id"})
    if resp.status_code != 200:
        print(f"  Warning: could not resolve user '{email}' ({resp.status_code}); skipping assignment.")
        cache[email] = None
        return None
    uid = resp.json()["id"]
    cache[email] = uid
    return uid


def build_assignments(client, emails_str, cache):
    if not emails_str:
        return {}
    assignments = {}
    for email in [e.strip() for e in emails_str.split(";") if e.strip()]:
        uid = resolve_user(client, email, cache)
        if uid:
            assignments[uid] = {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !",
            }
    return assignments


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


def to_graph_date(value):
    if not value:
        return None
    value = str(value).strip()
    try:
        dt = datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None
    return dt.strftime("%Y-%m-%dT00:00:00Z")


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

def set_task_details(client, task_id, description, checklist):
    """Returns True if anything was actually changed."""
    description = (description or "").strip()
    checklist = (checklist or "").strip()
    if not description and not checklist:
        return False

    resp = client.get(f"{GRAPH}/planner/tasks/{task_id}/details")
    check(resp)
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
        print(f"  Warning: failed to set details for task {task_id}: {r.status_code} {r.text}")
        return False
    return True


# ---------------------------------------------------------------------------
# Create / update
# ---------------------------------------------------------------------------

def create_task(client, plan_id, bucket_id, row, user_cache):
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
        print(f"  Failed to create '{row['title']}': {resp.status_code} {resp.text}")
        return None
    task = resp.json()

    set_task_details(client, task["id"], row.get("description"), row.get("checklist"))
    return task


def update_task(client, task, bucket_id, row, user_cache):
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
            print(f"  Failed to update '{row['title']}': {resp.status_code} {resp.text}")
            return "failed"

    details_changed = set_task_details(client, task["id"], row.get("description"), row.get("checklist"))

    return "updated" if (task_body_changed or details_changed) else "unchanged"


# ---------------------------------------------------------------------------
# Input file loading
# ---------------------------------------------------------------------------

def load_rows(path):
    if path.lower().endswith(".json"):
        with open(path) as f:
            data = json.load(f)
        return data if isinstance(data, list) else data.get("tasks", [])
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


# ---------------------------------------------------------------------------
# Receipt log
# ---------------------------------------------------------------------------

def write_receipt(receipt_rows, input_path):
    stem = os.path.splitext(os.path.basename(input_path))[0]
    out_path = f"{stem}_receipt.csv"
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["row", "title", "bucket", "action", "task_id", "task_url", "timestamp"]
        )
        writer.writeheader()
        writer.writerows(receipt_rows)
    return out_path


def task_url(plan_id, task_id):
    return f"https://planner.cloud.microsoft/webui/plan/{plan_id}/view/board/task/{task_id}"


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Bulk create/update tasks in Microsoft Planner.")
    parser.add_argument("input_file", help="Path to a .csv or .json tasks file")
    parser.add_argument("--plan-id", required=True, help="Target Planner plan ID")
    parser.add_argument("--dry-run", action="store_true", help="Preview without creating or updating anything")
    args = parser.parse_args()

    rows = load_rows(args.input_file)
    if not rows:
        sys.exit("No rows found in input file.")

    cfg = load_config()
    token = get_token(cfg)
    client = GraphClient(token)

    print(f"Loading existing buckets for plan {args.plan_id}...")
    buckets = get_buckets(client, args.plan_id)
    bucket_name_by_id = {v: k for k, v in buckets.items()}

    print("Loading existing tasks...")
    existing_tasks = get_existing_tasks(client, args.plan_id)
    tasks_by_id = {t["id"]: t for t in existing_tasks}
    tasks_by_bucket_title = {}
    for t in existing_tasks:
        bname = bucket_name_by_id.get(t["bucketId"], "")
        tasks_by_bucket_title[normalize_key(bname, t["title"])] = t

    state = load_state(args.plan_id)
    user_cache = {}
    receipt_rows = []
    created = updated = unchanged = failed = 0
    would_create = would_update = 0
    buckets_previewed = set()

    for i, raw_row in enumerate(rows, 1):
        row = {str(k).strip().lower(): v for k, v in raw_row.items()}
        title = (row.get("title") or "").strip()
        if not title:
            print(f"[{i}] Skipping row with no title.")
            continue
        row["title"] = title
        bucket_name = (row.get("bucket") or "To do").strip()

        key = row_key(row)
        match = None
        state_entry = state.get(key)
        if state_entry and state_entry.get("task_id") in tasks_by_id:
            match = tasks_by_id[state_entry["task_id"]]
        else:
            match = tasks_by_bucket_title.get(normalize_key(bucket_name, title))

        if args.dry_run:
            if bucket_name not in buckets and bucket_name not in buckets_previewed:
                print(f"  Would create bucket '{bucket_name}'")
                buckets_previewed.add(bucket_name)
            if match:
                print(f"[{i}] Would update '{title}' in bucket '{bucket_name}'")
                would_update += 1
            else:
                print(f"[{i}] Would create '{title}' in bucket '{bucket_name}'")
                would_create += 1
            continue

        if bucket_name not in buckets:
            print(f"  Creating bucket '{bucket_name}'")
            new_id = create_bucket(client, args.plan_id, bucket_name)
            buckets[bucket_name] = new_id
            bucket_name_by_id[new_id] = bucket_name
        bucket_id = buckets[bucket_name]

        if match:
            print(f"[{i}] Updating '{title}'...")
            action = update_task(client, match, bucket_id, row, user_cache)
            task_id = match["id"] if action != "failed" else ""
        else:
            print(f"[{i}] Creating '{title}'...")
            task = create_task(client, args.plan_id, bucket_id, row, user_cache)
            action = "created" if task else "failed"
            task_id = task["id"] if task else ""

        if task_id:
            state[key] = {"task_id": task_id, "bucket_id": bucket_id}

        receipt_rows.append({
            "row": i,
            "title": title,
            "bucket": bucket_name,
            "action": action,
            "task_id": task_id,
            "task_url": task_url(args.plan_id, task_id) if task_id else "",
            "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        })

        if action == "created":
            created += 1
        elif action == "updated":
            updated += 1
        elif action == "unchanged":
            unchanged += 1
        else:
            failed += 1

    if args.dry_run:
        print(f"\nWould create {would_create}, would update {would_update}.")
    else:
        save_state(args.plan_id, state)
        receipt_path = write_receipt(receipt_rows, args.input_file)
        print(f"\nDone. Created {created}, updated {updated}, unchanged {unchanged}, failed {failed}.")
        print(f"Receipt: {receipt_path}")
        print(f"Match state: {state_path(args.plan_id)} (keep this between runs)")


if __name__ == "__main__":
    main()