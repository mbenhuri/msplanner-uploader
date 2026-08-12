#!/usr/bin/env python3
"""
Microsoft Planner bulk task importer.

Reads tasks from a CSV or JSON file and creates them in a Microsoft Planner
plan via the Microsoft Graph API, using app-only (client credentials) auth.
No signed-in user or interactive login is required.

Setup (see README.md for full details):
  1. Register an app in Entra ID (Azure AD) admin center.
  2. Add these APPLICATION (not delegated) API permissions for Microsoft
     Graph, then grant admin consent:
       - Tasks.ReadWrite.All
       - Group.ReadWrite.All
       - User.Read.All
  3. Create a client secret for the app.
  4. Copy config.sample.json to config.json and fill in tenant_id,
     client_id, client_secret (or set them as environment variables
     PLANNER_TENANT_ID / PLANNER_CLIENT_ID / PLANNER_CLIENT_SECRET).
  5. pip install msal requests
  6. python planner_import.py tasks.csv --plan-id <your-plan-id> --dry-run
     python planner_import.py tasks.csv --plan-id <your-plan-id>

Input file columns / JSON keys (only "title" is required):
  title, bucket, due_date, start_date, priority, assigned_to, description,
  percent_complete, checklist

  - due_date / start_date: YYYY-MM-DD
  - priority: Urgent | Important | Medium | Low  (or a raw 0-10 number)
  - assigned_to: semicolon-separated emails, e.g. "a@co.com;b@co.com"
  - percent_complete: 0, 50, or 100
  - checklist: semicolon-separated item titles
  - bucket: created automatically if it doesn't already exist in the plan
"""

import argparse
import csv
import json
import os
import sys
import time
import uuid
from datetime import datetime

import requests
import msal

GRAPH = "https://graph.microsoft.com/v1.0"

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

    description = row.get("description")
    checklist = row.get("checklist")
    if description or checklist:
        set_task_details(client, task["id"], description, checklist)

    return task


def set_task_details(client, task_id, description, checklist):
    resp = client.get(f"{GRAPH}/planner/tasks/{task_id}/details")
    check(resp)
    etag = resp.headers.get("ETag")

    body = {}
    if description:
        body["description"] = description
    if checklist:
        items = [c.strip() for c in checklist.split(";") if c.strip()]
        body["checklist"] = {
            str(uuid.uuid4()): {
                "@odata.type": "#microsoft.graph.plannerChecklistItem",
                "title": item,
                "isChecked": False,
            }
            for item in items
        }
    if not body:
        return

    r = client.patch(
        f"{GRAPH}/planner/tasks/{task_id}/details",
        json=body,
        headers={"If-Match": etag},
    )
    if r.status_code not in (200, 204):
        print(f"  Warning: failed to set details for task {task_id}: {r.status_code} {r.text}")


def load_rows(path):
    if path.lower().endswith(".json"):
        with open(path) as f:
            data = json.load(f)
        return data if isinstance(data, list) else data.get("tasks", [])
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main():
    parser = argparse.ArgumentParser(description="Bulk-import tasks into Microsoft Planner.")
    parser.add_argument("input_file", help="Path to a .csv or .json tasks file")
    parser.add_argument("--plan-id", required=True, help="Target Planner plan ID")
    parser.add_argument("--dry-run", action="store_true", help="Preview without creating anything")
    args = parser.parse_args()

    rows = load_rows(args.input_file)
    if not rows:
        sys.exit("No rows found in input file.")

    cfg = load_config()
    token = get_token(cfg)
    client = GraphClient(token)

    print(f"Loading existing buckets for plan {args.plan_id}...")
    buckets = get_buckets(client, args.plan_id)
    user_cache = {}

    created, failed = 0, 0
    for i, raw_row in enumerate(rows, 1):
        row = {str(k).strip().lower(): v for k, v in raw_row.items()}
        title = (row.get("title") or "").strip()
        if not title:
            print(f"[{i}] Skipping row with no title.")
            continue
        row["title"] = title
        bucket_name = (row.get("bucket") or "To do").strip()

        if args.dry_run:
            print(f"[{i}] Would create '{title}' in bucket '{bucket_name}'")
            continue

        if bucket_name not in buckets:
            print(f"  Creating bucket '{bucket_name}'")
            buckets[bucket_name] = create_bucket(client, args.plan_id, bucket_name)

        print(f"[{i}] Creating '{title}'...")
        task = create_task(client, args.plan_id, buckets[bucket_name], row, user_cache)
        if task:
            created += 1
        else:
            failed += 1

    if not args.dry_run:
        print(f"\nDone. Created {created} tasks, {failed} failed.")


if __name__ == "__main__":
    main()