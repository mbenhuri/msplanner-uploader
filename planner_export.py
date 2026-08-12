#!/usr/bin/env python3
"""
Export an existing Microsoft Planner plan to a CSV in the exact format
planner_import.py expects — so you can edit it and re-import safely, even
if you've lost your original source file.

Uses the same config.json / environment variables as planner_import.py.

Usage:
  python planner_export.py --plan-id <plan-id>
  python planner_export.py --plan-id <plan-id> --out my_plan.csv

The exported "id" column is populated with each task's real Planner task
ID. Keep that column intact when you edit the file — planner_import.py
matches on it directly even before planner_state/ has caught up, so it's
safe to rename titles between export and re-import.
"""

import argparse
import csv
import json
import os
import sys

import requests
import msal

GRAPH = "https://graph.microsoft.com/v1.0"

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

    def get(self, url, **kw):
        return self.session.get(url, **kw)


def check(resp):
    if not resp.ok:
        print(f"  Graph error {resp.status_code} for GET {resp.url}")
        print(f"  Response body: {resp.text}")
    resp.raise_for_status()


def get_buckets(client, plan_id):
    resp = client.get(f"{GRAPH}/planner/plans/{plan_id}/buckets")
    check(resp)
    return {b["id"]: b["name"] for b in resp.json()["value"]}


def get_tasks(client, plan_id):
    resp = client.get(f"{GRAPH}/planner/plans/{plan_id}/tasks")
    check(resp)
    return resp.json()["value"]


def get_details(client, task_id):
    resp = client.get(f"{GRAPH}/planner/tasks/{task_id}/details")
    check(resp)
    return resp.json()


def resolve_user_label(client, user_id, cache):
    if user_id in cache:
        return cache[user_id]
    resp = client.get(f"{GRAPH}/users/{user_id}", params={"$select": "mail,userPrincipalName"})
    if resp.status_code != 200:
        cache[user_id] = user_id  # fall back to the raw ID rather than dropping it silently
        return cache[user_id]
    data = resp.json()
    label = data.get("mail") or data.get("userPrincipalName") or user_id
    cache[user_id] = label
    return label


def date_only(iso_value):
    if not iso_value:
        return ""
    return iso_value.split("T")[0]


def export_rows(client, plan_id):
    print(f"Fetching buckets for plan {plan_id}...")
    buckets = get_buckets(client, plan_id)

    print("Fetching tasks...")
    tasks = get_tasks(client, plan_id)

    user_cache = {}
    rows = []
    for t in tasks:
        print(f"  Reading details for '{t['title']}'...")
        details = get_details(client, t["id"])
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


def main():
    parser = argparse.ArgumentParser(
        description="Export a Planner plan to CSV, formatted for re-import with planner_import.py."
    )
    parser.add_argument("--plan-id", required=True, help="Planner plan ID to export")
    parser.add_argument("--out", default=None, help="Output CSV path (default: <plan-id>_export.csv)")
    args = parser.parse_args()

    cfg = load_config()
    token = get_token(cfg)
    client = GraphClient(token)

    rows = export_rows(client, args.plan_id)

    out_path = args.out or f"{args.plan_id}_export.csv"
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nExported {len(rows)} tasks to {out_path}")


if __name__ == "__main__":
    main()