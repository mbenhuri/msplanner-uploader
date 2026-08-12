#!/usr/bin/env python3
"""
Microsoft Planner bulk task importer — create + upsert (CLI).

Reads tasks from a CSV or JSON file and creates or updates them in a
Microsoft Planner plan via the Microsoft Graph API, using app-only (client
credentials) auth. Built to be run repeatedly against an evolving plan:
existing tasks are matched (by an explicit "id" column if present, otherwise
by bucket + title) so re-running with an edited file updates tasks in place
instead of creating duplicates.

All of the actual logic lives in planner_core.py, shared with the web app
(webapp/). This file is just argparse + app-only auth + stdout.

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
import sys

import planner_core as core
from cli_auth import build_token_provider, load_config


def main():
    parser = argparse.ArgumentParser(description="Bulk create/update tasks in Microsoft Planner.")
    parser.add_argument("input_file", help="Path to a .csv or .json tasks file")
    parser.add_argument("--plan-id", required=True, help="Target Planner plan ID")
    parser.add_argument("--dry-run", action="store_true", help="Preview without creating or updating anything")
    parser.add_argument(
        "--state-dir",
        default=core.DEFAULT_STATE_DIR,
        help=f"Directory holding the per-plan match state file (default: {core.DEFAULT_STATE_DIR})",
    )
    args = parser.parse_args()

    rows = core.load_rows(args.input_file)
    if not rows:
        sys.exit("No rows found in input file.")

    cfg = load_config()
    token_provider = build_token_provider(cfg)

    client_cls = core.ReadOnlyGraphClient if args.dry_run else core.GraphClient
    client = client_cls(token_provider, emit=print)

    try:
        summary = core.run_import(
            client,
            args.plan_id,
            rows,
            dry_run=args.dry_run,
            state_dir=args.state_dir,
            receipt_path=core.default_receipt_path(args.input_file),
            emit=print,
        )
    except core.GraphPermissionError as exc:
        sys.exit(f"\nAborted before writing anything: {exc}")

    if summary.dry_run:
        print(f"\nWould create {summary.would_create}, would update {summary.would_update}.")
    else:
        print(
            f"\nDone. Created {summary.created}, updated {summary.updated}, "
            f"unchanged {summary.unchanged}, failed {summary.failed}."
        )
        if summary.receipt_path:
            print(f"Receipt: {summary.receipt_path}")
        print(f"Match state: {summary.state_path} (keep this between runs)")


if __name__ == "__main__":
    main()
