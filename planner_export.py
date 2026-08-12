#!/usr/bin/env python3
"""
Export an existing Microsoft Planner plan to a CSV in the exact format
planner_import.py expects — so you can edit it and re-import safely, even
if you've lost your original source file.

Uses the same config.json / environment variables as planner_import.py.
The export logic itself lives in planner_core.py, shared with the web app.

Usage:
  python planner_export.py --plan-id <plan-id>
  python planner_export.py --plan-id <plan-id> --out my_plan.csv

The exported "id" column is populated with each task's real Planner task
ID. Keep that column intact when you edit the file — planner_import.py
matches on it directly even before planner_state/ has caught up, so it's
safe to rename titles between export and re-import.
"""

import argparse
import sys

import planner_core as core
from cli_auth import build_token_provider, load_config


def main():
    parser = argparse.ArgumentParser(
        description="Export a Planner plan to CSV, formatted for re-import with planner_import.py."
    )
    parser.add_argument("--plan-id", required=True, help="Planner plan ID to export")
    parser.add_argument("--out", default=None, help="Output CSV path (default: <plan-id>_export.csv)")
    args = parser.parse_args()

    cfg = load_config()
    token_provider = build_token_provider(cfg)

    # Export never writes to Planner, so it uses the read-only client too.
    client = core.ReadOnlyGraphClient(token_provider, emit=print)

    try:
        rows = core.run_export(client, args.plan_id, emit=print)
    except core.GraphPermissionError as exc:
        sys.exit(f"\nAborted: {exc}")

    out_path = core.write_export_csv(rows, args.out or f"{args.plan_id}_export.csv")
    print(f"\nExported {len(rows)} tasks to {out_path}")


if __name__ == "__main__":
    main()
