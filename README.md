# Planner Bulk Task Importer

Creates and updates Microsoft Planner tasks in bulk from a CSV or JSON file,
using the Microsoft Graph API directly (no Claude MCP connector involved).
Runs unattended with app-only (client credentials) authentication — no
interactive sign-in needed.

Designed for iterating with Claude: generate a plan, review the file, import
it, then edit the file and re-run — matching tasks are **updated in place**
instead of duplicated.

## 1. Register an Azure AD app

In the [Entra admin center](https://entra.microsoft.com) → **App registrations** → **New registration**:

1. Give it a name (e.g. `planner-import-script`), leave the rest as default, and register.
2. Note the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
3. Go to **Certificates & secrets** → **New client secret**. Copy the secret **value**
   immediately (it's only shown once).
4. Go to **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Application permissions**, and add:
   - `Tasks.ReadWrite.All`
   - `User.Read.All`

   `Group.ReadWrite.All` is **not** required for this script — it's only
   needed if you extend it to create brand-new plans, which this script
   doesn't do. Verified against the current per-endpoint Graph docs (list
   buckets, create bucket, create task, get/update task details) — all
   support application permissions with `Tasks.ReadWrite.All` alone. Note
   that Microsoft's *general* Planner permissions overview page currently
   claims application permissions aren't supported at all for Planner; that
   appears to be stale relative to the endpoint-specific docs, but confirm
   against a sandbox tenant/plan before treating this as final — real
   behavior has occasionally lagged the docs here in the past.
5. Click **Grant admin consent** (you'll need an admin, or ask one to click it —
   this is usually the step your IT team needs to approve).

## 2. Find your Plan ID

Open the plan in Planner on the web. The URL looks like this:

```
https://planner.cloud.microsoft/webui/plan/ptWtAvOsBUu6goF2P5Swk2UABmEl/view/board?tid=564f1b83-ec27-441b-8e1d-6b46fe305086
```

- The segment right after `/plan/` is your **Plan ID** (28 characters, letters and
  numbers) — that's the `--plan-id` value.
- The `?tid=` value is your **tenant ID**, not the plan ID — it should match the
  `tenant_id` you set in `config.json`.

(Older plans may still show the previous `tasks.office.com/.../plantaskboard?planId=...`
format, where the `planId=` query parameter is the value you want — same idea,
different URL shape.)

Alternatively, with your app credentials already set up, you can list plans
for a group with:

```
GET https://graph.microsoft.com/v1.0/groups/{group-id}/planner/plans
```

## 3. Configure credentials

Copy `config.sample.json` to `config.json` and fill in your values, **or**
set environment variables instead (useful if you don't want secrets in a
file):

```
export PLANNER_TENANT_ID="..."
export PLANNER_CLIENT_ID="..."
export PLANNER_CLIENT_SECRET="..."
```

`config.json` is git-ignored territory — don't commit it if this goes in a repo.

## 4. Set up a virtual environment and install dependencies

Using a venv keeps `msal`/`requests` isolated from your system Python and any
other projects.

```
python3 -m venv venv
```

Activate it (do this every time you open a new terminal to work on this):

```
source venv/bin/activate        # macOS/Linux
venv\Scripts\activate           # Windows (cmd/PowerShell)
```

Install dependencies from `requirements.txt`:

```
pip install -r requirements.txt
```

`venv/` is already covered by `.gitignore` — don't commit it.

## 5. Run it

Preview first (creates nothing, just shows what would happen):

```
python planner_import.py sample_tasks.csv --plan-id YOUR_PLAN_ID --dry-run
```

Then run for real:

```
python planner_import.py sample_tasks.csv --plan-id YOUR_PLAN_ID
```

JSON input works the same way — either a top-level array of task objects, or
`{"tasks": [...]}`, using the same field names as the CSV columns below.

## Lost your source file? Export the plan back to CSV

If you've lost the CSV/JSON you originally imported (or just want to pull a
plan's current state into an editable file), `planner_export.py` reads an
existing plan and writes it out in the exact format `planner_import.py`
expects:

```
python planner_export.py --plan-id YOUR_PLAN_ID --out recovered.csv
```

Edit `recovered.csv` and re-import it like any other file:

```
python planner_import.py recovered.csv --plan-id YOUR_PLAN_ID --dry-run
python planner_import.py recovered.csv --plan-id YOUR_PLAN_ID
```

The exported `id` column holds each task's real Planner task ID, so
matching works correctly on re-import even with no `planner_state/` file
present yet — including if you rename a title in the recovered file before
re-importing. Keep that `id` column intact when editing.

Notes on the export:
- Priority is written back as a label (`Urgent`/`Important`/`Medium`/`Low`),
  based on Microsoft's documented priority buckets (0–1, 2–4, 5–7, 8–10).
- Assignees are resolved back to email/UPN where possible; if a user can't
  be looked up, their raw ID is written instead so the assignment isn't
  silently dropped.
- Checklist items are written in title order (alphabetical), not the
  original manual ordering from Planner — item content and checked state
  round-trip correctly, ordering doesn't.

## Input file fields

| Field              | Required | Notes |
|--------------------|----------|-------|
| `id`               | No       | Optional stable key for matching this row to a Planner task across runs — see [Re-running (upsert)](#re-running-upsert) below. Populated automatically by `planner_export.py`. |
| `title`            | Yes      | Task name |
| `bucket`            | No       | Column/bucket name; created automatically if missing |
| `start_date`       | No       | `YYYY-MM-DD` |
| `due_date`         | No       | `YYYY-MM-DD` |
| `priority`         | No       | `Urgent`, `Important`, `Medium`, `Low`, or a raw 0–10 number |
| `assigned_to`      | No       | Semicolon-separated emails, e.g. `a@co.com;b@co.com` |
| `description`      | No       | Plain text task notes |
| `percent_complete` | No       | `0`, `50`, or `100`. Only applied when a task is first created — re-running never overwrites progress tracked directly in Planner |
| `checklist`        | No       | Semicolon-separated checklist item titles |

## Re-running (upsert)

Re-running the script with an edited file updates existing tasks instead of
creating duplicates:

- **Matching**: each row is matched to an existing task by its `id` column
  if present, otherwise by `(bucket, title)`. If your `id` value is a real
  Planner task ID (as written by `planner_export.py`), matching works even
  before `planner_state/` has an entry for it — so exported files re-import
  safely from a completely clean checkout. A custom, non-task-ID `id` value
  (e.g. `kickoff-meeting`) works too, but only once `planner_state/` has
  recorded it from a prior run.
- **What gets synced on update**: title, bucket, start/due dates, and
  priority always sync to match the file. Assignments, description, and
  checklist only sync when that row's column is non-blank — leaving a field
  blank on a later run does **not** clear a previously-set value, it just
  leaves it alone. There's currently no way to explicitly clear a field via
  the file; do that in Planner directly.
- **`percent_complete` is create-only** — re-running never resets progress
  you've since tracked in Planner.
- **Checklist merging**: matching items (by title, case-insensitive) keep
  their checked state; items no longer listed are removed; new items are
  added.
- **Matching survives a lost state file**: matches are tracked in
  `planner_state/<plan-id>.json`, created automatically. If that file is
  missing or out of date, the script falls back to a live bucket+title
  lookup against the plan, so you won't get duplicates even without it —
  but keep the file around between runs regardless, since it's the only
  way `id`-based matching survives a title rename. Don't commit it (add
  `planner_state/` to `.gitignore`).
- **Receipt log**: every real (non-`--dry-run`) run writes
  `<input-file>_receipt.csv` — one row per task with its ID, a direct link,
  and whether it was `created`, `updated`, or `unchanged`.

## Notes and limitations

- Assignee emails must match existing Microsoft 365 accounts (resolved via
  `/users/{email}`); unmatched emails are skipped with a warning, not fatal.
- The script batches nothing — it makes one API call per task (plus one for
  details if description/checklist are set, plus one for the initial bucket
  and task listing). For very large imports (500+ tasks) this will take a
  few minutes; the script automatically backs off and retries on HTTP 429
  rate-limit responses.
- Buckets are matched by exact name; renaming buckets in Planner after a
  first run will cause new ones to be created rather than reused. Tasks
  moved to a renamed bucket via the CSV will follow correctly, though —
  it's only the bucket *name-matching* that's exact.
- This uses application permissions, which apply tenant-wide (any plan the
  app is pointed at) rather than being scoped to one specific plan — that's a
  current limitation of the Planner API, not something this script can avoid.