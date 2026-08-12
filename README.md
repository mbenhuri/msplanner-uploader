# Planner Bulk Task Importer
 
Creates Microsoft Planner tasks in bulk from a CSV or JSON file, using the
Microsoft Graph API directly (no Claude MCP connector involved). Runs
unattended with app-only (client credentials) authentication — no interactive
sign-in needed.
 
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
 
Open the plan in Planner on the web. The URL contains the plan ID, e.g.:
 
```
https://tasks.office.com/.../Home/Planner/#/plantaskboard?planId=AAAbbbCCCddd123456789
```
 
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
 
## 4. Install dependencies
 
```
pip install msal requests
```
 
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
 
## Input file fields
 
| Field              | Required | Notes |
|--------------------|----------|-------|
| `title`            | Yes      | Task name |
| `bucket`           | No       | Column/bucket name; created automatically if missing |
| `start_date`       | No       | `YYYY-MM-DD` |
| `due_date`         | No       | `YYYY-MM-DD` |
| `priority`         | No       | `Urgent`, `Important`, `Medium`, `Low`, or a raw 0–10 number |
| `assigned_to`      | No       | Semicolon-separated emails, e.g. `a@co.com;b@co.com` |
| `description`      | No       | Plain text task notes |
| `percent_complete` | No       | `0`, `50`, or `100` |
| `checklist`        | No       | Semicolon-separated checklist item titles |
 
## Notes and limitations
 
- Assignee emails must match existing Microsoft 365 accounts (resolved via
  `/users/{email}`); unmatched emails are skipped with a warning, not fatal.
- The script batches nothing — it makes one API call per task (plus one for
  details if description/checklist are set). For very large imports (500+
  tasks) this will take a few minutes; the script automatically backs off and
  retries on HTTP 429 rate-limit responses.
- Buckets are matched by exact name; renaming buckets in Planner after a
  first run will cause new ones to be created rather than reused.
- This uses application permissions, which apply tenant-wide (any plan the
  app is pointed at) rather than being scoped to one specific plan — that's a
  current limitation of the Planner API, not something this script can avoid.
 