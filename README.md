# Planner Bulk Task Importer

Creates and updates Microsoft Planner tasks in bulk from a CSV or JSON file,
using the Microsoft Graph API directly. Two front ends, one set of logic:

| | Auth | Runs as | Use for |
|---|---|---|---|
| **Web app** (`webapp/`, `wsgi.py`) | Delegated — Authorization Code flow, you sign in | **You** | Day-to-day use. Scoped to plans you can already see. |
| **CLI** (`planner_import.py`, `planner_export.py`) | App-only — client credentials | A tenant-wide app identity | Unattended/scripted runs. |

All the Graph logic — matching, upsert, checklist diffing, priority/date
parsing, the match-state file, the receipt writer — lives in
**`planner_core.py`** and is shared. The only things that differ are *how a
token is obtained* and *how input and output are presented*.

```
planner_core.py     shared logic; takes a token provider, knows nothing about auth
cli_auth.py         app-only (client credentials) token provider, for the CLI
planner_import.py   CLI: argparse + stdout
planner_export.py   CLI: argparse + stdout
webapp/             Flask app: delegated auth + a form + streamed results
  config.py         all environment-specific settings (nothing hardcodes localhost)
  auth.py           MSAL confidential client, auth code flow, token cache
  views.py          routes; runs stream progress as they go
wsgi.py             entry point for both `python wsgi.py` and gunicorn
tests/              in-memory fake of Graph; no tenant needed
```

Designed for iterating: generate a plan, review the file, import it, then edit
the file and re-run — matching tasks are **updated in place** instead of
duplicated.

---

# Part 1 — The web app (delegated auth)

The token represents **you**, scoped to what you can already see in Planner,
rather than a tenant-wide app identity. A leaked app-only secret grants
immediate tenant-wide Graph access; a leaked confidential-client secret alone
grants nothing without also getting through your actual sign-in and any
Conditional Access/MFA policies attached to it.

Uses **Authorization Code flow** (the normal "sign in with Microsoft" browser
redirect). Device code flow is deliberately not used — it is a known phishing
vector (Storm-2372) and is disabled tenant-wide here.

> **Read this before you start:** under delegated permissions Planner
> authorizes against your **membership of the plan's Microsoft 365 group**.
> Being a tenant admin is *not* enough on its own. A plan you can see in the
> Planner UI but are not a member of will return 403 where the app-only CLI
> worked. That is the intended scoping, not a bug.

## 1. Register the web app in Entra

This is a **new, separate** registration from the CLI's app-only one. Keeping
them apart means the two auth models can be granted, audited, and revoked
independently, and a compromise of one doesn't imply the other.

In the [Entra admin center](https://entra.microsoft.com) → **App registrations**
→ **New registration**:

1. **Name**: something distinguishable from the CLI app, e.g.
   `planner-uploader-web`.
2. **Supported account types**: *Accounts in this organizational directory only
   (single tenant)*.
3. **Redirect URI**: set the platform dropdown to **Web** — not "Public
   client/native (mobile & desktop)", which is the type device code flow needs
   — and enter:

   ```
   http://localhost:5000/getAToken
   ```

   Entra allows plain `http` for loopback addresses. Port 5000 is free on
   Linux/WSL2 (the macOS AirPlay conflict doesn't apply). If you change the
   port, change it in **both** places — this redirect URI and
   `PLANNER_WEB_BASE_URL` — they must match exactly. Note that `localhost` and
   `127.0.0.1` are *different* URIs to Entra; use `localhost`.

4. Click **Register**, then copy the **Application (client) ID** and
   **Directory (tenant) ID** from the Overview page.

5. **Certificates & secrets** → **New client secret**. Copy the secret
   **Value** immediately — it's shown only once.

   A certificate is the stronger option and worth moving to before this runs
   anywhere but your own machine; a secret is fine to start.

6. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions**. Add:

   - `Tasks.ReadWrite` — plans, buckets, tasks, task details
   - `User.ReadBasic.All` — **required** to resolve assignee emails to user IDs

   `User.Read` is added by default; leave it. MSAL requests `openid`,
   `profile` and `offline_access` automatically — you don't add those.

   **Why `User.ReadBasic.All` is not optional:** the importer turns the
   `assigned_to` column into user IDs via `GET /users/{email}`, and the
   exporter turns IDs back into emails the same way. Delegated `User.Read`
   grants only `/me`. Without `User.ReadBasic.All` every assignment silently
   fails to resolve. The app now detects that case up front and refuses to run
   rather than importing a plan with no assignees — but the fix is to grant the
   permission.

   **`Group.Read.All` is deliberately not requested.** Listing buckets and
   tasks authorizes off group membership, not a group scope. If you ever do hit
   a 403 on a bucket call, read the Graph error body the app prints before
   adding scopes speculatively.

7. Click **Grant admin consent for &lt;your tenant&gt;**. User self-consent is
   disabled tenant-wide, so this must happen before the first sign-in or the
   sign-in will fail with a consent error.

8. **Authentication** blade → under *Advanced settings*, confirm **Allow public
   client flows** is set to **No**. This is a confidential client; leaving that
   on would re-enable exactly the public-client flows you turned off.

## 2. Configure it

Environment variables (recommended) — note the `PLANNER_WEB_` prefix, distinct
from the CLI's `PLANNER_` variables so the two registrations can't be crossed:

```bash
export PLANNER_WEB_TENANT_ID="..."
export PLANNER_WEB_CLIENT_ID="..."
export PLANNER_WEB_CLIENT_SECRET="..."
export PLANNER_WEB_SECRET_KEY="$(python -c 'import secrets; print(secrets.token_hex(32))')"
export PLANNER_WEB_BASE_URL="http://localhost:5000"
```

Or copy `config.web.sample.json` to `config.web.json` and fill it in — that
file is git-ignored.

`SECRET_KEY` signs the Flask session cookie. Set it to a fixed value rather
than regenerating it per start, or every restart signs you out.

Optional:

| Variable | Default | Notes |
|---|---|---|
| `PLANNER_WEB_BASE_URL` | `http://localhost:5000` | Must match the registered redirect URI |
| `PLANNER_WEB_DATA_DIR` | `./webapp_data` | Token cache, match state, run outputs |
| `PLANNER_WEB_HOST` | `127.0.0.1` | Bind address for the dev server |
| `PLANNER_WEB_DEBUG` | unset | `1` enables the Flask reloader. Don't use with real credentials |

## 3. Run it

```bash
source venv/bin/activate
pip install -r requirements.txt
python wsgi.py
```

Open <http://localhost:5000>. You'll be redirected to Entra (and on to your
federated IdP if your account is federated — nothing in this app is
IdP-specific), then back to a page with two forms:

- **Import** — plan ID, a CSV/JSON file, and a **Dry run** checkbox (checked by
  default). Progress streams row by row exactly as the CLI prints it, and a
  real run ends with a **Download receipt CSV** link.
- **Export** — plan ID; ends with a **Download plan CSV** link.

Long runs stream their output rather than blocking, so the connection keeps
producing bytes and neither the browser nor a proxy times out mid-import.

## 4. Security notes for the web app

- **The token cache is now the most sensitive file on disk.** MSAL's cache is
  serialized to `webapp_data/token_cache.bin` (mode `0600`, git-ignored). It
  holds a refresh token that acts as you, and redeeming it does *not* re-run
  interactive MFA — Conditional Access is re-evaluated at refresh, but that's a
  policy check, not a fresh sign-in. Treat that file the way you'd treat a
  password.
- **Single user only.** A file-backed cache and Flask's cookie session are fine
  for one admin on one machine. See "Moving to Azure App Service" below before
  anyone else uses it.
- **CSRF**: the forms have no CSRF token; the session cookie is `SameSite=Lax`,
  which stops a cross-site POST from carrying it. That's adequate for a
  single-user localhost tool and should become a real CSRF token if this is
  ever exposed.

## 5. Moving to Azure App Service (not done yet)

Structured so this is configuration, not a rewrite:

1. Add a second redirect URI to the same app registration:
   `https://<site>.azurewebsites.net/getAToken`, and set
   `PLANNER_WEB_BASE_URL` to that https URL. The session cookie automatically
   becomes `Secure` when the base URL is https.
2. Set `PLANNER_WEB_DATA_DIR=/home/data` — `/home` is the only path App Service
   persists across restarts.
3. Start with `gunicorn --bind 0.0.0.0:8000 --timeout 600 wsgi:app`. The long
   timeout matters because runs stream for as long as the import takes. App
   Service also kills any request producing no bytes for 230 seconds, which is
   why runs stream progress instead of blocking.
4. **Before more than one person uses it**: replace the file cache with
   server-side sessions (`Flask-Session` on Azure Cache for Redis), key the MSAL
   cache per user object ID, key `planner_state/` per user as well as per plan,
   and move the client secret to Key Vault — or drop the secret entirely in
   favour of a certificate or managed identity.

---

# Part 2 — The CLI (app-only auth)

Unchanged in behavior; it now imports its logic from `planner_core.py`.

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
  numbers) — that's the `--plan-id` value, and what the web app's form asks for.
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

Using a venv keeps `msal`/`requests`/`flask` isolated from your system Python and
any other projects.

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

---

# Reference

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

Re-running with an edited file updates existing tasks instead of creating
duplicates. This is identical in the CLI and the web app.

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
- **Unresolvable assignees don't wipe assignments**: if a row's `assigned_to`
  names people but *none* of them can be resolved (a typo, someone who has
  left), the existing assignments are left alone and a warning is printed,
  rather than the task being silently unassigned. If at least one resolves,
  assignments sync normally.
- **Matching survives a lost state file**: matches are tracked in
  `planner_state/<plan-id>.json` (`webapp_data/planner_state/` for the web
  app), created automatically and saved after every row, so a run that dies
  part-way still records what it wrote. If that file is missing or out of
  date, the tool falls back to a live bucket+title lookup against the plan, so
  you won't get duplicates even without it — but keep the file around between
  runs regardless, since it's the only way `id`-based matching survives a title
  rename. Don't commit it.
- **Receipt log**: every real (non-dry-run) run writes a receipt — one row per
  task with its ID, a direct link, and whether it was `created`, `updated`, or
  `unchanged`. The CLI writes `<input-file>_receipt.csv`; the web app offers it
  as a download. It's written even if the run fails part-way.

## Tests

The Graph endpoints are faked in memory (`tests/fake_graph.py`) rather than
hit for real, so the suite needs no tenant and no credentials:

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -q
```

Coverage is weighted toward what would break silently in a refactor: checklist
diffing, upsert matching (state file / explicit id / bucket+title fallback),
pagination, and the safety properties — dry run issues zero non-GET requests,
`percent_complete` is create-only, blank fields never clear. The web tests
drive the streaming import and export end to end against the same fake, with
auth stubbed. Auth and Flask plumbing itself is verified by hand against a
sandbox tenant.

## Notes and limitations

- Assignee emails must match existing Microsoft 365 accounts (resolved via
  `/users/{email}`); unmatched emails are skipped with a warning, not fatal.
- Nothing is batched — one API call per task (plus one for details if
  description/checklist are set, plus the initial bucket and task listing). For
  very large imports (500+ tasks) this takes a few minutes; the tool backs off
  and retries on HTTP 429, and renews its access token if a run outlives one.
- Buckets are matched by exact name; renaming buckets in Planner after a
  first run will cause new ones to be created rather than reused. Tasks
  moved to a renamed bucket via the CSV will follow correctly, though —
  it's only the bucket *name-matching* that's exact.
- The CLI uses application permissions, which apply tenant-wide (any plan the
  app is pointed at) rather than being scoped to one specific plan — that's a
  current limitation of the Planner API. The web app's delegated auth is the
  answer to that: it can only reach plans you're a member of.
