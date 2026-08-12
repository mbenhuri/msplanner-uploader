"""
Routes: sign-in, the single-page form, and the two streaming run endpoints.

Why the runs stream
-------------------
A few hundred rows is one to three Graph calls each plus throttling backoff, so
an import can easily run for minutes. A plain blocking POST would look hung
locally and would hard-fail on Azure App Service, which kills any request that
sends nothing for 230 seconds. So each run is handed to a worker thread that
pushes progress lines onto a queue, and the response streams those lines as
they arrive — the connection keeps producing bytes, and the user sees the same
per-row output the CLI prints. The receipt/export CSV is written server-side
under RUNS_DIR and offered as a download link at the end of the stream.

The worker thread has no request context, so everything it needs (file bytes,
plan id, account id) is passed in explicitly; it never touches `session`.
"""

import os
import queue
import re
import threading
import uuid
from functools import wraps

import requests
from flask import (
    Blueprint, Response, current_app, redirect, render_template, request,
    send_file, session, stream_with_context, url_for,
)
from markupsafe import escape
from werkzeug.utils import secure_filename

import planner_core as core
from .auth import AuthError

bp = Blueprint("main", __name__)

RUN_ID_RE = re.compile(r"^[0-9a-f]{32}$")
_SENTINEL = object()


def auth_manager():
    return current_app.extensions["planner_auth"]


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        account_id = session.get("home_account_id")
        if not auth_manager().has_valid_session(account_id):
            session.clear()
            return redirect(url_for("main.login"))
        return view(*args, **kwargs)

    return wrapped


# ---------------------------------------------------------------------------
# Auth routes
# ---------------------------------------------------------------------------

@bp.route("/login")
def login():
    flow = auth_manager().build_auth_code_flow()
    session["auth_flow"] = flow
    return redirect(flow["auth_uri"])


@bp.route("/getAToken")
def authorized():
    flow = session.pop("auth_flow", None)
    if not flow:
        return redirect(url_for("main.login"))
    try:
        result = auth_manager().complete_auth_code_flow(flow, request.args)
    except ValueError:
        # State mismatch or a replayed/stale callback.
        return redirect(url_for("main.login"))

    if "error" in result:
        return render_template(
            "error.html",
            title="Sign-in failed",
            message=result.get("error_description") or result["error"],
            detail=(
                "If this says consent is required, the app registration still needs admin "
                "consent granted for its delegated permissions — see the README."
            ),
        ), 401

    claims = result.get("id_token_claims") or {}
    session["home_account_id"] = auth_manager().home_account_id_from(result)
    session["user"] = {
        "name": claims.get("name"),
        "username": claims.get("preferred_username"),
    }
    return redirect(url_for("main.index"))


@bp.route("/logout")
def logout():
    account_id = session.get("home_account_id")
    if account_id:
        auth_manager().forget(account_id)
    session.clear()
    authority = current_app.config["AUTHORITY"]
    base = current_app.config["BASE_URL"]
    return redirect(f"{authority}/oauth2/v2.0/logout?post_logout_redirect_uri={base}")


# ---------------------------------------------------------------------------
# Main page
# ---------------------------------------------------------------------------

@bp.route("/")
@login_required
def index():
    return render_template("index.html", user=session.get("user") or {})


# ---------------------------------------------------------------------------
# Streaming plumbing
# ---------------------------------------------------------------------------

def _run_streamed(job, heading, subtitle):
    """
    job(emit) -> (summary_html_context). Runs on a worker thread; returns a
    streaming HTML response.
    """
    q: "queue.Queue" = queue.Queue()

    def emit(message):
        q.put(("log", str(message)))

    def worker():
        try:
            q.put(("done", job(emit)))
        except core.GraphPermissionError as exc:
            q.put(("error", ("Missing permission — nothing was written.", str(exc))))
        except AuthError as exc:
            q.put(("error", ("Sign-in expired.", str(exc))))
        except core.DryRunViolation as exc:
            q.put(("error", ("Dry run tried to write. This is a bug; nothing was written.", str(exc))))
        except requests.HTTPError as exc:
            resp = exc.response
            body = resp.text if resp is not None else str(exc)
            status = resp.status_code if resp is not None else "?"
            hint = ""
            if status == 403:
                hint = (
                    " Under delegated permissions, Planner authorizes off your membership "
                    "of the plan's Microsoft 365 group — being a tenant admin is not enough. "
                    "Check you are a member of this plan's group."
                )
            q.put(("error", (f"Graph returned {status}.{hint}", body)))
        except Exception as exc:  # noqa: BLE001 - surfaced to the page, not swallowed
            q.put(("error", (type(exc).__name__, str(exc))))
        finally:
            q.put(_SENTINEL)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    def generate():
        yield render_template("stream_head.html", heading=heading, subtitle=subtitle)
        result, error = None, None
        while True:
            item = q.get()
            if item is _SENTINEL:
                break
            kind, payload = item
            if kind == "log":
                yield f'<div class="line">{escape(payload)}</div>\n'
            elif kind == "done":
                result = payload
            elif kind == "error":
                error = payload
        yield render_template("stream_tail.html", result=result, error=error)

    resp = Response(stream_with_context(generate()), mimetype="text/html")
    resp.headers["X-Accel-Buffering"] = "no"  # don't let a reverse proxy buffer the stream
    resp.headers["Cache-Control"] = "no-cache"
    return resp


def _new_run_dir():
    run_id = uuid.uuid4().hex
    path = os.path.join(current_app.config["RUNS_DIR"], run_id)
    os.makedirs(path, exist_ok=True)
    return run_id, path


# ---------------------------------------------------------------------------
# Import / export
# ---------------------------------------------------------------------------

@bp.route("/import", methods=["POST"])
@login_required
def do_import():
    plan_id = (request.form.get("plan_id") or "").strip()
    dry_run = request.form.get("dry_run") == "on"
    upload = request.files.get("file")

    if not plan_id:
        return render_template("error.html", title="Missing plan ID",
                               message="Enter the plan ID to import into."), 400
    if not upload or not upload.filename:
        return render_template("error.html", title="No file",
                               message="Choose a CSV or JSON file to import."), 400

    filename = secure_filename(upload.filename)
    data = upload.read()

    # Parse before streaming so a bad file is a normal error page with a real
    # status code, not an error buried at the bottom of a 200 response.
    try:
        rows = core.load_rows(data, filename)
    except Exception as exc:  # noqa: BLE001
        return render_template("error.html", title="Could not read that file",
                               message=str(exc)), 400
    if not rows:
        return render_template("error.html", title="Empty file",
                               message="No rows found in the uploaded file."), 400

    run_id, run_dir = _new_run_dir()
    receipt_path = os.path.join(run_dir, "receipt.csv")
    state_dir = current_app.config["STATE_DIR"]
    # Captured now, on the request thread — the worker has no session access.
    provider = auth_manager().token_provider(session["home_account_id"])
    download_url = url_for("main.download", run_id=run_id, kind="receipt")

    def job(emit):
        client_cls = core.ReadOnlyGraphClient if dry_run else core.GraphClient
        client = client_cls(provider, emit=emit)
        summary = core.run_import(
            client, plan_id, rows,
            dry_run=dry_run, state_dir=state_dir,
            receipt_path=receipt_path, emit=emit,
        )
        if summary.dry_run:
            headline = f"Dry run: would create {summary.would_create}, would update {summary.would_update}."
            note = "Nothing was written to Planner."
        else:
            headline = (
                f"Created {summary.created}, updated {summary.updated}, "
                f"unchanged {summary.unchanged}, failed {summary.failed}."
            )
            note = f"Match state saved to {summary.state_path}"
        return {
            "headline": headline,
            "note": note,
            "download_url": download_url if summary.receipt_path else None,
            "download_label": "Download receipt CSV",
        }

    label = "Dry run" if dry_run else "Import"
    return _run_streamed(job, f"{label}: {filename}", f"Plan {plan_id} · {len(rows)} rows")


@bp.route("/export", methods=["POST"])
@login_required
def do_export():
    plan_id = (request.form.get("plan_id") or "").strip()
    if not plan_id:
        return render_template("error.html", title="Missing plan ID",
                               message="Enter the plan ID to export."), 400

    run_id, run_dir = _new_run_dir()
    export_path = os.path.join(run_dir, "export.csv")
    provider = auth_manager().token_provider(session["home_account_id"])
    download_url = url_for("main.download", run_id=run_id, kind="export")

    def job(emit):
        # Export never writes, so it uses the read-only client too.
        client = core.ReadOnlyGraphClient(provider, emit=emit)
        rows = core.run_export(client, plan_id, emit=emit)
        core.write_export_csv(rows, export_path)
        return {
            "headline": f"Exported {len(rows)} tasks.",
            "note": "The id column holds real Planner task IDs — keep it intact when editing.",
            "download_url": download_url,
            "download_label": "Download plan CSV",
        }

    return _run_streamed(job, f"Export: {plan_id}", "Reading buckets, tasks and task details")


@bp.route("/download/<run_id>/<kind>")
@login_required
def download(run_id, kind):
    if not RUN_ID_RE.match(run_id or "") or kind not in ("receipt", "export"):
        return render_template("error.html", title="Not found",
                               message="No such run."), 404
    filename = "receipt.csv" if kind == "receipt" else "export.csv"
    path = os.path.join(current_app.config["RUNS_DIR"], run_id, filename)
    if not os.path.exists(path):
        return render_template("error.html", title="Not found",
                               message="That file is no longer available."), 404
    return send_file(path, mimetype="text/csv", as_attachment=True, download_name=filename)
