"""
Web app plumbing tests.

Auth is stubbed (the real MSAL calls need a tenant — that part is verified by
hand against the sandbox), but the streaming run path, the safety-relevant
form handling, and the download routes are exercised for real against the
in-memory Graph fake.
"""

import io
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.fake_graph import FakeGraph  # noqa: E402
from webapp import create_app  # noqa: E402


class StubAuth:
    """Stands in for AuthManager without touching Entra."""

    def __init__(self):
        self.forgotten = []

    def has_valid_session(self, account_id):
        return account_id == "acct-1"

    def token_provider(self, account_id):
        return lambda force_refresh=False: "tok"

    def forget(self, account_id):
        self.forgotten.append(account_id)


@pytest.fixture
def app(tmp_path):
    app = create_app({
        "TENANT_ID": "tenant", "CLIENT_ID": "client", "CLIENT_SECRET": "secret",
        "SECRET_KEY": "test-key", "DATA_DIR": str(tmp_path / "data"),
        "BASE_URL": "http://localhost:5000",
    })
    app.config["TESTING"] = True
    app.extensions["planner_auth"] = StubAuth()
    return app


@pytest.fixture
def client(app):
    c = app.test_client()
    with c.session_transaction() as sess:
        sess["home_account_id"] = "acct-1"
        sess["user"] = {"name": "Admin", "username": "admin@example.com"}
    return c


@pytest.fixture
def graph(monkeypatch):
    return FakeGraph().install(monkeypatch)


# ---------------------------------------------------------------------------
# Config / auth gating
# ---------------------------------------------------------------------------

def test_redirect_uri_is_derived_from_base_url(app):
    assert app.config["REDIRECT_URI"] == "http://localhost:5000/getAToken"
    assert app.config["SESSION_COOKIE_SECURE"] is False


def test_https_base_url_marks_the_cookie_secure(tmp_path):
    app = create_app({
        "TENANT_ID": "t", "CLIENT_ID": "c", "CLIENT_SECRET": "s", "SECRET_KEY": "k",
        "DATA_DIR": str(tmp_path / "d"), "BASE_URL": "https://planner.example.net",
    })
    assert app.config["REDIRECT_URI"] == "https://planner.example.net/getAToken"
    assert app.config["SESSION_COOKIE_SECURE"] is True


def test_missing_config_is_a_clear_error(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    for key in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SECRET_KEY"):
        monkeypatch.delenv(f"PLANNER_WEB_{key}", raising=False)
    from webapp.config import ConfigError, load_config
    with pytest.raises(ConfigError) as exc:
        load_config()
    assert "TENANT_ID" in str(exc.value)


def test_anonymous_request_is_redirected_to_login(app):
    resp = app.test_client().get("/")
    assert resp.status_code == 302
    assert resp.headers["Location"].endswith("/login")


def test_signed_in_index_renders_both_forms(client):
    body = client.get("/").get_data(as_text=True)
    assert "Run import" in body and "Export to CSV" in body
    assert "Admin" in body and "Sign out" in body


# ---------------------------------------------------------------------------
# Input validation (before anything streams)
# ---------------------------------------------------------------------------

def test_import_without_plan_id_is_a_400(client):
    resp = client.post("/import", data={"file": (io.BytesIO(b"title\nA\n"), "t.csv")})
    assert resp.status_code == 400
    assert "plan ID" in resp.get_data(as_text=True)


def test_import_without_a_file_is_a_400(client):
    resp = client.post("/import", data={"plan_id": "PLAN1"})
    assert resp.status_code == 400


def test_unparseable_file_is_a_400_not_a_broken_stream(client):
    resp = client.post("/import", data={
        "plan_id": "PLAN1",
        "file": (io.BytesIO(b'{"tasks": [oops'), "t.json"),
    })
    assert resp.status_code == 400
    assert "Could not read that file" in resp.get_data(as_text=True)


def test_empty_file_is_a_400(client):
    resp = client.post("/import", data={
        "plan_id": "PLAN1", "file": (io.BytesIO(b"title\n"), "t.csv"),
    })
    assert resp.status_code == 400


# ---------------------------------------------------------------------------
# Streaming runs
# ---------------------------------------------------------------------------

def test_import_streams_progress_and_offers_a_receipt(client, graph):
    resp = client.post("/import", data={
        "plan_id": "PLAN1",
        "file": (io.BytesIO(b"title,bucket\nAlpha,Backlog\n"), "tasks.csv"),
    })
    body = resp.get_data(as_text=True)

    assert resp.status_code == 200
    assert "Creating &#39;Alpha&#39;" in body or "Creating 'Alpha'" in body
    assert "Created 1, updated 0" in body
    assert "Download receipt CSV" in body
    assert len(graph.tasks) == 1

    href = body.split('href="')[-2].split('"')[0] if "download" in body else None
    assert href and "/download/" in href
    csv_resp = client.get(href)
    assert csv_resp.status_code == 200
    assert "Alpha" in csv_resp.get_data(as_text=True)


def test_web_dry_run_writes_nothing(client, graph):
    bucket = graph.add_bucket("Backlog")
    graph.add_task("Alpha", bucket)

    resp = client.post("/import", data={
        "plan_id": "PLAN1", "dry_run": "on",
        "file": (io.BytesIO(b"title,bucket\nAlpha,Backlog\nBeta,New\n"), "tasks.csv"),
    })
    body = resp.get_data(as_text=True)

    assert graph.writes() == []
    assert "would create 1, would update 1" in body
    assert "Nothing was written to Planner." in body
    assert "Download receipt CSV" not in body
    assert len(graph.tasks) == 1


def test_export_streams_and_offers_the_csv(client, graph):
    bucket = graph.add_bucket("Backlog")
    graph.add_task("Alpha", bucket, priority=1)

    resp = client.post("/export", data={"plan_id": "PLAN1"})
    body = resp.get_data(as_text=True)

    assert "Exported 1 tasks." in body
    assert graph.writes() == []

    href = body.split('href="')[-2].split('"')[0]
    csv_resp = client.get(href)
    assert "Alpha" in csv_resp.get_data(as_text=True)
    assert "Urgent" in csv_resp.get_data(as_text=True)


def test_permission_error_is_reported_in_the_stream(client, graph):
    graph.users_status = 403
    resp = client.post("/import", data={
        "plan_id": "PLAN1",
        "file": (io.BytesIO(b"title,assigned_to\nAlpha,a@example.com\n"), "t.csv"),
    })
    body = resp.get_data(as_text=True)

    assert "Missing permission" in body
    assert "User.ReadBasic.All" in body
    assert graph.writes() == []


def test_forbidden_plan_explains_group_membership(client, graph, monkeypatch):
    from tests.fake_graph import FakeResponse

    def forbidden(session, method, url, **kwargs):
        return FakeResponse(403, {"error": {"code": "AccessDenied"}}, method=method, url=url)

    monkeypatch.setattr("requests.Sessi" "on.request", forbidden)
    resp = client.post("/export", data={"plan_id": "PLAN1"})
    body = resp.get_data(as_text=True)
    assert "Graph returned 403" in body
    assert "delegated permissions" in body  # the group-membership hint


# ---------------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("run_id,kind", [
    ("../../etc", "receipt"),
    ("not-a-uuid", "receipt"),
    ("0" * 32, "../../config"),
])
def test_download_rejects_bad_identifiers(client, run_id, kind):
    assert client.get(f"/download/{run_id}/{kind}").status_code == 404


def test_download_of_unknown_run_is_404(client):
    assert client.get(f"/download/{'a' * 32}/receipt").status_code == 404
