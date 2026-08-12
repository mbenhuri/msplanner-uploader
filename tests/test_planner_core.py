"""
Tests for the shared logic in planner_core.

Weighted toward the parts most likely to break silently during the refactor:
checklist diffing, upsert matching, and the safety properties (dry run writes
nothing, percent_complete is create-only, blank fields never clear).
"""

import json
import os
import sys

import pytest
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import planner_core as core  # noqa: E402
from tests.fake_graph import GRAPH, FakeGraph  # noqa: E402


@pytest.fixture
def graph(monkeypatch):
    return FakeGraph().install(monkeypatch)


def client_for(graph, dry_run=False, token="tok"):
    cls = core.ReadOnlyGraphClient if dry_run else core.GraphClient
    return cls(token, emit=lambda _msg: None)


def run(graph, rows, tmp_path, dry_run=False, receipt=True):
    client = client_for(graph, dry_run)
    return core.run_import(
        client,
        graph.plan_id,
        rows,
        dry_run=dry_run,
        state_dir=str(tmp_path / "state"),
        receipt_path=str(tmp_path / "receipt.csv") if receipt else None,
        emit=lambda _msg: None,
    )


# ---------------------------------------------------------------------------
# Checklist diffing
# ---------------------------------------------------------------------------

def test_checklist_unchanged_returns_none():
    existing = {"a": {"title": "Draft", "isChecked": True}}
    assert core.build_checklist_patch(["Draft"], existing) is None


def test_checklist_ignores_case_and_whitespace():
    existing = {"a": {"title": "Draft Spec", "isChecked": False}}
    assert core.build_checklist_patch(["  draft spec "], existing) is None


def test_checklist_add_keeps_existing_id_and_checked_state():
    existing = {"guid-1": {"title": "Draft", "isChecked": True}}
    patch = core.build_checklist_patch(["Draft", "Review"], existing)
    assert patch["guid-1"] == {
        "@odata.type": "#microsoft.graph.plannerChecklistItem",
        "title": "Draft",
        "isChecked": True,
    }
    added = [v for k, v in patch.items() if k != "guid-1"]
    assert len(added) == 1
    assert added[0]["title"] == "Review"
    assert added[0]["isChecked"] is False


def test_checklist_removal_sets_none():
    existing = {"guid-1": {"title": "Draft", "isChecked": False},
                "guid-2": {"title": "Obsolete", "isChecked": True}}
    patch = core.build_checklist_patch(["Draft"], existing)
    assert patch["guid-2"] is None
    assert patch["guid-1"]["title"] == "Draft"


def test_checklist_from_empty():
    patch = core.build_checklist_patch(["One", "Two"], None)
    assert len(patch) == 2
    assert {v["title"] for v in patch.values()} == {"One", "Two"}


def test_checklist_reorder_only_is_not_a_change():
    existing = {"a": {"title": "One", "isChecked": False},
                "b": {"title": "Two", "isChecked": False}}
    assert core.build_checklist_patch(["Two", "One"], existing) is None


# ---------------------------------------------------------------------------
# Field parsing
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("value,expected", [
    ("Urgent", 1), ("important", 3), ("MEDIUM", 5), ("Low", 9),
    ("7", 7), ("99", 10), ("-4", 0), ("", None), (None, None), ("nonsense", None),
])
def test_parse_priority(value, expected):
    assert core.parse_priority(value) == expected


@pytest.mark.parametrize("value,expected", [
    (0, "Urgent"), (1, "Urgent"), (2, "Important"), (4, "Important"),
    (5, "Medium"), (7, "Medium"), (8, "Low"), (10, "Low"),
    (None, ""), ("junk", ""),
])
def test_priority_to_label(value, expected):
    assert core.priority_to_label(value) == expected


def test_priority_round_trips_through_label():
    for label in ("Urgent", "Important", "Medium", "Low"):
        assert core.priority_to_label(core.parse_priority(label)) == label


def test_to_graph_date():
    assert core.to_graph_date("2026-03-04") == "2026-03-04T00:00:00Z"
    assert core.to_graph_date("04/03/2026") is None
    assert core.to_graph_date("") is None


def test_date_only():
    assert core.date_only("2026-03-04T00:00:00Z") == "2026-03-04"
    assert core.date_only(None) == ""


def test_row_key_prefers_explicit_id():
    assert core.row_key({"id": "abc", "title": "T", "bucket": "B"}) == "id:abc"
    assert core.row_key({"title": "T", "bucket": "B"}) == "bt:b::t"
    assert core.row_key({"title": "T"}) == "bt:to do::t"


# ---------------------------------------------------------------------------
# Input loading
# ---------------------------------------------------------------------------

def test_load_rows_csv_with_bom():
    data = "﻿title,bucket\nAlpha,Backlog\n".encode("utf-8")
    rows = core.load_rows(data, "tasks.csv")
    assert rows == [{"title": "Alpha", "bucket": "Backlog"}]


def test_load_rows_json_array_and_wrapper():
    assert core.load_rows(b'[{"title": "A"}]', "t.json") == [{"title": "A"}]
    assert core.load_rows(b'{"tasks": [{"title": "B"}]}', "t.json") == [{"title": "B"}]


def test_load_rows_sniffs_json_content_despite_csv_name():
    # An uploaded filename isn't trustworthy; content decides.
    assert core.load_rows(b'[{"title": "A"}]', "tasks.csv") == [{"title": "A"}]


def test_load_rows_from_path(tmp_path):
    path = tmp_path / "t.csv"
    path.write_text("title\nAlpha\n")
    assert core.load_rows(str(path)) == [{"title": "Alpha"}]


# ---------------------------------------------------------------------------
# Upsert matching
# ---------------------------------------------------------------------------

def test_creates_when_nothing_matches(graph, tmp_path):
    summary = run(graph, [{"title": "Alpha", "bucket": "Backlog"}], tmp_path)
    assert (summary.created, summary.updated) == (1, 0)
    assert [t["title"] for t in graph.tasks.values()] == ["Alpha"]
    assert "Backlog" in graph.buckets.values()


def test_matches_by_bucket_and_title_without_state_file(graph, tmp_path):
    bid = graph.add_bucket("Backlog")
    graph.add_task("Alpha", bid)
    summary = run(graph, [{"title": "Alpha", "bucket": "Backlog", "priority": "Urgent"}], tmp_path)
    assert (summary.created, summary.updated) == (0, 1)
    assert len(graph.tasks) == 1
    assert list(graph.tasks.values())[0]["priority"] == 1


def test_matches_by_explicit_task_id_even_after_rename(graph, tmp_path):
    bid = graph.add_bucket("Backlog")
    tid = graph.add_task("Old name", bid)
    summary = run(graph, [{"id": tid, "title": "New name", "bucket": "Backlog"}], tmp_path)
    assert (summary.created, summary.updated) == (0, 1)
    assert graph.tasks[tid]["title"] == "New name"
    assert len(graph.tasks) == 1


def test_matches_via_state_file_after_rename(graph, tmp_path):
    bid = graph.add_bucket("Backlog")
    tid = graph.add_task("Original", bid)
    state_dir = tmp_path / "state"
    state_dir.mkdir()
    (state_dir / f"{graph.plan_id}.json").write_text(
        json.dumps({"id:kickoff": {"task_id": tid, "bucket_id": bid}})
    )
    summary = run(graph, [{"id": "kickoff", "title": "Renamed", "bucket": "Backlog"}], tmp_path)
    assert (summary.created, summary.updated) == (0, 1)
    assert graph.tasks[tid]["title"] == "Renamed"


def test_duplicate_rows_in_one_file_do_not_create_twice(graph, tmp_path):
    rows = [{"title": "Alpha", "bucket": "Backlog"}, {"title": "Alpha", "bucket": "Backlog"}]
    summary = run(graph, rows, tmp_path)
    assert summary.created == 1
    assert len(graph.tasks) == 1


def test_paginated_task_list_still_matches(graph, tmp_path):
    """A task on page 2 must be found, or the importer duplicates it."""
    graph.page_size = 2
    bid = graph.add_bucket("Backlog")
    graph.add_task("One", bid)
    graph.add_task("Two", bid)
    graph.add_task("Target", bid)  # lands on the second page
    summary = run(graph, [{"title": "Target", "bucket": "Backlog"}], tmp_path)
    assert summary.created == 0
    assert len(graph.tasks) == 3


def test_paginated_bucket_list_is_fully_read(graph, tmp_path):
    graph.page_size = 1
    graph.add_bucket("First")
    graph.add_bucket("Second")
    run(graph, [{"title": "Alpha", "bucket": "Second"}], tmp_path)
    # No third bucket invented because "Second" was missed on page 1.
    assert sorted(graph.buckets.values()) == ["First", "Second"]


# ---------------------------------------------------------------------------
# Safety properties
# ---------------------------------------------------------------------------

def test_dry_run_makes_no_writes_at_all(graph, tmp_path):
    bid = graph.add_bucket("Backlog")
    graph.add_task("Alpha", bid)
    rows = [
        {"title": "Alpha", "bucket": "Backlog", "priority": "Urgent"},
        {"title": "Beta", "bucket": "Brand New Bucket", "checklist": "a;b"},
    ]
    summary = run(graph, rows, tmp_path, dry_run=True)

    assert graph.writes() == []
    assert (summary.would_update, summary.would_create) == (1, 1)
    assert len(graph.tasks) == 1
    assert "Brand New Bucket" not in graph.buckets.values()


def test_dry_run_writes_no_receipt_or_state(graph, tmp_path):
    summary = run(graph, [{"title": "Alpha"}], tmp_path, dry_run=True)
    assert summary.receipt_path is None
    assert not (tmp_path / "receipt.csv").exists()
    assert not os.path.exists(os.path.join(str(tmp_path / "state"), f"{graph.plan_id}.json"))


def test_read_only_client_refuses_writes(graph):
    client = core.ReadOnlyGraphClient("tok", emit=lambda _m: None)
    with pytest.raises(core.DryRunViolation):
        client.post(f"{GRAPH}/planner/tasks", json={})


def test_percent_complete_is_create_only(graph, tmp_path):
    row = {"title": "Alpha", "bucket": "Backlog", "percent_complete": "0"}
    run(graph, [row], tmp_path)
    task_id = next(iter(graph.tasks))

    # Someone marks it done in Planner...
    graph.tasks[task_id]["percentComplete"] = 100

    # ...and the same file is imported again.
    run(graph, [row], tmp_path)
    assert graph.tasks[task_id]["percentComplete"] == 100


def test_blank_fields_do_not_clear_existing_values(graph, tmp_path):
    alice = graph.add_user("alice@example.com")
    row = {
        "title": "Alpha", "bucket": "Backlog",
        "description": "Original notes", "checklist": "step one;step two",
        "assigned_to": "alice@example.com",
    }
    run(graph, [row], tmp_path)
    task_id = next(iter(graph.tasks))

    summary = run(graph, [{"title": "Alpha", "bucket": "Backlog"}], tmp_path)

    assert summary.unchanged == 1
    assert graph.details[task_id]["description"] == "Original notes"
    assert len(graph.details[task_id]["checklist"]) == 2
    assert alice in graph.tasks[task_id]["assignments"]


def test_unresolvable_assignee_does_not_wipe_existing_assignments(graph, tmp_path):
    """A typo'd email must not silently unassign everyone on the task."""
    alice = graph.add_user("alice@example.com")
    run(graph, [{"title": "Alpha", "assigned_to": "alice@example.com"}], tmp_path)
    task_id = next(iter(graph.tasks))
    assert alice in graph.tasks[task_id]["assignments"]

    run(graph, [{"title": "Alpha", "assigned_to": "typo@example.com"}], tmp_path)
    assert alice in graph.tasks[task_id]["assignments"]


def test_resolvable_assignee_change_still_syncs(graph, tmp_path):
    alice = graph.add_user("alice@example.com")
    bob = graph.add_user("bob@example.com")
    run(graph, [{"title": "Alpha", "assigned_to": "alice@example.com"}], tmp_path)
    task_id = next(iter(graph.tasks))

    run(graph, [{"title": "Alpha", "assigned_to": "bob@example.com"}], tmp_path)
    assert set(graph.tasks[task_id]["assignments"]) == {bob}
    assert alice not in graph.tasks[task_id]["assignments"]


def test_checked_state_survives_reimport(graph, tmp_path):
    row = {"title": "Alpha", "bucket": "Backlog", "checklist": "one;two"}
    run(graph, [row], tmp_path)
    task_id = next(iter(graph.tasks))

    for item in graph.details[task_id]["checklist"].values():
        if item["title"] == "one":
            item["isChecked"] = True

    run(graph, [{**row, "checklist": "one;two;three"}], tmp_path)
    checklist = graph.details[task_id]["checklist"]
    by_title = {v["title"]: v["isChecked"] for v in checklist.values()}
    assert by_title == {"one": True, "two": False, "three": False}


def test_unchanged_row_reports_unchanged(graph, tmp_path):
    row = {"title": "Alpha", "bucket": "Backlog", "priority": "Medium"}
    run(graph, [row], tmp_path)
    summary = run(graph, [row], tmp_path)
    assert (summary.unchanged, summary.updated, summary.created) == (1, 0, 0)


def test_rows_without_titles_are_skipped(graph, tmp_path):
    summary = run(graph, [{"title": ""}, {"title": "Alpha"}], tmp_path)
    assert (summary.skipped, summary.created) == (1, 1)


# ---------------------------------------------------------------------------
# Receipt and state
# ---------------------------------------------------------------------------

def test_receipt_records_every_row(graph, tmp_path):
    run(graph, [{"title": "Alpha", "bucket": "Backlog"}], tmp_path)
    text = (tmp_path / "receipt.csv").read_text()
    assert "row,title,bucket,action,task_id,task_url,timestamp" in text
    assert "Alpha" in text and "created" in text
    assert "planner.cloud.microsoft" in text


def test_state_is_saved_incrementally(graph, tmp_path):
    """A run that dies part-way still records what it wrote."""
    state_dir = tmp_path / "state"
    boom = RuntimeError("connection reset")

    original_create = core.create_task
    calls = {"n": 0}

    def flaky_create(*args, **kwargs):
        calls["n"] += 1
        if calls["n"] == 2:
            raise boom
        return original_create(*args, **kwargs)

    core.create_task = flaky_create
    try:
        with pytest.raises(RuntimeError):
            core.run_import(
                client_for(graph), graph.plan_id,
                [{"title": "Alpha"}, {"title": "Beta"}],
                state_dir=str(state_dir),
                receipt_path=str(tmp_path / "receipt.csv"),
                emit=lambda _m: None,
            )
    finally:
        core.create_task = original_create

    saved = json.loads((state_dir / f"{graph.plan_id}.json").read_text())
    assert "bt:to do::alpha" in saved
    # And the receipt for the completed row still landed.
    assert "Alpha" in (tmp_path / "receipt.csv").read_text()


# ---------------------------------------------------------------------------
# Permissions and transport
# ---------------------------------------------------------------------------

def test_missing_user_read_aborts_before_any_write(graph, tmp_path):
    graph.users_status = 403
    rows = [{"title": "Alpha", "bucket": "Backlog", "assigned_to": "alice@example.com"}]

    with pytest.raises(core.GraphPermissionError) as exc:
        run(graph, rows, tmp_path)

    assert "User.ReadBasic.All" in str(exc.value)
    assert graph.writes() == []          # nothing reached Planner
    assert graph.tasks == {}


def test_unknown_assignee_is_skipped_not_fatal(graph, tmp_path):
    graph.add_user("known@example.com")
    messages = []
    client = core.GraphClient("tok", emit=messages.append)
    summary = core.run_import(
        client, graph.plan_id,
        [{"title": "Alpha", "assigned_to": "ghost@example.com"}],
        state_dir=str(tmp_path / "state"),
        receipt_path=str(tmp_path / "receipt.csv"),
        emit=lambda _m: None,
    )
    assert summary.created == 1
    assert any("could not resolve user" in m.lower() for m in messages)
    assert list(graph.tasks.values())[0]["assignments"] == {}


def test_expired_token_is_refreshed_mid_run(graph, tmp_path):
    graph.valid_tokens = {"fresh"}
    issued = []

    def provider(force_refresh=False):
        issued.append(force_refresh)
        return "fresh" if force_refresh else "stale"

    client = core.GraphClient(provider, emit=lambda _m: None)
    summary = core.run_import(
        client, graph.plan_id, [{"title": "Alpha"}],
        state_dir=str(tmp_path / "state"),
        receipt_path=str(tmp_path / "receipt.csv"),
        emit=lambda _m: None,
    )
    assert summary.created == 1
    assert True in issued  # a forced refresh happened


def test_throttling_is_retried(graph, tmp_path, monkeypatch):
    monkeypatch.setattr(core.time, "sleep", lambda _s: None)
    graph.throttle_once = True
    summary = run(graph, [{"title": "Alpha"}], tmp_path)
    assert summary.created == 1


def test_graph_error_body_is_surfaced(graph):
    messages = []
    client = core.GraphClient("tok", emit=messages.append)
    with pytest.raises(requests.HTTPError):
        core.get_buckets(client, "no-such-plan-route/x")
    assert any("Response body" in m for m in messages)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def test_export_maps_tasks_back_to_rows(graph, tmp_path):
    alice = graph.add_user("alice@example.com")
    bid = graph.add_bucket("Backlog")
    tid = graph.add_task(
        "Alpha", bid, priority=1, percentComplete=50,
        dueDateTime="2026-03-04T00:00:00Z",
        assignments={alice: {"orderHint": " !"}},
    )
    graph.details[tid] = {
        "id": tid, "description": "Notes",
        "checklist": {"g1": {"title": "zeta", "isChecked": True},
                      "g2": {"title": "alpha", "isChecked": False}},
    }

    client = core.ReadOnlyGraphClient("tok", emit=lambda _m: None)
    rows = core.run_export(client, graph.plan_id, emit=lambda _m: None)

    assert rows == [{
        "id": tid, "title": "Alpha", "bucket": "Backlog",
        "start_date": "", "due_date": "2026-03-04",
        "priority": "Urgent", "assigned_to": "alice@example.com",
        "description": "Notes", "percent_complete": 50,
        "checklist": "alpha;zeta",  # sorted by title, as documented
    }]
    assert graph.writes() == []


def test_export_round_trips_into_an_unchanged_import(graph, tmp_path):
    alice = graph.add_user("alice@example.com")
    bid = graph.add_bucket("Backlog")
    tid = graph.add_task("Alpha", bid, priority=3, dueDateTime="2026-03-04T00:00:00Z",
                         assignments={alice: {"orderHint": " !"}})
    graph.details[tid] = {"id": tid, "description": "Notes",
                          "checklist": {"g1": {"title": "one", "isChecked": False}}}

    client = core.ReadOnlyGraphClient("tok", emit=lambda _m: None)
    rows = core.run_export(client, graph.plan_id, emit=lambda _m: None)

    summary = run(graph, rows, tmp_path)
    assert (summary.unchanged, summary.created) == (1, 0)


def test_export_write_csv(tmp_path):
    out = core.write_export_csv(
        [{k: "" for k in core.FIELDNAMES} | {"title": "Alpha"}],
        str(tmp_path / "out" / "export.csv"),
    )
    assert os.path.exists(out)
    assert "Alpha" in open(out).read()


def test_export_without_directory_read_permission_aborts(graph):
    alice = graph.add_user("alice@example.com")
    bid = graph.add_bucket("Backlog")
    graph.add_task("Alpha", bid, assignments={alice: {"orderHint": " !"}})
    graph.users_status = 403

    client = core.ReadOnlyGraphClient("tok", emit=lambda _m: None)
    with pytest.raises(core.GraphPermissionError):
        core.run_export(client, graph.plan_id, emit=lambda _m: None)
