"""
In-memory fake of the Graph endpoints planner_core touches.

Installed by monkeypatching requests.Session.request, which is the single seam
every call now goes through (import and export used to have separate clients;
they don't anymore). Lets the upsert, matching, checklist-diff, throttling and
permission paths be exercised without a tenant.
"""

import json
import re
import uuid
from urllib.parse import parse_qs, unquote, urlparse

import requests

GRAPH = "https://graph.microsoft.com/v1.0"


class FakeResponse:
    def __init__(self, status_code, body=None, headers=None, method="GET", url=""):
        self.status_code = status_code
        self._body = body
        self.headers = headers or {}
        self.url = url
        self.request = type("Req", (), {"method": method, "url": url})()

    @property
    def ok(self):
        return self.status_code < 400

    def json(self):
        return self._body

    @property
    def text(self):
        return json.dumps(self._body) if self._body is not None else ""

    def raise_for_status(self):
        if not self.ok:
            raise requests.HTTPError(f"{self.status_code} for {self.url}", response=self)


class FakeGraph:
    """
    Construct, then install with `fake.install(monkeypatch)`.

    Knobs for the failure paths:
      users_status  – status to return from GET /users/{...} (403 = no
                      directory read permission, 404 = unknown person)
      throttle_once – first mutating call returns 429
      expired_token – reject any bearer token not in `valid_tokens`
    """

    def __init__(self, plan_id="PLAN1", page_size=100):
        self.plan_id = plan_id
        self.page_size = page_size
        self.buckets = {}          # id -> name
        self.tasks = {}            # id -> task dict
        self.details = {}          # task id -> details dict
        self.users = {}            # email/id -> id
        self.calls = []            # (method, url)
        self.users_status = 200
        self.throttle_once = False
        self.valid_tokens = None   # None = accept anything
        self._throttled = False

    # -- fixture helpers ----------------------------------------------------

    def add_bucket(self, name, bucket_id=None):
        bucket_id = bucket_id or f"bucket-{len(self.buckets) + 1}"
        self.buckets[bucket_id] = name
        return bucket_id

    def add_task(self, title, bucket_id, task_id=None, **fields):
        task_id = task_id or f"task-{len(self.tasks) + 1}"
        task = {
            "id": task_id,
            "@odata.etag": f'W/"etag-{task_id}-1"',
            "title": title,
            "bucketId": bucket_id,
            "planId": self.plan_id,
            "percentComplete": 0,
            "priority": 5,
            "startDateTime": None,
            "dueDateTime": None,
            "assignments": {},
        }
        task.update(fields)
        self.tasks[task_id] = task
        self.details.setdefault(task_id, {"id": task_id, "description": "", "checklist": {}})
        return task_id

    def add_user(self, email, user_id=None):
        user_id = user_id or f"user-{len(self.users) + 1}"
        self.users[email] = user_id
        self.users[user_id] = user_id
        return user_id

    def install(self, monkeypatch):
        fake = self

        def _request(session, method, url, **kwargs):
            return fake.handle(session, method, url, **kwargs)

        monkeypatch.setattr(requests.Session, "request", _request, raising=True)
        return self

    # -- assertions ---------------------------------------------------------

    def writes(self):
        return [(m, u) for m, u in self.calls if m != "GET"]

    # -- dispatch -----------------------------------------------------------

    def handle(self, session, method, url, **kwargs):
        method = method.upper()
        self.calls.append((method, url))
        path = unquote(urlparse(url).path)  # Graph decodes percent-escapes in path segments
        query = parse_qs(urlparse(url).query)

        if self.valid_tokens is not None:
            token = (session.headers.get("Authorization") or "").removeprefix("Bearer ")
            if token not in self.valid_tokens:
                return FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken"}},
                                    method=method, url=url)

        if self.throttle_once and method != "GET" and not self._throttled:
            self._throttled = True
            return FakeResponse(429, {"error": {"code": "TooManyRequests"}},
                                headers={"Retry-After": "0"}, method=method, url=url)

        body = kwargs.get("json") or {}

        m = re.match(r"^/v1\.0/planner/plans/([^/]+)/buckets$", path)
        if m and method == "GET":
            return self._paged(
                [{"id": bid, "name": name} for bid, name in self.buckets.items()],
                url, query, method,
            )

        m = re.match(r"^/v1\.0/planner/plans/([^/]+)/tasks$", path)
        if m and method == "GET":
            return self._paged(list(self.tasks.values()), url, query, method)

        if path == "/v1.0/planner/buckets" and method == "POST":
            bid = f"bucket-{uuid.uuid4().hex[:6]}"
            self.buckets[bid] = body["name"]
            return FakeResponse(201, {"id": bid, "name": body["name"]}, method=method, url=url)

        if path == "/v1.0/planner/tasks" and method == "POST":
            tid = f"task-{uuid.uuid4().hex[:6]}"
            task = {
                "id": tid,
                "@odata.etag": f'W/"etag-{tid}-1"',
                "planId": body.get("planId"),
                "bucketId": body.get("bucketId"),
                "title": body.get("title"),
                "percentComplete": body.get("percentComplete", 0),
                "priority": body.get("priority", 5),
                "startDateTime": body.get("startDateTime"),
                "dueDateTime": body.get("dueDateTime"),
                "assignments": body.get("assignments", {}),
            }
            self.tasks[tid] = task
            self.details[tid] = {"id": tid, "description": "", "checklist": {}}
            return FakeResponse(201, task, method=method, url=url)

        m = re.match(r"^/v1\.0/planner/tasks/([^/]+)/details$", path)
        if m:
            tid = m.group(1)
            if tid not in self.details:
                return FakeResponse(404, {"error": {"code": "NotFound"}}, method=method, url=url)
            if method == "GET":
                return FakeResponse(200, dict(self.details[tid]),
                                    headers={"ETag": f'W/"details-{tid}"'}, method=method, url=url)
            if method == "PATCH":
                current = self.details[tid]
                if "description" in body:
                    current["description"] = body["description"]
                if "checklist" in body:
                    checklist = current.setdefault("checklist", {})
                    for key, value in body["checklist"].items():
                        if value is None:
                            checklist.pop(key, None)
                        else:
                            checklist[key] = {
                                "title": value["title"],
                                "isChecked": value.get("isChecked", False),
                            }
                return FakeResponse(204, None, method=method, url=url)

        m = re.match(r"^/v1\.0/planner/tasks/([^/]+)$", path)
        if m and method == "PATCH":
            tid = m.group(1)
            if tid not in self.tasks:
                return FakeResponse(404, {"error": {"code": "NotFound"}}, method=method, url=url)
            task = self.tasks[tid]
            for key, value in body.items():
                if key == "assignments":
                    assignments = task.setdefault("assignments", {})
                    for uid, obj in value.items():
                        if obj is None:
                            assignments.pop(uid, None)
                        else:
                            assignments[uid] = obj
                else:
                    task[key] = value
            return FakeResponse(204, None, method=method, url=url)

        m = re.match(r"^/v1\.0/users/(.+)$", path)
        if m and method == "GET":
            if self.users_status != 200:
                return FakeResponse(
                    self.users_status,
                    {"error": {"code": "Authorization_RequestDenied",
                               "message": "Insufficient privileges to complete the operation."}},
                    method=method, url=url,
                )
            key = m.group(1)
            if key not in self.users:
                return FakeResponse(404, {"error": {"code": "Request_ResourceNotFound"}},
                                    method=method, url=url)
            uid = self.users[key]
            email = next((e for e, v in self.users.items() if v == uid and "@" in e), None)
            return FakeResponse(200, {"id": uid, "mail": email, "userPrincipalName": email},
                                method=method, url=url)

        return FakeResponse(404, {"error": {"code": "UnknownRoute", "message": path}},
                            method=method, url=url)

    def _paged(self, items, url, query, method):
        page = int(query.get("_page", ["0"])[0])
        start = page * self.page_size
        chunk = items[start:start + self.page_size]
        body = {"value": chunk}
        if start + self.page_size < len(items):
            base = url.split("?")[0]
            body["@odata.nextLink"] = f"{base}?_page={page + 1}"
        return FakeResponse(200, body, method=method, url=url)
