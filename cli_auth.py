#!/usr/bin/env python3
"""
App-only (client credentials) auth for the CLI scripts.

This is the *only* thing that differs between the CLI tools and the web app:
the CLI runs unattended as a tenant-wide app identity, the web app runs as the
signed-in user via delegated auth code flow. Both then hand a token provider to
the same planner_core.GraphClient.

Config comes from config.json or PLANNER_TENANT_ID / PLANNER_CLIENT_ID /
PLANNER_CLIENT_SECRET. The web app reads its own separate config (see
webapp/config.py) so the two app registrations can't be crossed.
"""

import json
import os
import sys

import msal


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


def build_token_provider(cfg):
    """
    Returns a callable(force_refresh=False) -> access token string.

    MSAL caches the app token internally and renews it when it expires, so a
    long import that outlives one token keeps working; force_refresh is used
    by GraphClient when Graph rejects a token it still believed was valid.
    """
    app = msal.ConfidentialClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        client_credential=cfg["client_secret"],
    )
    scopes = ["https://graph.microsoft.com/.default"]

    def provider(force_refresh=False):
        try:
            result = app.acquire_token_for_client(scopes=scopes, force_refresh=force_refresh)
        except TypeError:
            # Older msal builds don't accept force_refresh.
            result = app.acquire_token_for_client(scopes=scopes)
        if "access_token" not in result:
            sys.exit(f"Auth failed: {result.get('error_description', result)}")
        return result["access_token"]

    return provider
