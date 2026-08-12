"""
Web app configuration.

Everything environment-specific lives here and comes from env vars (or an
optional config.web.json). Nothing in the app hardcodes localhost — moving to
Azure App Service means setting PLANNER_WEB_BASE_URL to the site's https URL,
adding that redirect URI to the app registration, and pointing
PLANNER_WEB_DATA_DIR at /home (the only path App Service persists).

The web app deliberately uses its own PLANNER_WEB_* variables rather than the
CLI's PLANNER_* ones: they are two separate app registrations with two
different auth models, and crossing them would be easy and confusing.
"""

import json
import os

CONFIG_FILE = "config.web.json"

# Delegated scopes. Kept in planner_core so the CLI docs and the web app can't
# drift apart.
from planner_core import DELEGATED_SCOPES  # noqa: E402

DEFAULTS = {
    "BASE_URL": "http://localhost:5000",
    "REDIRECT_PATH": "/getAToken",
    "DATA_DIR": os.path.abspath("webapp_data"),
    "MAX_CONTENT_LENGTH": 16 * 1024 * 1024,  # 16 MB uploads
}

REQUIRED = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SECRET_KEY"]


class ConfigError(RuntimeError):
    pass


def _from_file():
    if not os.path.exists(CONFIG_FILE):
        return {}
    with open(CONFIG_FILE) as f:
        raw = json.load(f)
    return {k.upper(): v for k, v in raw.items()}


def load_config(overrides=None):
    cfg = dict(DEFAULTS)
    cfg.update(_from_file())

    for key in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SECRET_KEY",
                "BASE_URL", "REDIRECT_PATH", "DATA_DIR"):
        env = os.environ.get(f"PLANNER_WEB_{key}")
        if env:
            cfg[key] = env

    if overrides:
        cfg.update(overrides)

    missing = [k for k in REQUIRED if not cfg.get(k)]
    if missing:
        raise ConfigError(
            "Missing web app config: "
            + ", ".join(missing)
            + ". Set them as PLANNER_WEB_* environment variables or in config.web.json "
              "(see config.web.sample.json). SECRET_KEY signs the Flask session cookie — "
              "generate one with: python -c \"import secrets; print(secrets.token_hex(32))\""
        )

    cfg["AUTHORITY"] = cfg.get("AUTHORITY") or f"https://login.microsoftonline.com/{cfg['TENANT_ID']}"
    cfg["BASE_URL"] = cfg["BASE_URL"].rstrip("/")
    cfg["REDIRECT_URI"] = cfg["BASE_URL"] + cfg["REDIRECT_PATH"]
    cfg["SCOPES"] = list(cfg.get("SCOPES") or DELEGATED_SCOPES)

    data_dir = os.path.abspath(cfg["DATA_DIR"])
    cfg["DATA_DIR"] = data_dir
    cfg["STATE_DIR"] = cfg.get("STATE_DIR") or os.path.join(data_dir, "planner_state")
    cfg["RUNS_DIR"] = cfg.get("RUNS_DIR") or os.path.join(data_dir, "runs")
    cfg["TOKEN_CACHE_PATH"] = cfg.get("TOKEN_CACHE_PATH") or os.path.join(data_dir, "token_cache.bin")

    for path in (cfg["DATA_DIR"], cfg["STATE_DIR"], cfg["RUNS_DIR"]):
        os.makedirs(path, exist_ok=True)

    return cfg
