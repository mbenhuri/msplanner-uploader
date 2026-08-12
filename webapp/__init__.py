"""
Flask application factory.

Nothing environment-specific lives in code: the base URL, redirect URI, data
directory and credentials all come from webapp/config.py. Moving to Azure App
Service is a config change (see README).
"""

import os
import sys

from flask import Flask

# Allow `import planner_core` when the app is launched from anywhere.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from .auth import AuthManager  # noqa: E402
from .config import load_config  # noqa: E402


def create_app(overrides=None):
    cfg = load_config(overrides)

    app = Flask(__name__)
    app.config.update(cfg)
    app.secret_key = cfg["SECRET_KEY"]
    app.config["MAX_CONTENT_LENGTH"] = cfg["MAX_CONTENT_LENGTH"]
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    # Lax, not Strict: Entra's redirect back to /getAToken is a cross-site
    # top-level GET, and Strict would drop the cookie carrying the auth flow.
    # Lax also means a cross-site POST can't ride along with the session
    # cookie, which is what stands in for CSRF tokens on this single-user tool.
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
    app.config["SESSION_COOKIE_SECURE"] = cfg["BASE_URL"].startswith("https://")

    app.extensions["planner_auth"] = AuthManager(cfg)

    from .views import bp  # noqa: WPS433 - avoids a circular import at module load
    app.register_blueprint(bp)

    return app
