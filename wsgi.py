#!/usr/bin/env python3
"""
Entry point for the web app.

Local:
    python wsgi.py                     # or: flask --app wsgi run --port 5000

Azure App Service (later — not configured this session):
    gunicorn --bind 0.0.0.0:8000 --timeout 600 wsgi:app

The long gunicorn timeout matters: runs stream for as long as the import takes.
App Service also kills any request that produces no bytes for 230 seconds,
which is exactly why runs stream progress rather than blocking.
"""

import os

from webapp import create_app

app = create_app()

if __name__ == "__main__":
    # Port comes from config's BASE_URL so it can't drift from the redirect URI
    # registered in Entra.
    from urllib.parse import urlparse

    parsed = urlparse(app.config["BASE_URL"])
    app.run(
        host=os.environ.get("PLANNER_WEB_HOST", "127.0.0.1"),
        port=parsed.port or 5000,
        debug=os.environ.get("PLANNER_WEB_DEBUG") == "1",
        threaded=True,
    )
