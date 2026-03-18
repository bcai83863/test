from __future__ import annotations

from importlib import import_module

# Render default start command is often `gunicorn app:app`.
dash_app_module = import_module("\u5c31\u696d\u91d1\u5361_dash.app")
app = dash_app_module.server

