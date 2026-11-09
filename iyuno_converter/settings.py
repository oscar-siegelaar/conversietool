"""Saved defaults for the cover page, kept per user on this computer.

macOS:   ~/Library/Application Support/Iyuno Converter/settings.json
Windows: %APPDATA%\\Iyuno Converter\\settings.json
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

APP_DIR = "Iyuno Converter"

# Cover fields that make sense as a default. Titles, episode, date and the SDI auth
# code differ per script (the auth code is issued per title), so they are left out.
DEFAULT_FIELDS = [
    ("series_title", "Series title"),
    ("translated_series_title", "Translated series title"),
    ("season", "Season #"),
    ("language", "Language"),
    ("translated_by", "Translated by"),
]


def settings_path() -> Path:
    override = os.environ.get("IYUNO_CONVERTER_SETTINGS")
    if override:
        return Path(override)
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    elif sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / APP_DIR / "settings.json"


def load_defaults() -> dict[str, str]:
    """The saved cover defaults; empty when nothing was saved or the file is unreadable."""
    try:
        data = json.loads(settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    cover = data.get("cover_defaults", {}) if isinstance(data, dict) else {}
    keys = {k for k, _ in DEFAULT_FIELDS}
    return {k: str(v).strip() for k, v in cover.items() if k in keys and str(v).strip()}


def save_defaults(values: dict[str, str]) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    keys = {k for k, _ in DEFAULT_FIELDS}
    clean = {k: v.strip() for k, v in values.items() if k in keys and v.strip()}
    path.write_text(json.dumps({"cover_defaults": clean}, indent=2, ensure_ascii=False), encoding="utf-8")
    return path
