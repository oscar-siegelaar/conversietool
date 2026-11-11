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

# Cover fields that make sense as a default. Titles, episode and date differ per
# script, so they are left out. The SDI auth code is the same for every script; the
# user enters it once and it is kept only in the settings file on this computer.
DEFAULT_FIELDS = [
    ("series_title", "Series title"),
    ("translated_series_title", "Translated series title"),
    ("season", "Season #"),
    ("language", "Language"),
    ("translated_by", "Translated by"),
    ("auth_code", "SDI auth code (from Iyuno)"),
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


def _load() -> dict:
    try:
        data = json.loads(settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _save(data: dict) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def load_defaults() -> dict[str, str]:
    """The saved cover defaults; empty when nothing was saved or the file is unreadable."""
    cover = _load().get("cover_defaults", {})
    if not isinstance(cover, dict):
        return {}
    keys = {k for k, _ in DEFAULT_FIELDS}
    return {k: str(v).strip() for k, v in cover.items() if k in keys and str(v).strip()}


def save_defaults(values: dict[str, str]) -> Path:
    keys = {k for k, _ in DEFAULT_FIELDS}
    data = _load()
    data["cover_defaults"] = {k: v.strip() for k, v in values.items() if k in keys and v.strip()}
    return _save(data)


def load_filename_template() -> str:
    from .naming import DEFAULT_TEMPLATE
    value = _load().get("filename_template")
    return value.strip() if isinstance(value, str) and value.strip() else DEFAULT_TEMPLATE


def save_filename_template(template: str) -> Path:
    data = _load()
    data["filename_template"] = template.strip()
    return _save(data)
