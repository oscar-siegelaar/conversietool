"""Build the Word document's file name from a template with cover-field labels.

A template is text with labels in square brackets, for example
'[LANGUAGE:upper]_[ORIGINAL TITLE]_[TRANSLATED BY:initials]' -> 'DUTCH_Puppy Pool Party_MO'.
A label may carry a style: ':upper' (CAPITALS) or ':initials'. Without a style the value
is used as it is.
"""
from __future__ import annotations

import datetime as dt
import re

LABELS = [
    ("LANGUAGE", "language"),
    ("ORIGINAL TITLE", "original_title"),
    ("TRANSLATED TITLE", "translated_title"),
    ("SERIES TITLE", "series_title"),
    ("TRANSLATED SERIES TITLE", "translated_series_title"),
    ("SEASON", "season"),
    ("EPISODE", "episode"),
    ("DATE", "date"),
    ("TRANSLATED BY", "translated_by"),
]
STYLES = [("", "as is"), ("initials", "initials"), ("upper", "CAPITALS")]
DEFAULT_TEMPLATE = "[LANGUAGE:upper]_[ORIGINAL TITLE]_[TRANSLATED BY:initials]"

_FIELD = {label: attr for label, attr in LABELS}
_TOKEN = re.compile(r"\[\s*([A-Za-z #]+?)\s*(?::\s*(initials|upper)\s*)?\]", re.I)


def initials(text: str) -> str:
    """'Mark Omvlee' -> 'MO'; 'Jan van der Berg' -> 'JvdB'. Anything after ' - ', a comma,
    '(' or an e-mail address is ignored, so 'Mark Omvlee - mark@… - 06…' still gives 'MO'."""
    name = re.split(r"\s+-\s+|,|\(|<|\S+@\S+", text, maxsplit=1)[0]
    return "".join(w[0] for w in name.split() if w[:1].isalnum())


def _value(cover, attr: str, today: dt.date) -> str:
    value = str(getattr(cover, attr, "") or "").strip()
    if attr == "date" and not value:
        value = f"{today.day}-{today.month}-{today.year}"
    return value


def safe_filename(name: str) -> str:
    """Replace characters that Mac or Windows do not allow in file names:
    ':' becomes ' -', '/ \\ * ? " < > |' become '_'; double spaces and a trailing dot
    or space go."""
    name = re.sub(r"\s*:\s*", " - ", name)
    name = re.sub(r'[\\/*?"<>|]', "_", name)
    name = re.sub(r"[\x00-\x1f]", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().rstrip(". ").strip()[:150]


def render(template: str, cover, today: dt.date | None = None) -> str:
    """The file name (without .docx) for this cover."""
    today = today or dt.date.today()

    def sub(m: re.Match) -> str:
        label = re.sub(r"\s+", " ", m.group(1).strip().upper().replace("#", "")).strip()
        attr = _FIELD.get(label)
        if attr is None:
            return m.group(0)
        value = _value(cover, attr, today)
        style = (m.group(2) or "").lower()
        if style == "initials":
            value = initials(value)
        elif style == "upper":
            value = value.upper()
        return value

    name = _TOKEN.sub(sub, template or DEFAULT_TEMPLATE)
    # Empty labels leave doubled separators ('DUTCH__MO'): tidy them up.
    name = re.sub(r"([_\-.])\1+", r"\1", name)
    name = safe_filename(name).strip("_-. ")
    return name or "script"


def unique_name(stem: str, taken: set[str]) -> str:
    """'name', then 'name_2', 'name_3', ... (case-insensitive, for Mac and Windows)."""
    candidate, n = stem, 2
    while candidate.lower() in taken:
        candidate, n = f"{stem}_{n}", n + 1
    taken.add(candidate.lower())
    return candidate
