"""Read a SyncScript Excel export (Iyuno/SDI layout)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl


@dataclass
class ScriptRow:
    character: str
    tc: str
    note: str
    text: str


@dataclass
class CharacterInfo:
    name: str
    original_name: str = ""
    gender: str = ""  # as exported: Male / Female / Male-Female
    singing: bool = False


@dataclass
class Episode:
    rows: list[ScriptRow] = field(default_factory=list)
    characters: dict[str, CharacterInfo] = field(default_factory=dict)
    original_title: str = ""
    translated_title: str = ""
    translator: str = ""
    translation_date: str = ""
    series_title: str = ""
    season: str = ""
    episode: str = ""


def _s(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def _header_index(header: tuple) -> dict[str, int]:
    return {_s(h).lower(): i for i, h in enumerate(header) if _s(h)}


def read_export(path: str | Path) -> Episode:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ep = Episode()

    ws = wb["script"] if "script" in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    idx = _header_index(next(rows))
    try:
        ci, ti, ni, xi = (idx["character"], idx["tc"], idx["note"], idx["text"])
    except KeyError as exc:
        raise ValueError(
            "Het tabblad 'script' mist een van de kolommen Character, TC, Note, TEXT"
        ) from exc
    for r in rows:
        r = tuple(r) + (None,) * 4
        row = ScriptRow(_s(r[ci]), _s(r[ti]), _s(r[ni]), _s(r[xi]))
        if row.character or row.text:
            ep.rows.append(row)

    if "character count" in wb.sheetnames:
        rows = wb["character count"].iter_rows(values_only=True)
        idx = _header_index(next(rows))
        for r in rows:
            name = _s(r[idx.get("character", 0)])
            if not name:
                continue
            get = lambda k: _s(r[idx[k]]) if k in idx and idx[k] < len(r) else ""
            singing = get("singing").lower() in ("1", "yes", "ja", "true", "x", "y")
            ep.characters[name] = CharacterInfo(
                name, get("original name"), get("gender"), singing
            )

    if "metadata" in wb.sheetnames:
        rows = wb["metadata"].iter_rows(values_only=True)
        idx = _header_index(next(rows))
        for r in rows:
            key = _s(r[0]).lower()
            trans = _s(r[idx["translation"]]) if "translation" in idx else ""
            orig = _s(r[idx["original"]]) if "original" in idx else ""
            if key == "episode title":
                ep.original_title = orig or trans
                ep.translated_title = trans
            elif key == "translator":
                ep.translator = trans or orig
            elif key == "translation date":
                ep.translation_date = trans or orig

    _guess_from_filename(ep, Path(path).stem)
    return ep


def _guess_from_filename(ep: Episode, stem: str) -> None:
    """SyncScript names files '<project> S<season>E<episode> <title>'.

    The series title is the project name: the text before ' S'. Often the season
    slot holds a name instead of a number, for example
    'Spellbound SSpellboundE01 Spellbound FTR' or
    'Zombies Addisons Monster Mystery SZombies 4E01 ZOMBIES 4- ...'.
    """
    m = re.match(r"(.+?)\s+S(\d+)\s*E(\d+)\b", stem)
    if m:
        ep.series_title = m.group(1).strip()
        ep.season, ep.episode = str(int(m.group(2))), str(int(m.group(3)))
        return
    m = re.search(r"(?<=\S)\s*E(\d+)\b", stem)
    if not m:
        return
    ep.episode = str(int(m.group(1)))
    prefix = stem[: m.start()]
    splits = [(prefix[: s.start()].strip(), prefix[s.end():].strip())
              for s in re.finditer(r"\s+S(?=\S)", prefix)]
    # Best: the name after the S starts with the project name again.
    for project, after in splits:
        if project and after.startswith(project):
            ep.series_title = project
            return
    for project, after in splits:
        if project and (after[:1].isupper() or after[:1].isdigit()):
            ep.series_title = project
            return
