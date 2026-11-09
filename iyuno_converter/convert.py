"""Turn a SyncScript export into the rows and statistics the SDI/Iyuno template expects.

The statistics reproduce the Word macro "SDI Template v. 5" (CountScriptLinesProc,
count method 5 "New HL method + QuickDub", 8 words per line, song verses not summed
into the character totals).
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

from .reader import Episode, ScriptRow

INSERT_TOKEN = "Insert "
SONG_TOKEN = "Song "
WORDS_PER_LINE = 8
GENDER_CODES = {"male": "M", "female": "V", "male-female": "MV", "female-male": "MV"}

_SONG_START = re.compile(r"^\{?\s*song\s*(\d+)\b(.*?)\}?\s*$", re.S | re.I)
_TC_AT_END = re.compile(r"(?:^|\s)((?:\d{1,2}:)?\d{1,2}:\d{2})\s*$")


# ---------------------------------------------------------------- helpers


def tc_seconds(tc: str) -> int | None:
    parts = tc.strip().split(":")
    if not tc.strip() or not all(p.isdigit() for p in parts):
        return None
    secs = 0
    for p in parts[:3]:
        secs = secs * 60 + int(p)
    return secs if len(parts) >= 2 else None


def clear_character_name(name: str) -> str:
    name = name.strip()
    if name[-1:] in (":", "+", ",", "/"):
        return name[:-1].strip()
    return name


def split_character_names(name: str) -> list[str]:
    for delim in ("+", ",", "/"):
        if delim in name:
            return [n for n in (clear_character_name(p) for p in name.split(delim)) if n]
    return [clear_character_name(name)]


def count_words(text: str) -> int:
    """Word count of method 5 in the macro."""
    t = text
    for a, b in ((" - ", " "), (" – ", " "), (" -", " "), ("- ", " "),
                 ("-\x0b", ""), ("...", " "), (",", " "), (".", " "),
                 ("M/E:LLÄ", "M_E_LLÄ")):
        t = t.replace(a, b)
    for ch in ':;?!#$%^&*+”„[]{}()|/\\ \r\t\n\x0b\x0c':
        t = t.replace(ch, " ")
    t = t.replace("@", "").replace("' ", " ")
    t = t.strip()
    for n in (6, 5, 4, 3, 2):
        t = t.replace(" " * n, " ")
    return t.count(" ") + 1


def verses(words: int) -> int:
    return math.ceil(words / WORDS_PER_LINE)


# ---------------------------------------------------------------- transform


@dataclass
class Song:
    mark: str                  # "Song 1"
    titles: str                # "Original // Local"
    tc_in: str
    tc_out: str
    singers: set[str] = field(default_factory=set)


def parse_song_note(note: str) -> Song | None:
    """Read a song note in any of the forms translators use.

    'Song 1 Orig//Local 09:35', 'song 4 30:49', '{song 4 30:49}', 'Song 4', or the
    three-line form 'Song 1' / 'Orig//Local' / '09:35'. The mark is normalised to
    'Song N', which the macro and Iyuno's system look for.
    """
    m = _SONG_START.match(note.strip())
    if not m:
        return None
    mark = f"Song {int(m.group(1))}"
    rest = [l.strip() for l in m.group(2).split("\n") if l.strip()]
    titles, tc_out = "", ""
    for i, line in enumerate(rest):
        t = _TC_AT_END.search(line)
        if t and (i == len(rest) - 1):
            tc_out = t.group(1)
            line = line[: t.start(1)].strip()
        if line and not titles:
            titles = line
    return Song(mark, titles, "", tc_out)


def song_note(song: Song, long: bool) -> str:
    """Long form is always three paragraphs (mark, titles, TC out): the macro reads
    them by position, so an unknown title stays an empty line."""
    if not long:
        return song.mark
    return "\n".join((song.mark, song.titles, song.tc_out))


def transform(ep: Episode) -> tuple[list[ScriptRow], list[str]]:
    """Apply the manual steps translators do today. Returns rows and warnings."""
    warnings: list[str] = []
    rows = [ScriptRow(r.character, r.tc, r.note, r.text) for r in ep.rows]

    # 1. The first row carries the episode title at 00:00:00: it belongs on the cover.
    if rows and rows[0].character.startswith(INSERT_TOKEN.strip()) and tc_seconds(rows[0].tc) == 0:
        title = rows.pop(0).text
        if "//" in title and not ep.translated_title:
            ep.original_title, ep.translated_title = (p.strip() for p in title.split("//", 1))

    # Rows without text that only carry a label (e.g. "Synopsis") are not script lines.
    rows = [r for r in rows if r.text.strip() or r.character.lower() not in ("synopsis",)]

    # 2. Number the inserts consecutively.
    n = 0
    for r in rows:
        if re.fullmatch(r"Insert(\s+\d+)?\s*", r.character):
            n += 1
            r.character = f"{INSERT_TOKEN}{n}"

    # 3. Songs: find every song from its explicit note.
    songs: dict[str, Song] = {}
    # Songs whose later lines already carry a short "Song N" note in the export: the
    # translator marked every sung line, so trust that and infer nothing.
    marked_by_export: set[str] = set()
    for r in rows:
        s = parse_song_note(r.note)
        if not s:
            continue
        if s.mark in songs and not s.titles and not s.tc_out:
            marked_by_export.add(s.mark)
        known = songs.setdefault(s.mark, s)
        if not known.tc_in:
            known.tc_in = r.tc
        if not known.tc_out and s.tc_out:
            known.tc_out = s.tc_out
        if not known.titles and s.titles:
            known.titles = s.titles
        known.singers.update(split_character_names(r.character))

    def in_song(sec: int | None, s: Song) -> bool:
        if s.mark in marked_by_export:
            return False
        a, b = tc_seconds(s.tc_in), tc_seconds(s.tc_out) if s.tc_out else None
        return sec is not None and a is not None and sec >= a and (b is None or sec <= b)

    # Singers: flagged in the export or named on a row with a song note. A character
    # who is never heard outside a song (e.g. "Zanger") also sings, but only on rows
    # that directly follow another song row, so speakers inside a long song range
    # are not picked up.
    flagged = {n for n, c in ep.characters.items() if c.singing}
    for s in songs.values():
        s.singers.update(flagged)
    appears_outside: set[str] = set()
    for r in rows:
        sec = tc_seconds(r.tc)
        if not any(in_song(sec, s) for s in songs.values()):
            appears_outside.update(split_character_names(r.character))
    prev_song: Song | None = None
    for r in rows:
        sec = tc_seconds(r.tc)
        names = split_character_names(r.character)
        explicit = parse_song_note(r.note)
        current = songs[explicit.mark] if explicit else next(
            (s for s in songs.values() if in_song(sec, s) and any(n in s.singers for n in names)),
            None,
        )
        if current is None and prev_song is not None and in_song(sec, prev_song) and all(
            n not in appears_outside and not n.startswith(INSERT_TOKEN) for n in names
        ):
            prev_song.singers.update(names)
            current = prev_song
        prev_song = current

    # 4. Song notes: first line of each singer gets the long form, the rest "Song N".
    seen: set[tuple[str, str]] = set()
    for r in rows:
        explicit = parse_song_note(r.note)
        sec = tc_seconds(r.tc)
        names = split_character_names(r.character)
        song = songs[explicit.mark] if explicit else next(
            (s for s in songs.values() if in_song(sec, s) and any(n in s.singers for n in names)),
            None,
        )
        if not song:
            continue
        first = any((song.mark, n) not in seen for n in names)
        seen.update((song.mark, n) for n in names)
        r.note = song_note(song, long=first)

    for s in songs.values():
        if not s.tc_out:
            warnings.append(f"{s.mark}: geen TC Out gevonden in de note.")
        if not s.titles:
            warnings.append(f"{s.mark}: geen titel (Origineel//Vertaald) in de note; "
                            "de titelregel blijft leeg.")
    return rows, warnings


# ---------------------------------------------------------------- statistics


@dataclass
class CharStat:
    name: str
    in_tcs: int = 0
    words: int = 0
    pages: set[int] = field(default_factory=set)
    first_tc: str = ""
    sings: bool = False


@dataclass
class InsertStat:
    name: str
    tc: str
    page: int
    original: str
    translated: str


@dataclass
class SongStat:
    mark: str
    original_title: str
    local_title: str
    characters: list[str]
    words: int
    tc_in: str
    tc_out: str


@dataclass
class Statistics:
    characters: list[CharStat]
    inserts: list[InsertStat]
    songs: list[SongStat]


def _lines(text: str) -> list[str]:
    return [l.strip() for l in text.split("\n")]


def compute_statistics(rows: list[ScriptRow], pages: list[int]) -> Statistics:
    chars: dict[str, CharStat] = {}
    inserts: dict[str, InsertStat] = {}
    songs: dict[str, SongStat] = {}

    for row, page in zip(rows, pages):
        post = row.character.strip()
        if len(post) <= 1:
            continue
        words = count_words(row.text) if row.text.strip() else 0
        has_text = words > 0 and bool(row.text.strip())
        is_song = row.note.startswith(SONG_TOKEN)
        singers = split_character_names(post)

        if is_song:
            nl = _lines(row.note)
            mark = nl[0]
            if mark not in songs:
                titles = nl[1] if len(nl) > 1 else ""
                orig, _, local = (p.strip() for p in titles.partition("//"))
                songs[mark] = SongStat(mark, orig, local, [], words, row.tc,
                                       nl[2] if len(nl) > 2 else "")
            else:
                songs[mark].words += words
            for s in singers:
                if s not in songs[mark].characters:
                    songs[mark].characters.append(s)

        for name in singers:
            c = chars.get(name)
            if c is None:
                c = chars[name] = CharStat(name, first_tc=_lines(row.tc)[0])
                if has_text:
                    c.in_tcs = 1
                    if not is_song:
                        c.words = words
            elif has_text:
                c.in_tcs += 1
                if not is_song:
                    c.words += words
            c.pages.add(page)
            if is_song:
                c.sings = True
            if name.startswith(INSERT_TOKEN) and name not in inserts:
                orig, _, trans = (p.strip() for p in row.text.partition("//"))
                inserts[name] = InsertStat(name, c.first_tc, page, orig, trans)

    people = sorted((c for n, c in chars.items() if not n.startswith(INSERT_TOKEN)),
                    key=lambda c: c.name.lower())
    return Statistics(people, list(inserts.values()), list(songs.values()))


def gender_code(gender: str) -> str:
    g = gender.strip()
    return GENDER_CODES.get(g.lower(), g)
