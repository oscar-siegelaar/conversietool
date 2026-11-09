"""Tests use a small made-up export, so no client scripts live in the repo."""
from pathlib import Path

import docx
import openpyxl
import pytest

from iyuno_converter.convert import compute_statistics, count_words, split_character_names, transform
from iyuno_converter.pipeline import convert, default_cover
from iyuno_converter.reader import read_export

SCRIPT = [
    ("Insert 1", "00:00:00", None, "Original Title // Vertaalde titel"),
    ("Insert", "00:00:10", None, "Bakkerij"),
    ("Anna", "00:00:12", None, "Hallo, hoe gaat het met je vandaag?"),
    ("Ben", "00:00:15", None, "Goed! (LACHT) En met jou?\nPrima."),
    ("Koor", "00:01:00", "Song 1 Big Song//Groot Lied 01:30", "La la la\nLa la la"),
    ("Zanger", "00:01:05", None, "Ik zing een lied\nvoor jou"),
    ("Koor", "00:01:10", None, "Aaaa"),
    ("Insert", "00:01:20", None, "Ziekenhuis"),
    ("Anna+Ben", "00:01:40", None, "Samen!"),
    ("Synopsis", "00:02:00", None, None),
]
CHARS = [
    ("character", "count", "first seen", "original name", "description", "age", "gender", "singing"),
    ("Anna", 2, "00:12", "Anne", "", "", "Female", None),
    ("Ben", 2, "00:15", "Ben", "", "", "Male", None),
    ("Koor", 2, "01:00", "Chorus", "", "", "Male-Female", None),
    ("Zanger", 1, "01:05", "Singer", "", "", "Male", None),
]
META = [
    ("metadata", "part", "song", "translation", "original"),
    ("Episode title", "A", None, "Vertaalde titel", "Original Title"),
    ("Translator", "A", None, "Test Vertaler", None),
]


@pytest.fixture
def export(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "script"
    ws.append(("Character", "TC", "Note", "TEXT"))
    for r in SCRIPT:
        ws.append(r)
    for title, rows in (("character count", CHARS), ("metadata", META)):
        s = wb.create_sheet(title)
        for r in rows:
            s.append(r)
    p = tmp_path / "My Show S02E07 Pilot.xlsx"
    wb.save(p)
    return p


def test_word_count_matches_macro_method_5():
    assert count_words("Hallo, hoe gaat het?") == 4
    assert count_words("(LACHT) Ja - nee... oké!") == 4
    assert count_words("one  two   three") == 3


def test_split_names():
    assert split_character_names("Anna+Ben") == ["Anna", "Ben"]
    assert split_character_names("Anna, Ben,") == ["Anna", "Ben"]
    assert split_character_names("Kevin:") == ["Kevin"]


def test_reader_metadata_and_filename(export):
    ep = read_export(export)
    assert ep.original_title == "Original Title"
    assert ep.translator == "Test Vertaler"
    assert (ep.series_title, ep.season, ep.episode) == ("My Show", "2", "7")
    assert ep.characters["Koor"].original_name == "Chorus"


def test_transform_inserts_and_songs(export):
    ep = read_export(export)
    rows, warnings = transform(ep)
    assert not warnings
    assert rows[0].character == "Insert 1" and rows[0].text == "Bakkerij"
    assert [r.character for r in rows if r.character.startswith("Insert")] == ["Insert 1", "Insert 2"]
    assert all(r.character != "Synopsis" for r in rows)
    notes = {(r.character, r.tc): r.note for r in rows}
    long_note = "Song 1\nBig Song//Groot Lied\n01:30"
    assert notes[("Koor", "00:01:00")] == long_note
    assert notes[("Zanger", "00:01:05")] == long_note      # first line of this singer
    assert notes[("Koor", "00:01:10")] == "Song 1"
    assert notes[("Anna+Ben", "00:01:40")] == ""           # after TC out


def test_statistics(export):
    ep = read_export(export)
    rows, _ = transform(ep)
    stats = compute_statistics(rows, [1] * len(rows))
    by = {c.name: c for c in stats.characters}
    assert set(by) == {"Anna", "Ben", "Koor", "Zanger"}
    assert by["Anna"].in_tcs == 2 and by["Anna"].words == 8   # 7 + 1 words
    assert by["Koor"].words == 0 and by["Koor"].sings          # songs not summed
    assert stats.songs[0].characters == ["Koor", "Zanger"]
    assert stats.songs[0].words == 6 + 6 + 1
    assert [i.name for i in stats.inserts] == ["Insert 1", "Insert 2"]


def test_full_document(export, tmp_path):
    ep = read_export(export)
    cover = default_cover(ep)
    cover.auth_code = "ABC 123"
    out, _ = convert(export, tmp_path / "out.docx", cover, ep)
    d = docx.Document(out)
    front = {r.cells[0].text.strip(): r.cells[1].text for r in d.tables[0].rows}
    assert front["ORIGINAL TITLE"] == "Original Title"
    assert front["SDI AUTH CODE"] == "ABC123"
    script = d.tables[1]
    assert [c.text for c in script.rows[0].cells] == ["Character", "TC", "Note", "TEXT"]
    assert len(script.rows) == 1 + 8
    note_cell = [r for r in script.rows if r.cells[1].text == "00:01:00"][0].cells[2]
    assert len(note_cell.paragraphs) == 3          # one paragraph per line, as the macro needs
    heads = [t.rows[0].cells[0].text for t in d.tables[2:]]
    assert heads == ["Script name", "Insert number", "Song mark", "Banned word"]
    chars = {r.cells[0].text: [c.text for c in r.cells] for r in d.tables[2].rows[1:]}
    assert chars["Anna"][1:3] == ["Anne", "V"]
    assert chars["Koor"][2] == "MV"


def test_convert_folder(export, tmp_path):
    from iyuno_converter.pipeline import convert_folder
    (export.parent / "broken.xlsx").write_text("not excel")
    out_dir, report = convert_folder(export.parent, tmp_path / "out")
    assert (out_dir / (export.stem + ".docx")).exists()
    assert (out_dir / "rapport.txt").exists()
    assert sorted(r.split()[0] for r in report) == ["FOUT", "OK"]


@pytest.mark.parametrize("note, expected", [
    ("Song 1 Big Song//Groot Lied 01:30", ("Song 1", "Big Song//Groot Lied", "01:30")),
    ("song 4 30:49", ("Song 4", "", "30:49")),
    ("{song 4 30:49}", ("Song 4", "", "30:49")),
    ("SONG 2", ("Song 2", "", "")),
    ("Song 3\nOrig//Lokaal\n1:02:03", ("Song 3", "Orig//Lokaal", "1:02:03")),
    ("Songtekst nakijken", None),
])
def test_parse_song_note_variants(note, expected):
    from iyuno_converter.convert import parse_song_note
    s = parse_song_note(note)
    assert (s and (s.mark, s.titles, s.tc_out)) == expected if expected else s is None


def test_lowercase_song_note_without_titles(tmp_path):
    """Real case: SyncScript export with 'song 4 30:49' first and 'Song 4' after."""
    from iyuno_converter.reader import Episode, ScriptRow
    ep = Episode(rows=[
        ScriptRow("Ellian", "00:26:44", "", "(ZUCHT)"),
        ScriptRow("Ellian", "00:27:29", "song 4 30:49", "Muren kaal, lege gangen"),
        ScriptRow("Ellian", "00:27:52", "Song 4", "Elluk feest, de gesprekken"),
        ScriptRow("Ellian", "00:30:40", "", "(BLAZEN)"),
    ])
    rows, warnings = transform(ep)
    assert rows[1].note == "Song 4\n\n30:49"       # three paragraphs, empty title line
    assert rows[2].note == "Song 4"
    assert rows[3].note == ""
    assert not any("TC Out" in w for w in warnings)
    assert any("titel" in w for w in warnings)
    stats = compute_statistics(rows, [1] * len(rows))
    assert stats.songs[0].tc_out == "30:49" and stats.songs[0].characters == ["Ellian"]


def test_series_from_broken_filenames():
    from iyuno_converter.reader import Episode, _guess_from_filename

    cases = {
        "Show A SShow AE01 Some Title": ("Show A", "", "1"),
        "Project Name SOther Series 4E02 TITLE": ("Other Series 4", "", "2"),
        "My Show S02E05 Title": ("My Show", "2", "5"),
        "Sam Show SSam ShowE03 x": ("Sam Show", "", "3"),
    }
    for stem, expected in cases.items():
        ep = Episode()
        _guess_from_filename(ep, stem)
        assert (ep.series_title, ep.season, ep.episode) == expected, stem
