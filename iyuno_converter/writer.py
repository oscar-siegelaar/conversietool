"""Write the Iyuno/SDI Word document from the bundled template."""
from __future__ import annotations

import copy
import datetime as dt
from dataclasses import dataclass
from importlib import resources
from pathlib import Path

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from .convert import Statistics, gender_code, verses
from .reader import CharacterInfo, ScriptRow

STATS_TITLE = "STATISTICS (v. 5) - New HL method + QuickDub"


@dataclass
class CoverInfo:
    original_title: str = ""
    translated_title: str = ""
    series_title: str = ""
    translated_series_title: str = ""
    season: str = ""
    episode: str = ""
    language: str = "Dutch"
    date: str = ""
    translated_by: str = ""
    auth_code: str = ""


# ---------------------------------------------------------------- cell helpers


def _set_cell_lines(cell, text: str) -> None:
    """Fill a cell with one paragraph per line, keeping the cell's formatting.

    The macro reads the parts of a song note by paragraph (vbCr), so each line must
    be its own paragraph, not a soft line break.
    """
    paras = cell.paragraphs
    proto = paras[0]._p
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    run_proto = proto.find(qn("w:r"))
    rpr = copy.deepcopy(run_proto.find(qn("w:rPr"))) if run_proto is not None else None
    for r in proto.findall(qn("w:r")):
        proto.remove(r)

    lines = text.split("\n") if text else [""]
    prev = proto
    for i, line in enumerate(lines):
        p = proto if i == 0 else copy.deepcopy(proto)
        if i:
            for r in p.findall(qn("w:r")):
                p.remove(r)
            prev.addnext(p)
        if line:
            r = OxmlElement("w:r")
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = line
            r.append(t)
            p.append(r)
        prev = p


def _border_xml(tbl, outside_sz=12, inside_sz=4):
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (("top", outside_sz), ("left", outside_sz), ("bottom", outside_sz),
                     ("right", outside_sz), ("insideH", inside_sz), ("insideV", inside_sz)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def _cell_border(cell, edge: str, val: str = "double", sz: int = 4):
    tcPr = cell._tc.get_or_add_tcPr()
    b = tcPr.find(qn("w:tcBorders"))
    if b is None:
        b = OxmlElement("w:tcBorders")
        tcPr.append(b)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "000000")
    b.append(el)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def _stats_table(doc, header: list[str], body: list[list[str]], align_center=True,
                 total: list[str] | None = None):
    rows = body + ([total] if total else [])
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    _border_xml(t)
    for r_i, values in enumerate([header] + rows):
        row = t.rows[r_i]
        for c_i, v in enumerate(values):
            cell = row.cells[c_i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if align_center or r_i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(v))
            run.font.name = "Arial"
            run.font.size = Pt(10)
            run.bold = r_i == 0 or (total is not None and r_i == len(rows))
            if r_i == 0:
                _cell_border(cell, "bottom")
            if total is not None and r_i == len(rows):
                _cell_border(cell, "top")
    _repeat_header(t.rows[0])
    t.autofit = True
    return t


def _heading(doc, text, size, center=False, space_before=True):
    if space_before:
        doc.add_paragraph()
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(size)
    r.bold = True
    return p


# ---------------------------------------------------------------- main writer


def write_document(
    out_path: str | Path,
    cover: CoverInfo,
    rows: list[ScriptRow],
    stats: Statistics,
    characters: dict[str, CharacterInfo],
    now: dt.datetime | None = None,
) -> Path:
    now = now or dt.datetime.now()
    with resources.files(__package__).joinpath("template.docx").open("rb") as fh:
        doc = docx.Document(fh)

    # Cover table: find the value cell by its label.
    labels = {
        "ORIGINAL TITLE": cover.original_title,
        "TRANSLATED TITLE": cover.translated_title,
        "SERIES TITLE": cover.series_title,
        "TRANSLATED SERIES TITLE": cover.translated_series_title,
        "SEASON#": cover.season,
        "EPISODE#": cover.episode,
        "LANGUAGE": cover.language,
        "DATE": cover.date or f"{now.day}-{now.month}-{now.year}",
        "TRANSLATED BY": cover.translated_by,
        "SDI AUTH CODE": cover.auth_code.replace(" ", ""),
    }
    for row in doc.tables[0].rows:
        key = row.cells[0].text.strip().upper().replace(" #", "#")
        if key in labels:
            _set_cell_lines(row.cells[1], labels[key])

    # Script table: header + one prototype row in the template.
    script = doc.tables[-1]
    proto = script.rows[1]._tr
    _repeat_header(script.rows[0])
    anchor = proto
    for r in rows:
        tr = copy.deepcopy(proto)
        anchor.addnext(tr)
        anchor = tr
        cells = tr.findall(qn("w:tc"))
        from docx.table import _Cell
        for tc, value in zip(cells, (r.character, r.tc, r.note, r.text)):
            _set_cell_lines(_Cell(tc, script), value)
    proto.getparent().remove(proto)

    # Statistics in a new section, like the macro.
    sect = doc.add_section(WD_SECTION.NEW_PAGE)
    for el in sect._sectPr.findall(qn("w:pgNumType")):  # continue page numbers
        sect._sectPr.remove(el)
    _heading(doc, STATS_TITLE, 14, center=True, space_before=False)
    doc.add_paragraph()

    body, tot_tcs, tot_verses = [], 0, 0
    for c in stats.characters:
        info = characters.get(c.name)
        v = verses(c.words)
        tot_tcs += c.in_tcs
        tot_verses += v
        body.append([
            c.name,
            (info.original_name if info and info.original_name else c.name),
            gender_code(info.gender) if info else "",
            "",
            c.in_tcs,
            v,
            ", ".join(str(p) for p in sorted(c.pages)),
            c.first_tc,
            "Yes" if c.sings else "",
        ])
    _stats_table(
        doc,
        ["Script name", "Original name", "Gender", "Actor", "In TCs", "Verses", "Pages",
         "First TC In", "Sings"],
        body,
        total=["", "", "", "Total", tot_tcs, tot_verses, "", "", ""],
    )

    if stats.inserts:
        _heading(doc, "List of inserts:", 12)
        _stats_table(
            doc,
            ["Insert number", "TC", "Page", "Original insert", "Translated insert"],
            [[i.name, i.tc, i.page, i.original, i.translated] for i in stats.inserts],
        )

    if stats.songs:
        _heading(doc, "List of songs:", 12)
        _stats_table(
            doc,
            ["Song mark", "Original song title", "Local song title", "Characters", "Verses",
             "TC In", "TC Out"],
            [[s.mark, s.original_title, s.local_title, ", ".join(s.characters),
              verses(s.words), s.tc_in, s.tc_out] for s in stats.songs],
        )

    _heading(doc, "List of banned words:", 12)
    _stats_table(doc, ["Banned word", "TC"], [["", ""]])

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Statistics timestamp: " + now.strftime("%Y-%m-%d %H:%M:%S"))
    r.font.name = "Arial"
    r.font.size = Pt(8)

    out_path = Path(out_path)
    doc.save(out_path)
    return out_path
