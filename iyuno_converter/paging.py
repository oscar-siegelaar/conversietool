"""Estimate on which page each script row lands.

Word decides the real page breaks. Without Word we estimate them from the template's
layout: column widths, Calibri 11 pt, 1.5 line spacing, cell padding and the usable
page height. The estimate is tuned against LibreOffice's rendering of the same file.
"""
from __future__ import annotations

from .reader import ScriptRow

TWIP_PER_PT = 20

# Template layout (twips) – see template.docx, script table and section 2.
PAGE_HEIGHT = 16838
TOP_MARGIN = 794
BOTTOM_MARGIN = 567
HEADER_ALLOWANCE = 0          # header sits inside the top margin
CELL_PADDING = 57 * 2         # top + bottom cell margin
CELL_SIDE_PADDING = 108 * 2
COLUMN_WIDTHS = (1906, 1276, 1134, 5387)  # Character, TC, Note, TEXT
LINE_HEIGHT = 22.0 * TWIP_PER_PT       # Calibri 11 pt at 1.5 lines (measured)
AVG_CHAR_WIDTH = 5.0 * TWIP_PER_PT      # Calibri 11 pt, mixed text (measured)
HEADER_ROW = LINE_HEIGHT + CELL_PADDING

# Page 1 of the script section starts with the script table at the top.
FIRST_PAGE_OFFSET = 0


def _wrapped_lines(text: str, width_twips: float) -> int:
    usable = max(width_twips - CELL_SIDE_PADDING, AVG_CHAR_WIDTH * 4)
    per_line = max(int(usable // AVG_CHAR_WIDTH), 1)
    total = 0
    for para in text.split("\n") or [""]:
        if not para:
            total += 1
            continue
        # greedy word wrap
        lines, cur = 1, 0
        for word in para.split(" "):
            w = len(word)
            if cur == 0:
                cur = w
            elif cur + 1 + w <= per_line:
                cur += 1 + w
            else:
                lines += 1
                cur = w
            while cur > per_line:   # very long word
                lines += 1
                cur -= per_line
        total += lines
    return total


def row_height(row: ScriptRow) -> float:
    cells = (row.character, row.tc, row.note, row.text)
    lines = max(_wrapped_lines(c, w) for c, w in zip(cells, COLUMN_WIDTHS))
    return lines * LINE_HEIGHT + CELL_PADDING


def estimate_pages(rows: list[ScriptRow], start_page: int = 1) -> list[int]:
    usable = PAGE_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN - HEADER_ALLOWANCE
    page = start_page
    y = FIRST_PAGE_OFFSET + HEADER_ROW
    result = []
    for r in rows:
        h = row_height(r)
        # Rows may split across pages in the template; Word puts the start of the
        # row on the page where it begins, which is what the macro reads.
        if y + min(h, LINE_HEIGHT + CELL_PADDING) > usable:
            page += 1
            y = HEADER_ROW  # header row repeats on every page
        result.append(page)
        y += h
        while y > usable:   # row continues on the next page
            page += 1
            y = HEADER_ROW + (y - usable)
    return result
