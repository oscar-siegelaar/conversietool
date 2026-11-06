"""One call from export file to Word document."""
from __future__ import annotations

from pathlib import Path

from .convert import compute_statistics, transform
from .paging import estimate_pages
from .reader import Episode, read_export
from .writer import CoverInfo, write_document


def default_cover(ep: Episode) -> CoverInfo:
    return CoverInfo(
        original_title=ep.original_title,
        translated_title=ep.translated_title,
        series_title=ep.series_title,
        season=ep.season,
        episode=ep.episode,
        translated_by=ep.translator,
    )


def default_output_name(ep: Episode, language: str = "Dutch") -> str:
    series = "".join(w[0] for w in ep.series_title.replace("&", " ").split()).upper()
    parts = [language.upper(), series or "SCRIPT", ep.episode or "XX"]
    return "_".join(parts) + ".docx"


def convert(src: str | Path, out: str | Path | None = None, cover: CoverInfo | None = None,
            ep: Episode | None = None) -> tuple[Path, list[str]]:
    src = Path(src)
    ep = ep or read_export(src)
    rows, warnings = transform(ep)
    cover = cover or default_cover(ep)
    pages = estimate_pages(rows)
    stats = compute_statistics(rows, pages)
    out = Path(out) if out else src.with_name(default_output_name(ep, cover.language))
    write_document(out, cover, rows, stats, ep.characters)
    if not cover.auth_code:
        warnings.append("SDI AUTH CODE is leeg: vul de code van Iyuno in.")
    return out, warnings
