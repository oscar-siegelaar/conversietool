"""One call from export file to Word document."""
from __future__ import annotations

from pathlib import Path

from .convert import compute_statistics, transform
from .naming import render, unique_name
from .paging import estimate_pages
from .reader import Episode, read_export
from .writer import CoverInfo, write_document


def default_cover(ep: Episode, defaults: dict[str, str] | None = None) -> CoverInfo:
    """Cover data from the export. A filled-in saved default (Settings) wins."""
    cover = CoverInfo(
        original_title=ep.original_title,
        translated_title=ep.translated_title,
        series_title=ep.series_title,
        translated_series_title=ep.series_title,
        season=ep.season,
        episode=ep.episode,
        translated_by=ep.translator,
    )
    for key, value in (defaults or {}).items():
        if value and hasattr(cover, key):
            setattr(cover, key, value)
    return cover


def output_name(cover: CoverInfo, template: str | None = None) -> str:
    """File name for the Word document, from the Settings › Document name template."""
    return render(template or "", cover) + ".docx"


def convert(src: str | Path, out: str | Path | None = None, cover: CoverInfo | None = None,
            ep: Episode | None = None, template: str | None = None) -> tuple[Path, list[str]]:
    src = Path(src)
    ep = ep or read_export(src)
    rows, warnings = transform(ep)
    cover = cover or default_cover(ep)
    pages = estimate_pages(rows)
    stats = compute_statistics(rows, pages)
    out = Path(out) if out else src.with_name(output_name(cover, template))
    write_document(out, cover, rows, stats, ep.characters)
    if not cover.auth_code:
        warnings.append("SDI AUTH CODE is leeg: vul de code van Iyuno in.")
    return out, warnings


def convert_folder(folder: str | Path, out_dir: str | Path | None = None,
                   defaults: dict[str, str] | None = None,
                   template: str | None = None) -> tuple[Path, list[str]]:
    """Convert every .xlsx export in a folder.

    Cover data comes from each export plus the saved defaults (including the SDI auth
    code when it is saved). Each document is named with the same file name template,
    filled in with that export's data; a name that is already taken gets _2, _3, ...
    Returns the output folder and a report with one line per file.
    """
    folder = Path(folder)
    out_dir = Path(out_dir) if out_dir else folder / "iyuno-word"
    out_dir.mkdir(parents=True, exist_ok=True)
    report: list[str] = []
    taken: set[str] = set()
    for src in sorted(folder.glob("*.xlsx")):
        if src.name.startswith("~$"):  # Excel lock file
            continue
        try:
            ep = read_export(src)
            cover = default_cover(ep, defaults)
            stem = unique_name(output_name(cover, template)[: -len(".docx")], taken)
            out, warnings = convert(src, out_dir / (stem + ".docx"), cover, ep)
            report.append(f"OK    {src.name}  ->  {out.name}"
                          + (f"  ({'; '.join(warnings)})" if warnings else ""))
        except Exception as exc:  # keep going, report the failure
            report.append(f"FOUT  {src.name}: {exc}")
    (out_dir / "rapport.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    return out_dir, report
