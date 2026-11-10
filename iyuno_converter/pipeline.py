"""One call from export file to Word document."""
from __future__ import annotations

from pathlib import Path

from .convert import compute_statistics, transform
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


def convert_folder(folder: str | Path, out_dir: str | Path | None = None,
                   defaults: dict[str, str] | None = None) -> tuple[Path, list[str]]:
    """Convert every .xlsx export in a folder (for test runs on many scripts).

    Cover data comes from each export plus the saved defaults; the SDI auth code
    stays empty. Returns the
    output folder and a report with one line per file.
    """
    folder = Path(folder)
    out_dir = Path(out_dir) if out_dir else folder / "iyuno-word"
    out_dir.mkdir(parents=True, exist_ok=True)
    report: list[str] = []
    for src in sorted(folder.glob("*.xlsx")):
        if src.name.startswith("~$"):  # Excel lock file
            continue
        try:
            ep = read_export(src)
            out, warnings = convert(src, out_dir / (src.stem + ".docx"), default_cover(ep, defaults), ep)
            notes = [w for w in warnings if "AUTH CODE" not in w]
            report.append(f"OK    {src.name}" + (f"  ({'; '.join(notes)})" if notes else ""))
        except Exception as exc:  # keep going, report the failure
            report.append(f"FOUT  {src.name}: {exc}")
    (out_dir / "rapport.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    return out_dir, report
