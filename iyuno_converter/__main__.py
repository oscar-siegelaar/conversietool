"""Command line: python -m iyuno_converter EXPORT.xlsx [-o OUT.docx] [--auth-code CODE].
Without arguments the window opens."""
import argparse
import sys


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    argv = [a for a in argv if not a.startswith("-psn")]  # macOS Finder launch
    if not argv:
        from .gui import run
        return run()
    from .pipeline import convert, default_cover
    from .reader import read_export
    from .settings import load_defaults, load_filename_template
    ap = argparse.ArgumentParser(prog="iyuno-converter")
    ap.add_argument("export", help="export .xlsx, or a folder to convert every export in it")
    ap.add_argument("-o", "--output")
    ap.add_argument("--auth-code", default="")
    ap.add_argument("--language", default="")
    a = ap.parse_args(argv)
    from pathlib import Path
    if Path(a.export).is_dir():
        from .pipeline import convert_folder
        out_dir, report = convert_folder(a.export, a.output, load_defaults(), load_filename_template())
        print("\n".join(report))
        print(out_dir)
        return
    ep = read_export(a.export)
    cover = default_cover(ep, load_defaults())
    if a.auth_code:
        cover.auth_code = a.auth_code
    if a.language:
        cover.language = a.language
    out, warnings = convert(a.export, a.output, cover, ep, load_filename_template())
    for w in warnings:
        print("Let op:", w)
    print(out)


if __name__ == "__main__":
    sys.exit(main())
