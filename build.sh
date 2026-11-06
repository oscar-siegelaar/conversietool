#!/usr/bin/env bash
# Build the macOS app for the architecture of this Mac.
set -euo pipefail
cd "$(dirname "$0")"
python3 -m pip install -q openpyxl python-docx tkinterdnd2 pyinstaller
python3 -m PyInstaller --noconfirm --windowed --name "Iyuno Converter" \
  --add-data "iyuno_converter/template.docx:iyuno_converter" \
  --collect-all tkinterdnd2 \
  run_app.py
echo "Built: dist/Iyuno Converter.app"
