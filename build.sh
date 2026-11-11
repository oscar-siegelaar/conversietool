#!/usr/bin/env bash
# Build the macOS app for the architecture of this Mac.
set -euo pipefail
cd "$(dirname "$0")"
python3 -m pip install -q openpyxl python-docx tkinterdnd2 pyinstaller
python3 -m PyInstaller --noconfirm --windowed --name "Iyuno Converter" \
  --add-data "iyuno_converter/template.docx:iyuno_converter" \
  --collect-all tkinterdnd2 \
  run_app.py
# Put the app version into the bundle so Finder's Get Info shows it.
APP="dist/Iyuno Converter.app"
VERSION=$(python3 -c "import iyuno_converter as m; print(m.__version__)")
plutil -replace CFBundleShortVersionString -string "$VERSION" "$APP/Contents/Info.plist"
plutil -replace CFBundleVersion -string "$VERSION" "$APP/Contents/Info.plist"
codesign --force --deep --sign - "$APP"
echo "Built: $APP (version $VERSION)"
