#!/usr/bin/env bash
# Package dist/Iyuno Converter.app as a drag-to-install DMG (macOS only).
set -euo pipefail
cd "$(dirname "$0")"
out="${1:-dist/Iyuno-Converter.dmg}"
stage="$(mktemp -d)"
cp -R "dist/Iyuno Converter.app" "$stage/"
ln -s /Applications "$stage/Apps"
cat > "$stage/LEES MIJ.txt" <<'TXT'
1. Sleep "Iyuno Converter" naar "Apps".
2. Open de app vanuit Apps.
3. Zegt de Mac dat de app niet geopend kan worden?
   Ga naar Systeeminstellingen > Privacy en beveiliging,
   scrol naar beneden en klik bij "Iyuno Converter" op "Open toch".
   Dit hoeft maar een keer.
TXT
rm -f "$out"
hdiutil create -volname "Iyuno Converter" -srcfolder "$stage" -ov -format UDZO "$out"
rm -rf "$stage"
echo "Built: $out"
