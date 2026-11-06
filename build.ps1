# Build the Windows app.
Set-Location $PSScriptRoot
python -m pip install -q openpyxl python-docx tkinterdnd2 pyinstaller
python -m PyInstaller --noconfirm --windowed --onefile --name "Iyuno Converter" `
  --add-data "iyuno_converter/template.docx;iyuno_converter" `
  --collect-all tkinterdnd2 `
  run_app.py
Write-Host "Built: dist\Iyuno Converter.exe"
