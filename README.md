# Iyuno converter

A small stand-alone app for macOS and Windows. You give it a SyncScript Excel export
(the SDI/Iyuno export, type 5) and it gives you the Word document that Iyuno's system
reads in. It does the work of the Word macro "SDI Template v. 5" and of the manual steps
translators do today, so neither Word nor the macro is needed.

## What it does

1. **Cover page.** Fills in the original and translated title, series, season, episode,
   language, date, translator and the **SDI auth code**. Iyuno issues the auth code per
   title and translator. The app never creates one: the user types it in.
2. **Script table** (Character, TC, Note, TEXT):
   - drops the title row at 00:00:00 (the title goes on the cover);
   - numbers the inserts `Insert 1`, `Insert 2`, ...;
   - songs: the first line of each singer gets the long note (`Song 1` ⏎
     `Original//Local` ⏎ `TC out`), and every later sung line gets `Song 1`. Each part
     is its own paragraph, which the macro and Iyuno's system need;
   - drops label-only rows such as `Synopsis`.
3. **Statistics** in a new section, exactly like the macro: characters (with original
   name and gender from the export's *character count* sheet, In TCs, verses, pages,
   first TC, sings), inserts, songs, banned words and a timestamp. Verses use the
   macro's method 5: punctuation is stripped, words are counted, and verses are
   words / 8 rounded up. Sung lines do not count toward a character's verses.

### Known limits

- **Pages** are estimated from the template layout, because only Word decides the
  real page breaks. Compared with a LibreOffice rendering of the same file, 312 of
  319 rows land on the right page. The rest are one page off at a page boundary.
- **Gender and original name** in SyncScript are often defaults. The window lists
  every character so you can correct them before saving.
- The output is `.docx`. The old macro template is a Word 97 `.doc`.

## Use

The window has a **Convert whole folder…** button for test runs. It converts every
`.xlsx` in a folder into a subfolder `iyuno-word` and writes `rapport.txt` with one line
per file.

```bash
pip install -e .            # or: pip install openpyxl python-docx
python -m iyuno_converter                 # opens the window
python -m iyuno_converter export.xlsx --auth-code CODE -o out.docx   # command line
python -m iyuno_converter path/to/folder   # every export in a folder, plus rapport.txt
```

For drag and drop in the window, also install `tkinterdnd2` (`pip install -e .[dnd]`).
Without it, click the box to choose a file.

## Build the apps

```bash
./build.sh        # macOS: dist/Iyuno Converter.app (built for the Mac it runs on)
build.ps1         # Windows: dist\Iyuno Converter.exe
```

An Apple Silicon Mac builds an arm64 app, and an Intel Mac (for example the iMac on
Ventura) builds an x86_64 app. The apps are not signed. On first launch on a Mac,
right-click the app and choose **Open**. The workflow
`.github/workflows/build.yml` builds all three: run it by hand from the Actions tab (Run workflow).

## Tests

```bash
pip install pytest && python -m pytest tests
```

The tests use a small made-up export, so no client scripts are stored in the repo.
`iyuno_converter/template.docx` is the SDI template with the macros and personal data
removed.
