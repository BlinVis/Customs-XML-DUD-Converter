# Perkthim → DUD XML converter

Desktop app (Python 3.10+, Tkinter, openpyxl) for a Kosovo customs agent. It turns a supplier's
**perkthim** Excel (translated invoice lines) into the item XML that the **old ASYCUDA client**
imports, replacing a manual routine done in Excel with the "XML FORMA" template
(`tests/fixtures/XML FORMA.xltx`, which has the XML map). Users are customs staff with weak Excel
skills, so the UI must stay simple, step-by-step and bilingual (Albanian default, English).

## Layout

- `dud_converter/core.py`: all conversion logic, no GUI. `convert(path, Options) -> Result`,
  `inspect_sheet()` for live column feedback, `build_xml()`, `output_filename()`.
- `dud_converter/gui.py`: Tkinter window (clam theme, scrolling body): header with Shqip/English
  switch, steps 1 file, 2 columns (combobox of sheet titles + ✓/✗), 3 questions (00/90, 1000/4000
  or a typed code, MSA/GBMTL), 4 how to create/check, fixed footer with the Create button, result
  popup. Create always opens a Save-as dialog (starts in the last folder used). The wheel over a column
  combobox scrolls the page instead of changing the pick (`ScrollArea.scroll_only`). Settings (column
  names, MSA words, fixed values, language, last output folder) are saved in
  `%APPDATA%\DUDConverter\settings.json`.
- `dud_converter/i18n.py`: every user-visible string, `{"sq": ..., "en": ...}`; `t(lang, key, **p)`.
  Errors/warnings from core are keys + params, translated in the GUI. A test checks both languages.
- `app.py` entry point; `build_exe.bat` (PyInstaller one-file); `run_app.bat`.
- `.github/workflows/build.yml`: pytest on Ubuntu, then builds `DUD-Converter.exe` on Windows and
  uploads it as the **DUD-Converter-windows** artifact on every push/PR. That's how the user gets the
  exe; several runs have artifacts with the same name, so link the specific run.

## Conversion rules (agreed with the user, verified against their hand-made DUDs)

- Columns are found by the **exact header text, case-insensitive, extra spaces ignored**; the user
  types/picks the names (supplier sheets name columns inconsistently). The header row is the first row
  (any sheet, first 30 rows) holding the most of the names. **Every column is optional** (user's
  choice): a missing one is skipped and its element left empty in every item, with one
  `warn_column_missing` warning (none for `Kodi`/`Pref`). Only a sheet with none of the names errors.
- Rows with neither perkthim nor article code (e.g. the totals row) are skipped.
- `Commercial_Description` = perkthim text **exactly as typed** (spaces kept, like Excel CONCAT; only
  line breaks flattened) + `=` + quantity + `PCE`. If the text doesn't start with the article code,
  prefix `"<kodi> - "`. The user's DUDs use both `PCE` and `pce`; app writes `PCE`.
- `Amount_foreign_currency` = **Totali only, never "Totali me zbritje"**, rounded to 2 decimals.
  Formulas (`=E2*D2`) use Excel's cached value, else a small safe evaluator (+-*/, SUM, ROUND).
  Text numbers like `1.234,56 €` are parsed.
- `Precision_1` (00/90) and `Extended_customs_procedure` (1000/4000) are asked every time, no default;
  the user can also type any other 2-digit / 4-digit code.
- Origin codes are upper-cased and `UK` is written as `GB` (the old ASYCUDA has no UK).
- **MSA** (`Preference_code`): only for EU origin countries; GB origins get `GBMTL` instead (user's
  rule, "I think"), handled the same way (indicator column or ticked). With an indicator column, positive values
  (default list: Po, Yes, Y, Da, Ja, +, X, MSA, True, 1) on an EU origin → `MSA`; positive on non-EU →
  warning, no MSA. Without the column, the user ticks which EU origins in that invoice get MSA (the
  user's choice; in the Bosch sample BG got MSA but DE didn't, so it can't be inferred). Not remembered.
- Fixed values ("dragged down" in the template, editable under ⚙): Number_of_packages 0,
  Kind_of_packages_code PA, Kind_of_packages_name Paketa, National_customs_procedure 000,
  Valuation_method_code 1, Rate_of_adjustement 1, Currency_code EUR,
  Currency_name "Nuk ka valutë të huaj", Currency_rate 1.
- Always empty/omitted (the user deletes them in the template): Marks2_of_packages,
  Summary_declaration (so `<Previous_doc/>` is empty), Statistical_value.
- Output name `DUDI-<excel file stem>-<DDMMYYYY>.xml`, saved next to the Excel unless a folder is chosen.

## XML format (must stay byte-compatible with Excel's export)

`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, root
`<ASYCUDA xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`, one `<Item>` per row, **tab
indentation, CRLF line endings, no trailing newline**, empty elements omitted (Preference_code only on
MSA rows), numbers without trailing zeros (`478.8`, `99`). Element order is in `build_xml()`.

## Testing

`pip install -r requirements.txt pytest && pytest`. The key tests regenerate the user's real DUDs
from the sample Excels byte for byte: `DUDIBILSTEIN17102026.xml` (Bilstein, has a `Pref` Po/Jo column)
and `DUDIBOSCH22082026.xml` (Bosch, no Pref column, MSA on BG). `tests/fixtures/*` is marked binary in
`.gitattributes`; don't let line endings change. Any rule change should be checked against both.

GUI check without Windows: Tk isn't in the default python here; `apt-get install python3-tk`, make a
python3.12 venv with `--system-site-packages`, `pip install openpyxl`, run under `xvfb-run` and
screenshot with `import -window root`. Patch `messagebox` functions in the script, and delete any
earlier output file first, or the overwrite prompt blocks.

## Working with the user

- Blin writes in English; the app's default language is Albanian. Keep all UI text in `i18n.py`.
- Repo is private (`BlinVis/Customs-XML-DUD-Converter`, default branch `main`); changes go through PRs.
- When the user sends a perkthim + their hand-made DUD + the app's DUD, compare element by element,
  first checking all three come from the same invoice (filenames can mismatch).
