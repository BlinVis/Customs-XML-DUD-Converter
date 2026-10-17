"""Albanian (sq) and English (en) texts for the app."""

LANGUAGES = {"sq": "Shqip", "en": "English"}

TEXTS = {
    "title": {"sq": "Perkthim → DUD XML (ASYCUDA)", "en": "Perkthim → DUD XML (ASYCUDA)"},
    "language": {"sq": "Gjuha", "en": "Language"},
    "excel_file": {"sq": "Skedari Excel (perkthimi)", "en": "Excel file (perkthim)"},
    "browse": {"sq": "Zgjidh…", "en": "Browse…"},
    "columns_section": {
        "sq": "Emrat e kolonave në Excel (saktë, pa dallim shkronjash të mëdha/vogla)",
        "en": "Column names in the Excel file (exact, not case sensitive)",
    },
    "optional": {"sq": "(opsionale)", "en": "(optional)"},
    "field_kodi": {"sq": "Emri i kolonës së kodit të artikullit", "en": "Name of the article code column"},
    "field_perkthim": {"sq": "Emri i kolonës së përkthimit", "en": "Name of the translation (perkthim) column"},
    "field_sasia": {"sq": "Emri i kolonës së sasisë", "en": "Name of the quantity column"},
    "field_totali": {"sq": "Emri i kolonës Totali (pa zbritje)", "en": "Name of the Totali column (not the discounted one)"},
    "field_pesha_bruto": {"sq": "Emri i kolonës së peshës bruto", "en": "Name of the gross weight column"},
    "field_pesha_neto": {"sq": "Emri i kolonës së peshës neto", "en": "Name of the net weight column"},
    "field_origjina": {"sq": "Emri i kolonës së vendit të origjinës", "en": "Name of the country of origin column"},
    "field_pref": {"sq": "Emri i kolonës së preferencës (MSA)", "en": "Name of the preference (MSA) column"},
    "field_kodi_tarifor": {"sq": "Emri i kolonës së kodit tarifor", "en": "Name of the tariff code column"},
    "questions_section": {"sq": "Pyetje", "en": "Questions"},
    "q_precision": {"sq": "Precision_1: 00 apo 90?", "en": "Precision_1: 00 or 90?"},
    "q_extended": {
        "sq": "Extended_customs_procedure: 1000 apo 4000?",
        "en": "Extended_customs_procedure: 1000 or 4000?",
    },
    "q_indicators": {
        "sq": "Cilat vlera në kolonën e preferencës do të thonë MSA? (të ndara me presje)",
        "en": "Which values in the preference column mean MSA? (comma separated)",
    },
    "indicators_hint": {
        "sq": "MSA vendoset vetëm kur treguesi është pozitiv DHE vendi i origjinës është në BE.",
        "en": "MSA is set only when the indicator is positive AND the country of origin is in the EU.",
    },
    "q_msa_countries": {
        "sq": "Pa kolonë preference: cilat vende origjine marrin MSA? (p.sh. BG, DE; bosh = asnjë)",
        "en": "No preference column: which origin countries get MSA? (e.g. BG, DE; empty = none)",
    },
    "output_folder": {"sq": "Dosja ku ruhet XML", "en": "Folder to save the XML"},
    "output_same": {"sq": "Lëreni bosh për ta ruajtur pranë skedarit Excel.", "en": "Leave empty to save it next to the Excel file."},
    "constants": {"sq": "Vlerat fikse…", "en": "Fixed values…"},
    "constants_title": {"sq": "Vlerat fikse të çdo artikulli", "en": "Fixed values for every item"},
    "save": {"sq": "Ruaj", "en": "Save"},
    "reset": {"sq": "Kthe parazgjedhjet", "en": "Reset to defaults"},
    "cancel": {"sq": "Anulo", "en": "Cancel"},
    "create": {"sq": "Krijo XML", "en": "Create XML"},
    "show_headers": {"sq": "Shfaq kolonat e Excel-it", "en": "Show Excel columns"},
    "log": {"sq": "Rezultati", "en": "Result"},
    "excel_filter": {"sq": "Skedarë Excel", "en": "Excel files"},
    "all_files": {"sq": "Të gjithë skedarët", "en": "All files"},
    "need_file": {"sq": "Zgjidhni fillimisht skedarin Excel.", "en": "Choose the Excel file first."},
    "need_precision": {"sq": "Zgjidhni Precision_1: 00 apo 90.", "en": "Choose Precision_1: 00 or 90."},
    "need_extended": {
        "sq": "Zgjidhni Extended_customs_procedure: 1000 apo 4000.",
        "en": "Choose Extended_customs_procedure: 1000 or 4000.",
    },
    "overwrite": {"sq": "Skedari {name} ekziston. Ta zëvendësoj?", "en": "{name} already exists. Replace it?"},
    "error": {"sq": "Gabim", "en": "Error"},
    "done": {"sq": "U krye", "en": "Done"},
    "saved": {"sq": "XML u ruajt:\n{path}", "en": "XML saved:\n{path}"},
    "summary": {
        "sq": "Fleta \"{sheet}\", rreshti i titujve {row}.\n"
              "Artikuj: {items}   |   me MSA: {msa}\n"
              "Totali: {amount}   |   Pesha bruto: {gross}   |   Pesha neto: {net}\n"
              "Krahasoni këto totale me rreshtin e totaleve në Excel.",
        "en": "Sheet \"{sheet}\", header row {row}.\n"
              "Items: {items}   |   with MSA: {msa}\n"
              "Total: {amount}   |   Gross weight: {gross}   |   Net weight: {net}\n"
              "Compare these totals with the totals row in the Excel file.",
    },
    "indicators_found": {
        "sq": "Vlerat në kolonën e preferencës: {values}",
        "en": "Values in the preference column: {values}",
    },
    "positive": {"sq": "MSA", "en": "MSA"},
    "negative": {"sq": "jo MSA", "en": "no MSA"},
    "warnings": {"sq": "Kujdes:", "en": "Warnings:"},
    "headers_in": {"sq": "Fleta \"{sheet}\" (rreshti {row}): {headers}", "en": "Sheet \"{sheet}\" (row {row}): {headers}"},
    # Warnings
    "warn_empty": {"sq": "Rreshti {row}: kolona \"{column}\" është bosh.", "en": "Row {row}: column \"{column}\" is empty."},
    "warn_no_pref_column": {
        "sq": "Nuk ka kolonë preference dhe nuk u dha asnjë vend për MSA, prandaj asnjë artikull nuk ka MSA.",
        "en": "There is no preference column and no MSA countries were given, so no item gets MSA.",
    },
    "warn_msa_country_not_eu": {
        "sq": "{country} nuk është vend i BE-së, prandaj nuk merr MSA.",
        "en": "{country} is not an EU country, so it does not get MSA.",
    },
    "warn_pref_not_eu": {
        "sq": "Rreshti {row}: treguesi është pozitiv por origjina {country} nuk është në BE, MSA nuk u vendos.",
        "en": "Row {row}: the indicator is positive but origin {country} is not in the EU, so MSA was not set.",
    },
    "warn_unknown_indicator": {
        "sq": "Rreshti {row}: vlera \"{value}\" në kolonën e preferencës nuk njihet, u trajtua si jo MSA.",
        "en": "Row {row}: value \"{value}\" in the preference column is not recognised, treated as no MSA.",
    },
    # Errors
    "err_xls": {
        "sq": "Skedarët e vjetër .xls nuk lexohen. Hapeni në Excel dhe ruajeni si .xlsx.",
        "en": "Old .xls files can't be read. Open it in Excel and save it as .xlsx.",
    },
    "err_file_missing": {"sq": "Skedari nuk u gjet: {path}", "en": "File not found: {path}"},
    "err_file_unreadable": {
        "sq": "Skedari nuk mund të lexohet ({detail}). A është i dëmtuar apo nuk është skedar Excel?",
        "en": "The file can't be read ({detail}). Is it damaged or not an Excel file?",
    },
    "err_column_name_empty": {
        "sq": "Shkruani emrin e kolonës: {field}.",
        "en": "Enter the column name: {field}.",
    },
    "err_no_header": {
        "sq": "Asnjë nga kolonat {names} nuk u gjet në Excel.",
        "en": "None of the columns {names} were found in the Excel file.",
    },
    "err_columns_missing": {
        "sq": "Këto kolona nuk u gjetën në fletën \"{sheet}\": {missing}.\nKolonat që ka fleta: {found}",
        "en": "These columns were not found in sheet \"{sheet}\": {missing}.\nColumns the sheet has: {found}",
    },
    "err_formula": {
        "sq": "Rreshti {row}, kolona \"{column}\": formula nuk mund të llogaritet. Hapeni skedarin në Excel, ruajeni dhe provoni përsëri.",
        "en": "Row {row}, column \"{column}\": the formula can't be calculated. Open the file in Excel, save it and try again.",
    },
    "err_not_number": {
        "sq": "Rreshti {row}, kolona \"{column}\": \"{value}\" nuk është numër.",
        "en": "Row {row}, column \"{column}\": \"{value}\" is not a number.",
    },
    "err_no_rows": {"sq": "Nuk u gjet asnjë artikull në fletën \"{sheet}\".", "en": "No items were found in sheet \"{sheet}\"."},
    "err_precision": {"sq": "Precision_1 duhet të jetë 00 ose 90.", "en": "Precision_1 must be 00 or 90."},
    "err_extended": {
        "sq": "Extended_customs_procedure duhet të jetë 1000 ose 4000.",
        "en": "Extended_customs_procedure must be 1000 or 4000.",
    },
}


def t(lang: str, key: str, **params) -> str:
    entry = TEXTS.get(key)
    if entry is None:
        return key
    text = entry.get(lang) or entry["en"]
    clean = {k: ", ".join(map(str, v)) if isinstance(v, (list, tuple)) else v for k, v in params.items()}
    if "field" in clean:
        clean["field"] = t(lang, "field_" + clean["field"])
    return text.format(**clean)
