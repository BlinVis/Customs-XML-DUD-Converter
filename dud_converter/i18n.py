"""Albanian (sq) and English (en) texts for the app."""

LANGUAGES = {"sq": "Shqip", "en": "English"}

TEXTS = {
    "title": {"sq": "Perkthim → DUD XML (ASYCUDA)", "en": "Perkthim → DUD XML (ASYCUDA)"},
    "optional": {"sq": "(opsionale)", "en": "(optional)"},
    "field_kodi": {"sq": "Kodi i artikullit", "en": "Article code"},
    "field_perkthim": {"sq": "Përkthimi (përshkrimi)", "en": "Translation (description)"},
    "field_sasia": {"sq": "Sasia", "en": "Quantity"},
    "field_totali": {"sq": "Totali (jo ai me zbritje)", "en": "Total (not the discounted one)"},
    "field_pesha_bruto": {"sq": "Pesha bruto", "en": "Gross weight"},
    "field_pesha_neto": {"sq": "Pesha neto", "en": "Net weight"},
    "field_origjina": {"sq": "Vendi i origjinës", "en": "Country of origin"},
    "field_pref": {"sq": "Preferenca (MSA)", "en": "Preference (MSA)"},
    "field_kodi_tarifor": {"sq": "Kodi tarifor", "en": "Tariff code"},
    "q_precision": {"sq": "Precision 1 — 00 apo 90?", "en": "Precision 1 — 00 or 90?"},
    "q_extended": {
        "sq": "Procedura doganore e zgjeruar — 1000 apo 4000?",
        "en": "Extended customs procedure — 1000 or 4000?",
    },
    "indicators_hint": {
        "sq": "MSA vendoset vetëm kur treguesi është pozitiv DHE vendi i origjinës është në BE.",
        "en": "MSA is set only when the indicator is positive AND the country of origin is in the EU.",
    },
    "constants": {"sq": "Vlerat fikse", "en": "Fixed values"},
    "constants_title": {"sq": "Vlerat fikse të çdo artikulli", "en": "Fixed values for every item"},
    "save": {"sq": "Ruaj", "en": "Save"},
    "reset": {"sq": "Kthe parazgjedhjet", "en": "Reset to defaults"},
    "cancel": {"sq": "Anulo", "en": "Cancel"},
    "create": {"sq": "Krijo XML-in  →", "en": "Create the XML  →"},
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
    # Window
    "subtitle": {
        "sq": "Krijoni DUD-in XML nga përkthimi në Excel, në 3 hapa.",
        "en": "Create the DUD XML from the Excel translation in 3 steps.",
    },
    "step1_title": {"sq": "Zgjidhni Excel-in e përkthimit", "en": "Choose the translation Excel file"},
    "step1_hint": {
        "sq": "Skedari që dërgon furnitori (perkthimi), p.sh. Perkthimi_Bosch_KS.xlsx.",
        "en": "The file the supplier sends (perkthim), e.g. Perkthimi_Bosch_KS.xlsx.",
    },
    "choose_file": {"sq": "Zgjidh skedarin", "en": "Choose file"},
    "no_file": {"sq": "Asnjë skedar i zgjedhur", "en": "No file chosen"},
    "step1_none": {"sq": "Klikoni butonin dhe zgjidhni skedarin Excel.", "en": "Click the button and pick the Excel file."},
    "file_ready": {"sq": "Skedari u lexua (fleta \"{sheet}\"). Të gjitha kolonat u gjetën.",
                   "en": "File read (sheet \"{sheet}\"). All columns were found."},
    "columns_missing_short": {
        "sq": "Disa kolona nuk u gjetën. Rregullojini te hapi 2.",
        "en": "Some columns were not found. Fix them in step 2.",
    },
    "step2_title": {"sq": "Kontrolloni kolonat", "en": "Check the columns"},
    "step2_hint": {
        "sq": "Programi i gjen kolonat sipas titullit në rreshtin e parë të Excel-it. Nëse diçka del me ✗, zgjidhni titullin e saktë nga lista. Programi i mban mend për herën tjetër.",
        "en": "The app finds each column by its title in the Excel's first row. If one shows ✗, pick the right title from the list. The app remembers it for next time.",
    },
    "found": {"sq": "U gjet", "en": "Found"},
    "not_found": {"sq": "Nuk u gjet", "en": "Not found"},
    "not_used": {"sq": "Nuk ka, s'ka problem", "en": "Not there, that's fine"},
    "step3_title": {"sq": "Përgjigjuni pyetjeve", "en": "Answer the questions"},
    "step3_hint": {
        "sq": "Këto vlera vendosen në çdo artikull të DUD-it.",
        "en": "These values go into every item of the DUD.",
    },
    "msa_need_file": {"sq": "Zgjidhni fillimisht Excel-in për të parë opsionet e MSA.",
                      "en": "Choose the Excel file first to see the MSA options."},
    "msa_with_column": {
        "sq": "Kolona \"{column}\" tregon ku vendoset MSA. Cilat vlera në të do të thonë PO? (të ndara me presje)",
        "en": "The \"{column}\" column shows where MSA goes. Which values in it mean YES? (comma separated)",
    },
    "msa_pick_countries": {
        "sq": "Ky Excel nuk ka kolonë për MSA. Shënoni vendet e origjinës që marrin MSA:",
        "en": "This Excel has no MSA column. Tick the countries of origin that get MSA:",
    },
    "msa_no_eu": {
        "sq": "Asnjë artikull nuk ka origjinë nga BE, prandaj nuk ka MSA.",
        "en": "No item has an EU origin, so there is no MSA.",
    },
    "msa_outside_eu": {"sq": "Jashtë BE-së (pa MSA): {countries}", "en": "Outside the EU (no MSA): {countries}"},
    "save_to": {"sq": "XML do të ruhet në: {folder}", "en": "The XML will be saved in: {folder}"},
    "save_to_default": {"sq": "XML do të ruhet në të njëjtën dosje me Excel-in.",
                        "en": "The XML will be saved in the same folder as the Excel file."},
    "change_folder": {"sq": "Ndrysho dosjen", "en": "Change folder"},
    "constants_hint": {
        "sq": "Këto vlera janë të njëjta për çdo artikull. Ndryshojini vetëm nëse e dini që duhet.",
        "en": "These are the same for every item. Only change them if you know they must change.",
    },
    "missing_title": {"sq": "Mungon diçka", "en": "Something is missing"},
    "err_save": {
        "sq": "XML nuk u ruajt ({detail}). A është skedari i hapur në një program tjetër?",
        "en": "The XML could not be saved ({detail}). Is the file open in another program?",
    },
    "result_title": {"sq": "XML u krijua me sukses", "en": "The XML was created"},
    "stat_items": {"sq": "Artikuj", "en": "Items"},
    "stat_msa": {"sq": "Artikuj me MSA", "en": "Items with MSA"},
    "stat_total": {"sq": "Totali", "en": "Total"},
    "stat_gross": {"sq": "Pesha bruto", "en": "Gross weight"},
    "stat_net": {"sq": "Pesha neto", "en": "Net weight"},
    "compare_totals": {
        "sq": "Krahasoni këto shuma me rreshtin e totaleve në Excel para se ta ngarkoni në ASYCUDA.",
        "en": "Compare these sums with the totals row in Excel before loading it into ASYCUDA.",
    },
    "warnings_title": {"sq": "Kontrolloni {count} gjëra", "en": "Check {count} things"},
    "open_folder": {"sq": "Hap dosjen", "en": "Open folder"},
    "close": {"sq": "Mbyll", "en": "Close"},
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
