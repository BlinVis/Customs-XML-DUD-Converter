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
    "field_pref": {
        "sq": "Preferenca (MSA / GBMTL)",
        "en": "Preference (MSA / GBMTL)",
    },
    "field_kodi_tarifor": {"sq": "Kodi tarifor", "en": "Tariff code"},
    "q_precision": {
        "sq": "Precision 1 (zakonisht 00 ose 90)",
        "en": "Precision 1 (usually 00 or 90)",
    },
    "q_extended": {
        "sq": "Procedura doganore e zgjeruar (zakonisht 1000 ose 4000)",
        "en": "Extended customs procedure (usually 1000 or 4000)",
    },
    "indicators_hint": {
        "sq": "Preferenca vendoset vetëm kur vlera është PO dhe origjina është nga BE-ja (merr MSA) ose GB (merr GBMTL). Vendet e tjera nuk marrin preferencë.",
        "en": "A preference is set only when the value means YES and the origin is in the EU (gets MSA) or GB (gets GBMTL). Other countries get no preference.",
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
    "need_precision": {
        "sq": "Zgjidhni Precision 1: 00, 90 ose shkruani një kod tjetër te «Tjetër».",
        "en": "Choose Precision 1: 00, 90, or type another code in «Other».",
    },
    "need_extended": {
        "sq": "Zgjidhni procedurën e zgjeruar: 1000, 4000 ose shkruani një kod tjetër te «Tjetër».",
        "en": "Choose the extended procedure: 1000, 4000, or type another code in «Other».",
    },
    "error": {"sq": "Gabim", "en": "Error"},
    "done": {"sq": "U krye", "en": "Done"},
    # Window
    "subtitle": {
        "sq": "Krijoni DUD-in XML nga Excel-i i përkthimit, në 4 hapa.",
        "en": "Create the DUD XML from the translation Excel in 4 steps.",
    },
    "step1_title": {"sq": "Zgjidhni Excel-in e përkthimit", "en": "Choose the translation Excel file"},
    "step1_hint": {
        "sq": "Klikoni «Zgjidh skedarin» dhe gjeni Excel-in e përkthimit për këtë faturë (p.sh. Perkthimi_Bosch_KS.xlsx). Nëse e keni ndryshuar Excel-in, ruajeni para se ta zgjidhni.",
        "en": "Click «Choose file» and find the translation Excel for this invoice (e.g. Perkthimi_Bosch_KS.xlsx). If you changed the Excel, save it before choosing it.",
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
        "sq": "Programi e gjen secilën kolonë sipas titullit të saj në Excel. ✓ do të thotë që kolona u gjet. Nëse del ✗, klikoni shigjetën ▾ në atë rresht dhe zgjidhni titullin e saktë nga lista. Kolonat «(opsionale)» mund të mungojnë. Programi i mban mend zgjedhjet për herën tjetër.",
        "en": "The app finds each column by its title in the Excel. ✓ means the column was found. If you see ✗, click the ▾ arrow on that row and pick the right title from the list. Columns marked «(optional)» may be missing. The app remembers your choices for next time.",
    },
    "found": {"sq": "U gjet", "en": "Found"},
    "not_found": {"sq": "Nuk u gjet", "en": "Not found"},
    "not_used": {"sq": "Nuk ka, s'ka problem", "en": "Not there, that's fine"},
    "step3_title": {"sq": "Përgjigjuni pyetjeve", "en": "Answer the questions"},
    "step3_hint": {
        "sq": "Këto vlera vendosen në çdo artikull të DUD-it. Programi i pyet çdo herë, sepse mund të ndryshojnë nga fatura në faturë.",
        "en": "These values go into every item of the DUD. The app asks every time, because they can change from invoice to invoice.",
    },
    "msa_need_file": {"sq": "Zgjidhni fillimisht Excel-in për të parë opsionet e MSA.",
                      "en": "Choose the Excel file first to see the MSA options."},
    "msa_with_column": {
        "sq": 'Kolona "{column}" tregon cilët artikuj marrin preferencë. Cilat vlera në të do të thonë PO? (të ndara me presje)',
        "en": 'The "{column}" column shows which items get a preference. Which values in it mean YES? (comma separated)',
    },
    "msa_pick_countries": {
        "sq": "Ky Excel nuk ka kolonë për preferencën. Shënoni vendet e origjinës që marrin preferencë në këtë faturë: vendet e BE-së marrin MSA, GB merr GBMTL. Nëse nuk shënoni asnjë, asnjë artikull nuk merr preferencë.",
        "en": "This Excel has no preference column. Tick the countries of origin that get a preference in this invoice: EU countries get MSA, GB gets GBMTL. If you tick none, no item gets a preference.",
    },
    "msa_no_eu": {
        "sq": "Asnjë artikull nuk ka origjinë nga BE-ja ose GB, prandaj nuk ka preferencë.",
        "en": "No item comes from the EU or GB, so there is no preference.",
    },
    "msa_outside_eu": {
        "sq": "Pa preferencë (jashtë BE-së dhe GB): {countries}",
        "en": "No preference (outside the EU and GB): {countries}",
    },
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
    "stat_msa": {
        "sq": "Artikuj me preferencë (MSA / GBMTL)",
        "en": "Items with a preference (MSA / GBMTL)",
    },
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
    "q_code_hint": {
        "sq": "Klikoni njërin nga butonat. Nëse kodi që ju duhet nuk është aty, klikoni «Tjetër» dhe shkruani kodin me {digits} shifra.",
        "en": "Click one of the buttons. If the code you need is not there, click «Other» and type the {digits}-digit code.",
    },
    "other_code": {
        "sq": "Tjetër:",
        "en": "Other:",
    },
    "need_other_code": {
        "sq": "Kodi te «Tjetër» duhet të ketë saktësisht {digits} shifra.",
        "en": "The code in «Other» must have exactly {digits} digits.",
    },
    "msa_title": {
        "sq": "Preferenca (MSA / GBMTL)",
        "en": "Preference (MSA / GBMTL)",
    },
    "save_hint": {
        "sq": "Pasi të klikoni «Krijo XML-in», programi ju pyet ku ta ruani skedarin.",
        "en": "After you click «Create the XML», the app asks where to save the file.",
    },
    "save_title": {
        "sq": "Ku ta ruaj XML-in?",
        "en": "Where should the XML be saved?",
    },
    "xml_filter": {
        "sq": "Skedarë XML",
        "en": "XML files",
    },
    "step4_title": {
        "sq": "Krijoni XML-in",
        "en": "Create the XML",
    },
    "step4_line1": {
        "sq": "Klikoni butonin e gjelbër «Krijo XML-in» poshtë djathtas.",
        "en": "Click the green «Create the XML» button at the bottom right.",
    },
    "step4_line2": {
        "sq": "Zgjidhni dosjen ku do ta ruani dhe klikoni «Save» (Ruaj). Emri i skedarit plotësohet vetë.",
        "en": "Choose the folder to save it in and click Save. The file name is filled in for you.",
    },
    "step4_line3": {
        "sq": "Krahasoni numrin e artikujve, totalin dhe peshat me rreshtin e totaleve në Excel.",
        "en": "Compare the number of items, the total and the weights with the totals row in Excel.",
    },
    "step4_line4": {
        "sq": "Nëse programi shfaq gjëra për t'u kontrolluar (⚠), lexojini para se ta ngarkoni XML-in në ASYCUDA.",
        "en": "If the app lists things to check (⚠), read them before loading the XML into ASYCUDA.",
    },
    # Warnings
    "warn_empty": {"sq": "Rreshti {row}: kolona \"{column}\" është bosh.", "en": "Row {row}: column \"{column}\" is empty."},
    "warn_no_pref_column": {
        "sq": "Nuk ka kolonë preference dhe nuk u shënua asnjë vend, prandaj asnjë artikull nuk ka MSA ose GBMTL.",
        "en": "There is no preference column and no country was ticked, so no item gets MSA or GBMTL.",
    },
    "warn_msa_country_not_eu": {
        "sq": "{country} nuk është në BE dhe nuk është GB, prandaj nuk merr preferencë.",
        "en": "{country} is not in the EU and is not GB, so it gets no preference.",
    },
    "warn_pref_not_eu": {
        "sq": "Rreshti {row}: vlera është PO, por origjina {country} nuk është nga BE-ja ose GB, prandaj preferenca nuk u vendos.",
        "en": "Row {row}: the value means YES, but origin {country} is not in the EU or GB, so no preference was set.",
    },
    "warn_unknown_indicator": {
        "sq": 'Rreshti {row}: vlera "{value}" në kolonën e preferencës nuk njihet, u trajtua si pa preferencë.',
        "en": 'Row {row}: value "{value}" in the preference column is not recognised, treated as no preference.',
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
    "err_precision": {
        "sq": "Precision 1 duhet të ketë 2 shifra (p.sh. 00 ose 90).",
        "en": "Precision 1 must have 2 digits (e.g. 00 or 90).",
    },
    "err_extended": {
        "sq": "Procedura e zgjeruar duhet të ketë 4 shifra (p.sh. 1000 ose 4000).",
        "en": "The extended procedure must have 4 digits (e.g. 1000 or 4000).",
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
