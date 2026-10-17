# Perkthim → DUD XML (ASYCUDA)

Desktop app that turns a supplier's **perkthim** Excel sheet into the item XML
that the old ASYCUDA client imports (the same file Excel produces from the
"XML FORMA" template). Albanian and English interface.

## Si përdoret / How to use

1. Hapni `DUD-Converter.exe` dhe klikoni **Zgjidh skedarin** / open it and click **Choose file**.
2. Te hapi 2 çdo kolonë duhet të ketë ✓. Nëse ka ✗, zgjidhni titullin e saktë nga lista
   (mbahet mend për herën tjetër). / In step 2 every column should show ✓; if one shows ✗,
   pick the right title from the list (remembered for next time).
3. Zgjidhni **00 / 90** dhe **1000 / 4000**. Për MSA: nëse Excel-i ka kolonë preference, shkruani
   cilat vlera do të thonë PO (p.sh. `Po`, `+`); nëse s'ka, shënoni vendet e BE-së që marrin MSA.
   / Pick **00 / 90** and **1000 / 4000**. For MSA: with a preference column, enter the values that
   mean YES; without one, tick the EU origin countries that get MSA.
4. **Krijo XML-in / Create the XML** → `DUDI-<emri i Excel-it>-<DDMMYYYY>.xml`, pranë Excel-it ose
   në dosjen e zgjedhur. / saved next to the Excel file or in the chosen folder.

The result window shows the item count, MSA count and the totals (amount, gross,
net) so they can be checked against the totals row of the Excel sheet, plus any
warnings (empty cells, a positive MSA indicator on a non-EU origin, unknown
indicator values).

## What the app does with each row

| XML element | Source |
|---|---|
| Commodity_code | tariff code column (digits only) |
| Precision_1 / Extended_customs_procedure | your answers |
| Preference_code | `MSA` only when the indicator is positive **and** the origin is an EU country; otherwise left out |
| Preference_code, sheet without an indicator column | `MSA` for the origin countries you type in (EU countries only); leave the field empty for no MSA |
| Country_of_origin_code | origin column (upper case) |
| Commercial_Description | `perkthim=<sasia>PCE`; the article code is put in front (`<kodi> - perkthim`) when the text doesn't already start with it |
| Gross_weight_itm / Net_weight_itm | weight columns |
| Amount_foreign_currency | **Totali** only (never the discounted total), rounded to 2 decimals |
| Number_of_packages, Kind_of_packages_code/name, National_customs_procedure, Valuation_method_code, Rate_of_adjustement, Currency_code/name/rate | fixed values (button **Vlerat fikse / Fixed values**; defaults 0, PA, Paketa, 000, 1, 1, EUR, "Nuk ka valutë të huaj", 1) |
| Marks2_of_packages, Summary_declaration, Statistical_value | always left empty |

Rows with neither a perkthim text nor an article code (for example the totals
row) are skipped. `Totali` formulas such as `=E2*D2` are read from Excel's saved
results, or calculated if the file has none. Old `.xls` files must be saved as
`.xlsx` first.

## Running from source / building the .exe

```bat
run_app.bat        :: needs Python 3.10+
build_exe.bat      :: produces dist\DUD-Converter.exe
```

Every push also builds the Windows `.exe` on GitHub Actions (artifact
**DUD-Converter-windows** on the workflow run).

Tests: `pip install -r requirements.txt pytest && pytest`. They regenerate
`tests/fixtures/DUDIBILSTEIN17102026.xml` from the sample perkthim and require a
byte-for-byte match.
