import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from dud_converter import core
from dud_converter.i18n import TEXTS, t

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = FIXTURES / "Perkthimi_ThyssenkrupBilstein_KS.xlsx"
REFERENCE = FIXTURES / "DUDIBILSTEIN17102026.xml"


def reference_xml():
    return REFERENCE.read_bytes().decode("utf-8")


def test_sample_reproduces_reference_byte_for_byte():
    result = core.convert(SAMPLE, core.Options("00", "4000"))
    assert result.xml == reference_xml()
    assert len(result.items) == 22
    assert result.msa_count == 5
    assert result.total_amount == 4436.86
    assert result.warnings == []


def test_written_file_matches_reference(tmp_path):
    result = core.convert(SAMPLE, core.Options("00", "4000"))
    out = tmp_path / "out.xml"
    core.write_xml(result.xml, out)
    assert out.read_bytes() == REFERENCE.read_bytes()


def test_formulas_without_cached_values(tmp_path):
    # openpyxl drops Excel's cached formula results, so Totali must be computed.
    copy = tmp_path / "nocache.xlsx"
    openpyxl.load_workbook(SAMPLE).save(copy)
    assert core.convert(copy, core.Options("00", "4000")).xml == reference_xml()


def test_column_names_are_case_insensitive(tmp_path):
    names = {k: v.upper() + "  " for k, v in core.DEFAULT_COLUMN_NAMES.items()}
    result = core.convert(SAMPLE, core.Options("00", "4000", column_names=names))
    assert result.xml == reference_xml()


def test_answers_go_into_every_item():
    xml = core.convert(SAMPLE, core.Options("90", "1000")).xml
    assert xml.count("<Precision_1>90</Precision_1>") == 22
    assert xml.count("<Extended_customs_procedure>1000</Extended_customs_procedure>") == 22


def test_missing_column_reports_what_the_sheet_has():
    names = dict(core.DEFAULT_COLUMN_NAMES, origjina="Country")
    with pytest.raises(core.ConversionError) as exc:
        core.convert(SAMPLE, core.Options("00", "4000", column_names=names))
    assert exc.value.key == "err_columns_missing"
    assert exc.value.params["missing"] == ["Country"]
    assert "Origjina" in exc.value.params["found"]


def _make_sheet(path, rows, headers=("Kodi", "Perkthim", "Sasia", "Totali", "Pesha neto",
                                     "Pesha bruto", "Origjina", "Pref", "Kodi tarifor")):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return path


def test_description_prefix_msa_rules_and_text_numbers(tmp_path):
    path = _make_sheet(tmp_path / "p.xlsx", [
        ("AB-1", "Filter vaji", 3, "1.234,50", 2, 2.5, "de", "+", "8421 23 00"),  # code missing in text
        ("AB-2", "AB-2 Filter ajri", 1.0, 10, 1, 1.2, "TR", "+", 84213100),         # positive but not EU
        ("AB-3", "Fshirëse", 2, 5, 1, 1, "IT", "?", 85123010),                     # unknown indicator
        (None, None, None, 1249.5, 4, 4.7, None, None, None),                      # totals row is skipped
    ])
    names = dict(core.DEFAULT_COLUMN_NAMES, totali="TOTALI")
    result = core.convert(path, core.Options("00", "4000", column_names=names, positive_indicators=["+"]))
    a, b, c = result.items
    assert a.description == "AB-1 - Filter vaji=3PCE"
    assert a.amount == "1234.5" and a.commodity_code == "84212300"
    assert a.preference_code == "MSA" and a.country == "DE"
    assert b.description == "AB-2 Filter ajri=1PCE" and b.preference_code == ""
    assert c.preference_code == ""
    assert [w.key for w in result.warnings] == ["warn_pref_not_eu", "warn_unknown_indicator"]
    assert "<Preference_code>" in result.xml and result.xml.count("<Preference_code>") == 1


def test_without_pref_column_no_msa(tmp_path):
    path = _make_sheet(tmp_path / "p.xlsx", [("K1", "Tekst", 1, 2, 1, 1, "DE", 1234)],
                       headers=("Kodi", "Perkthim", "Sasia", "Totali", "Pesha neto", "Pesha bruto",
                                "Origjina", "Kodi tarifor"))
    names = dict(core.DEFAULT_COLUMN_NAMES, pref="")
    result = core.convert(path, core.Options("00", "4000", column_names=names))
    assert result.msa_count == 0
    assert [w.key for w in result.warnings] == ["warn_no_pref_column"]


def test_xml_special_characters_are_escaped(tmp_path):
    path = _make_sheet(tmp_path / "p.xlsx", [("K1", 'Kapak <A&B> "x"', 1, 2, 1, 1, "DE", "Po", 1234)])
    xml = core.convert(path, core.Options("00", "4000")).xml
    assert "K1 - Kapak &lt;A&amp;B&gt; \"x\"=1PCE" in xml


def test_output_filename():
    assert core.output_filename(Path("x") / "Perkthimi_Bilstein.xlsx", dt.date(2026, 10, 17)) == \
        "DUDI-Perkthimi_Bilstein-17102026.xml"


def test_every_text_has_both_languages():
    for key, entry in TEXTS.items():
        assert set(entry) == {"sq", "en"}, key
    assert "Kujdes" in t("sq", "warnings")
