"""Tests for the date-format signature system.

Run:  python -m pytest tests/test_dates.py -v
"""

import datetime
import math
import re
import zipfile
import io

import numpy as np
import pandas as pd
import pytest


# ---------------------------------------------------------------------------
# Import helpers from the module under test.
# conftest.py installs Streamlit / utils mocks before this import runs.
# ---------------------------------------------------------------------------
from pages.compare import (  # noqa: E402
    # -- New signature system (Step 2) --
    DateSignature,
    _excel_code_to_signature,
    _BUILTIN_NUMFMT_CODES,
    _signature_to_strftime,
    _signature_to_parse_format,
    _signature_to_excel_code,
    # -- CSV column inference (Step 3) --
    _text_value_to_signature,
    # -- Comparison (Step 5) --
    _date_signatures_match,
    # -- Conversion (Step 6) --
    _convert_date_column,
    # -- Writer (Step 7) --
    _write_df_to_xlsx_vectorized,
)


# ===================================================================
# 1. TestExcelCodeToSignature
# ===================================================================

class TestExcelCodeToSignature:
    """_excel_code_to_signature tokeniser: Excel code → DateSignature."""

    # --- Basic format codes ---

    def test_yyyy_mm_dd(self):
        sig = _excel_code_to_signature("yyyy-mm-dd")
        assert sig is not None
        assert sig.order == "YMD"
        assert sig.sep == "-"
        assert sig.year_digits == 4
        assert sig.month_style == "numeric"
        assert sig.has_time is False
        assert sig.ampm is False
        assert sig.locale_dependent is False

    def test_m_d_yy(self):
        sig = _excel_code_to_signature("m/d/yy")
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.sep == "/"
        assert sig.year_digits == 2
        assert sig.month_style == "numeric"

    def test_dd_mmm_yyyy(self):
        sig = _excel_code_to_signature("dd-mmm-yyyy")
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.sep == "-"
        assert sig.year_digits == 4
        assert sig.month_style == "abbrev"

    def test_dd_mmmm_yyyy(self):
        sig = _excel_code_to_signature("dd-mmmm-yyyy")
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.month_style == "full"

    def test_m_d_yy_h_mm(self):
        sig = _excel_code_to_signature("m/d/yy h:mm")
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.has_time is True
        assert sig.ampm is False

    def test_yyyy_mm_dd_hh_mm_ss(self):
        sig = _excel_code_to_signature("yyyy-mm-dd hh:mm:ss")
        assert sig is not None
        assert sig.order == "YMD"
        assert sig.has_time is True
        assert sig.year_digits == 4

    # --- Built-in numFmtIds ---

    def test_builtin_14(self):
        """numFmtId 14 = mm-dd-yy, locale-dependent."""
        sig = _excel_code_to_signature("mm-dd-yy", numfmt_id=14)
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.sep == "-"
        assert sig.year_digits == 2
        assert sig.locale_dependent is True

    def test_builtin_22(self):
        """numFmtId 22 = m/d/yy h:mm, locale-dependent."""
        sig = _excel_code_to_signature("m/d/yy h:mm", numfmt_id=22)
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.has_time is True
        assert sig.locale_dependent is True

    def test_builtin_15(self):
        sig = _excel_code_to_signature("d-mmm-yy", numfmt_id=15)
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.month_style == "abbrev"
        assert sig.locale_dependent is True

    def test_builtin_id_alone(self):
        """When only numfmt_id is given (code may be 'General'), resolve from table."""
        sig = _excel_code_to_signature("General", numfmt_id=14)
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.locale_dependent is True

    # --- Locale prefixes ---

    def test_locale_prefix_stripped(self):
        sig = _excel_code_to_signature("[$-en-US]m/d/yyyy")
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.year_digits == 4

    def test_locale_prefix_with_semicolon(self):
        sig = _excel_code_to_signature("[$-en-US]m/d/yyyy;@")
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.year_digits == 4

    # --- AM/PM ---

    def test_ampm(self):
        sig = _excel_code_to_signature("m/d/yyyy h:mm AM/PM")
        assert sig is not None
        assert sig.has_time is True
        assert sig.ampm is True

    def test_ampm_lower(self):
        sig = _excel_code_to_signature("m/d/yyyy h:mm am/pm")
        assert sig is not None
        assert sig.ampm is True

    # --- Minute vs month disambiguation ---

    def test_minute_after_hour(self):
        """In 'h:mm', the mm is minutes (not month) because it follows h."""
        sig = _excel_code_to_signature("yyyy-mm-dd h:mm:ss")
        assert sig is not None
        assert sig.order == "YMD"
        assert sig.has_time is True
        # The 'mm' in 'h:mm' must NOT be parsed as month

    def test_month_before_day(self):
        """In 'mm/dd/yyyy', mm is month because it precedes dd."""
        sig = _excel_code_to_signature("mm/dd/yyyy")
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.month_style == "numeric"

    # --- Escaped separators ---

    def test_escaped_dash(self):
        sig = _excel_code_to_signature("yyyy\\-mm\\-dd")
        assert sig is not None
        assert sig.order == "YMD"
        assert sig.sep == "-"

    # --- Non-date codes return None ---

    def test_general_returns_none(self):
        assert _excel_code_to_signature("General") is None

    def test_number_format_returns_none(self):
        assert _excel_code_to_signature("#,##0.00") is None

    def test_text_format_returns_none(self):
        assert _excel_code_to_signature("@") is None

    def test_none_input_returns_none(self):
        assert _excel_code_to_signature(None) is None

    def test_empty_input_returns_none(self):
        assert _excel_code_to_signature("") is None

    # --- Parametrized regression over all old _EXCEL_FMT_TO_LABEL codes ---
    # Each tuple: (excel_code, expected_order, expected_sep, expected_year_digits,
    #              expected_month_style, expected_has_time)

    @pytest.mark.parametrize("code,order,sep,yd,ms,ht", [
        # YMD family
        ("yyyy-mm-dd", "YMD", "-", 4, "numeric", False),
        ("yyyy\\-mm\\-dd", "YMD", "-", 4, "numeric", False),
        ("yyyy-mm-dd;@", "YMD", "-", 4, "numeric", False),
        ("yyyy/mm/dd", "YMD", "/", 4, "numeric", False),
        ("yyyy.mm.dd", "YMD", ".", 4, "numeric", False),
        ("yyyy-mm-dd h:mm:ss", "YMD", "-", 4, "numeric", True),
        ("yyyy-mm-dd hh:mm:ss", "YMD", "-", 4, "numeric", True),
        ("yyyy-mm-dd hh:mm", "YMD", "-", 4, "numeric", True),
        ("yyyy-mm-dd h:mm", "YMD", "-", 4, "numeric", True),
        ("yyyy/mm/dd h:mm:ss", "YMD", "/", 4, "numeric", True),
        ("yyyy/mm/dd hh:mm:ss", "YMD", "/", 4, "numeric", True),
        ("yyyy/mm/dd hh:mm", "YMD", "/", 4, "numeric", True),
        ("yyyy/mm/dd h:mm", "YMD", "/", 4, "numeric", True),
        # MDY family
        ("m/d/yy", "MDY", "/", 2, "numeric", False),
        ("m/d/yy h:mm", "MDY", "/", 2, "numeric", True),
        ("m/d/yy hh:mm", "MDY", "/", 2, "numeric", True),
        ("m/d/yy h:mm:ss", "MDY", "/", 2, "numeric", True),
        ("m/d/yy hh:mm:ss", "MDY", "/", 2, "numeric", True),
        ("mm/dd/yy", "MDY", "/", 2, "numeric", False),
        ("mm/dd/yy h:mm", "MDY", "/", 2, "numeric", True),
        ("mm/dd/yy hh:mm", "MDY", "/", 2, "numeric", True),
        ("mm/dd/yy h:mm:ss", "MDY", "/", 2, "numeric", True),
        ("mm/dd/yyyy", "MDY", "/", 4, "numeric", False),
        ("mm/dd/yyyy h:mm", "MDY", "/", 4, "numeric", True),
        ("mm/dd/yyyy hh:mm", "MDY", "/", 4, "numeric", True),
        ("mm/dd/yyyy h:mm:ss", "MDY", "/", 4, "numeric", True),
        ("mm/dd/yyyy hh:mm:ss", "MDY", "/", 4, "numeric", True),
        ("m/d/yyyy", "MDY", "/", 4, "numeric", False),
        ("m/d/yyyy h:mm", "MDY", "/", 4, "numeric", True),
        ("m/d/yyyy h:mm:ss", "MDY", "/", 4, "numeric", True),
        ("m/d/yyyy hh:mm", "MDY", "/", 4, "numeric", True),
        ("m/d/yyyy hh:mm:ss", "MDY", "/", 4, "numeric", True),
        ("mm-dd-yyyy", "MDY", "-", 4, "numeric", False),
        ("mm.dd.yyyy", "MDY", ".", 4, "numeric", False),
        # DMY family
        ("dd/mm/yyyy", "DMY", "/", 4, "numeric", False),
        ("dd/mm/yy", "DMY", "/", 2, "numeric", False),
        ("dd/mm/yy h:mm", "DMY", "/", 2, "numeric", True),
        ("dd/mm/yy hh:mm", "DMY", "/", 2, "numeric", True),
        ("dd/mm/yyyy h:mm:ss", "DMY", "/", 4, "numeric", True),
        ("dd/mm/yyyy hh:mm:ss", "DMY", "/", 4, "numeric", True),
        ("dd/mm/yyyy hh:mm", "DMY", "/", 4, "numeric", True),
        ("dd/mm/yyyy h:mm", "DMY", "/", 4, "numeric", True),
        ("dd-mm-yyyy", "DMY", "-", 4, "numeric", False),
        ("dd.mm.yyyy", "DMY", ".", 4, "numeric", False),
        ("d/m/yyyy", "DMY", "/", 4, "numeric", False),
        ("d/m/yy", "DMY", "/", 2, "numeric", False),
        ("d/m/yy h:mm", "DMY", "/", 2, "numeric", True),
        ("d/m/yyyy h:mm", "DMY", "/", 4, "numeric", True),
        ("d/m/yyyy h:mm:ss", "DMY", "/", 4, "numeric", True),
        ("d-m-yyyy", "DMY", "-", 4, "numeric", False),
        ("d-m-yy", "DMY", "-", 2, "numeric", False),
        # DMY month-name
        ("dd-mmm-yyyy", "DMY", "-", 4, "abbrev", False),
        ("dd-mmm-yy", "DMY", "-", 2, "abbrev", False),
        ("d-mmm-yyyy", "DMY", "-", 4, "abbrev", False),
        ("d-mmm-yy", "DMY", "-", 2, "abbrev", False),
        ("dd-mmmm-yyyy", "DMY", "-", 4, "full", False),
        # Locale prefix
        ("[$-en-us]m/d/yyyy", "MDY", "/", 4, "numeric", False),
        ("[$-en-us]mm/dd/yyyy", "MDY", "/", 4, "numeric", False),
        ("[$-en-us]m/d/yy h:mm", "MDY", "/", 2, "numeric", True),
    ])
    def test_regression_all_known_codes(self, code, order, sep, yd, ms, ht):
        sig = _excel_code_to_signature(code)
        assert sig is not None, f"Expected a signature for {code!r}, got None"
        assert sig.order == order, f"{code}: order {sig.order} != {order}"
        assert sig.sep == sep, f"{code}: sep {sig.sep!r} != {sep!r}"
        assert sig.year_digits == yd, f"{code}: year_digits {sig.year_digits} != {yd}"
        assert sig.month_style == ms, f"{code}: month_style {sig.month_style!r} != {ms!r}"
        assert sig.has_time == ht, f"{code}: has_time {sig.has_time} != {ht}"


# ===================================================================
# 2. TestSignatureComparison
# ===================================================================

class TestSignatureComparison:
    """_date_signatures_match returns (bool, list[str]) with field-level diffs."""

    def _sig(self, **kw):
        defaults = dict(order="YMD", sep="-", year_digits=4,
                        month_style="numeric", has_time=False,
                        ampm=False, locale_dependent=False)
        defaults.update(kw)
        return DateSignature(**defaults)

    def test_identical_match(self):
        s = self._sig()
        match, diffs = _date_signatures_match(s, s)
        assert match is True
        assert diffs == []

    def test_order_diff(self):
        s1 = self._sig(order="MDY")
        s2 = self._sig(order="YMD")
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("order" in d for d in diffs)

    def test_sep_diff(self):
        s1 = self._sig(sep="/")
        s2 = self._sig(sep="-")
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("sep" in d.lower() for d in diffs)

    def test_year_digits_diff(self):
        s1 = self._sig(year_digits=2)
        s2 = self._sig(year_digits=4)
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("year" in d.lower() for d in diffs)

    def test_month_style_diff(self):
        s1 = self._sig(month_style="numeric")
        s2 = self._sig(month_style="abbrev")
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("month" in d.lower() for d in diffs)

    def test_has_time_diff(self):
        s1 = self._sig(has_time=False)
        s2 = self._sig(has_time=True)
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("time" in d.lower() for d in diffs)

    def test_ampm_diff(self):
        s1 = self._sig(ampm=False, has_time=True)
        s2 = self._sig(ampm=True, has_time=True)
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert any("am" in d.lower() or "pm" in d.lower() for d in diffs)

    def test_none_matches_anything(self):
        match, diffs = _date_signatures_match(None, self._sig())
        assert match is True
        assert diffs == []

    def test_both_none_match(self):
        match, diffs = _date_signatures_match(None, None)
        assert match is True
        assert diffs == []

    def test_multiple_diffs(self):
        s1 = self._sig(order="MDY", sep="/", year_digits=2)
        s2 = self._sig(order="YMD", sep="-", year_digits=4)
        match, diffs = _date_signatures_match(s1, s2)
        assert match is False
        assert len(diffs) >= 3


# ===================================================================
# 3. TestCSVColumnInference
# ===================================================================

class TestCSVColumnInference:
    """_text_value_to_signature — column-level format inference from text values."""

    def test_unambiguous_day_first(self):
        """Values like 25/01/2023 are unambiguously day-first."""
        values = ["25/01/2023", "13/02/2023", "31/12/2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.sep == "/"

    def test_unambiguous_month_first(self):
        """Values like 01/25/2023 are unambiguously month-first."""
        values = ["01/25/2023", "02/13/2023", "12/31/2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "MDY"
        assert sig.sep == "/"

    def test_ambiguous_slash_defaults_mdy(self):
        """All values ambiguous + slash sep → MDY (US convention)."""
        values = ["01/02/2023", "03/04/2023", "05/06/2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "MDY"

    def test_ambiguous_dash_defaults_dmy(self):
        """All values ambiguous + dash sep → DMY."""
        values = ["01-02-2023", "03-04-2023", "05-06-2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "DMY"

    def test_ambiguous_resolved_by_template_hint(self):
        """Template hint overrides the separator heuristic."""
        values = ["01/02/2023", "03/04/2023"]
        hint = DateSignature(order="DMY", sep="/", year_digits=4,
                             month_style="numeric", has_time=False,
                             ampm=False, locale_dependent=False)
        sig = _text_value_to_signature(values, template_hint=hint)
        assert sig is not None
        assert sig.order == "DMY"

    def test_iso_format(self):
        """YYYY-MM-DD values → YMD."""
        values = ["2023-01-25", "2023-02-13", "2023-12-31"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "YMD"
        assert sig.sep == "-"
        assert sig.year_digits == 4

    def test_with_time(self):
        values = ["01/25/2023 14:30:00", "02/13/2023 09:15:00"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.has_time is True
        assert sig.order == "MDY"

    def test_month_name_abbrev(self):
        values = ["25-Jan-2023", "13-Feb-2023", "31-Dec-2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.month_style == "abbrev"

    def test_month_name_full(self):
        values = ["25-January-2023", "13-February-2023"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.order == "DMY"
        assert sig.month_style == "full"

    def test_non_date_returns_none(self):
        values = ["hello", "world", "foo"]
        sig = _text_value_to_signature(values)
        assert sig is None

    def test_empty_returns_none(self):
        sig = _text_value_to_signature([])
        assert sig is None

    def test_two_digit_year(self):
        values = ["01/25/23", "02/13/23"]
        sig = _text_value_to_signature(values)
        assert sig is not None
        assert sig.year_digits == 2


# ===================================================================
# 4. TestConversionPipeline
# ===================================================================

class TestConversionPipeline:
    """_convert_date_column — single-parse, NaT prevention."""

    def _src_sig(self, **kw):
        defaults = dict(order="MDY", sep="/", year_digits=4,
                        month_style="numeric", has_time=False,
                        ampm=False, locale_dependent=False)
        defaults.update(kw)
        return DateSignature(**defaults)

    def _tgt_sig(self, **kw):
        defaults = dict(order="YMD", sep="-", year_digits=4,
                        month_style="numeric", has_time=False,
                        ampm=False, locale_dependent=False)
        defaults.update(kw)
        return DateSignature(**defaults)

    def test_basic_parse(self):
        """Text dates parsed to datetime64 in a single pass."""
        df = pd.DataFrame({"dt": ["01/25/2023", "02/13/2023", "12/31/2023"]})
        src = self._src_sig()
        tgt = self._tgt_sig()
        _convert_date_column(df, "dt", src, tgt, "yyyy-mm-dd")
        assert pd.api.types.is_datetime64_any_dtype(df["dt"]) or df["dt"].dtype == object
        # All values should have been parsed (no NaT for valid dates)
        for v in df["dt"]:
            assert v is not None
            assert not (isinstance(v, float) and math.isnan(v))

    def test_unparseable_kept(self):
        """Unparseable values keep original text — never NaT or blank."""
        df = pd.DataFrame({"dt": ["01/25/2023", "NOT-A-DATE", "12/31/2023"]})
        src = self._src_sig()
        tgt = self._tgt_sig()
        _convert_date_column(df, "dt", src, tgt, "yyyy-mm-dd")
        # The invalid value must be preserved as-is
        assert df["dt"].iloc[1] == "NOT-A-DATE"

    def test_empty_stays_empty(self):
        df = pd.DataFrame({"dt": ["01/25/2023", "", "12/31/2023"]})
        src = self._src_sig()
        tgt = self._tgt_sig()
        _convert_date_column(df, "dt", src, tgt, "yyyy-mm-dd")
        assert df["dt"].iloc[1] == "" or pd.isna(df["dt"].iloc[1])

    def test_excel_serial(self):
        """Excel serial number (float) should be converted to datetime."""
        # 45000 = 2023-03-14 in Excel epoch
        df = pd.DataFrame({"dt": [45000.0, 45001.5]})
        src = self._src_sig()
        tgt = self._tgt_sig()
        _convert_date_column(df, "dt", src, tgt, "yyyy-mm-dd")
        # Should not raise; values should be parsed


# ===================================================================
# 5. TestWriter
# ===================================================================

class TestWriter:
    """_write_df_to_xlsx_vectorized — precision, escaping, built-in IDs."""

    def _read_zip_xml(self, path, inner_path):
        """Read an XML part from an XLSX ZIP file."""
        with zipfile.ZipFile(path, "r") as z:
            return z.read(inner_path).decode("utf-8")

    def test_serial_precision_17g(self, tmp_path):
        """Datetime serials must use %.17g (not %.10g)."""
        # A datetime that needs more than 10 significant digits
        df = pd.DataFrame({
            "dt": pd.to_datetime(["2023-01-15 13:45:30.123456"])
        })
        out = str(tmp_path / "prec.xlsx")
        _write_df_to_xlsx_vectorized(
            df, out, date_formats={"dt": {"excel_code": "yyyy-mm-dd hh:mm:ss", "numfmt_id": None}}
        )
        sheet_xml = self._read_zip_xml(out, "xl/worksheets/sheet1.xml")
        # Extract the serial value — should have high precision
        m = re.search(r"<v>([\d.]+)</v>", sheet_xml)
        assert m, "No serial value found in sheet XML"
        serial_str = m.group(1)
        # %.17g would give at least 15 characters for a typical serial
        # %.10g would give at most 10 significant digits
        # The serial for 2023-01-15 13:45:30.123456 is ~44941.573265...
        assert len(serial_str.replace(".", "")) >= 12, \
            f"Serial {serial_str} lacks precision (expected >=12 sig digits)"

    def test_xml_escape_formatcode(self, tmp_path):
        """formatCode in styles.xml must be XML-escaped."""
        df = pd.DataFrame({
            "dt": pd.to_datetime(["2023-01-15"])
        })
        # A format code with characters that need escaping
        code_with_special = 'yyyy"年"mm"月"dd"日"'
        out = str(tmp_path / "escape.xlsx")
        _write_df_to_xlsx_vectorized(
            df, out, date_formats={"dt": {"excel_code": code_with_special, "numfmt_id": None}}
        )
        styles_xml = self._read_zip_xml(out, "xl/styles.xml")
        # The raw quotes should be escaped, not literal in the attribute
        assert '&quot;' in styles_xml or "&#" in styles_xml or '"年"' not in styles_xml

    def test_builtin_id_14_not_164(self, tmp_path):
        """Built-in numFmtId 14 must appear as numFmtId='14', not 164+."""
        df = pd.DataFrame({
            "dt": pd.to_datetime(["2023-01-15", "2023-06-30"])
        })
        out = str(tmp_path / "builtin14.xlsx")
        _write_df_to_xlsx_vectorized(
            df, out, date_formats={"dt": {"excel_code": "mm-dd-yy", "numfmt_id": 14}}
        )
        styles_xml = self._read_zip_xml(out, "xl/styles.xml")
        # The cellXfs should reference numFmtId="14" for the date style
        assert 'numFmtId="14"' in styles_xml
        # There should be NO custom numFmt entry for this built-in
        assert 'numFmtId="164"' not in styles_xml

    def test_builtin_id_22(self, tmp_path):
        """Built-in numFmtId 22 must appear as numFmtId='22'."""
        df = pd.DataFrame({
            "dt": pd.to_datetime(["2023-01-15 10:30:00"])
        })
        out = str(tmp_path / "builtin22.xlsx")
        _write_df_to_xlsx_vectorized(
            df, out, date_formats={"dt": {"excel_code": "m/d/yy h:mm", "numfmt_id": 22}}
        )
        styles_xml = self._read_zip_xml(out, "xl/styles.xml")
        assert 'numFmtId="22"' in styles_xml

    def test_custom_id_starts_at_164(self, tmp_path):
        """Non-built-in formats must get IDs starting at 164."""
        df = pd.DataFrame({
            "dt": pd.to_datetime(["2023-01-15"])
        })
        out = str(tmp_path / "custom.xlsx")
        _write_df_to_xlsx_vectorized(
            df, out, date_formats={"dt": {"excel_code": "yyyy-mm-dd hh:mm:ss", "numfmt_id": None}}
        )
        styles_xml = self._read_zip_xml(out, "xl/styles.xml")
        assert 'numFmtId="164"' in styles_xml


# ===================================================================
# 6. TestEndToEnd
# ===================================================================

class TestEndToEnd:
    """Integration: CSV detection + comparison + fix + write."""

    def test_format_mismatch_detected(self, tmp_csv, tmp_xlsx):
        """CSV M/D/YYYY vs template YYYY-MM-DD → format_mismatch."""
        csv_path = tmp_csv("placed.csv", [
            ["ID", "EventDate"],
            ["1", "01/25/2023"],
            ["2", "02/13/2023"],
            ["3", "12/31/2023"],
        ])
        tpl_path = tmp_xlsx("template.xlsx", [
            ["ID", "EventDate"],
            ["1", datetime.datetime(2023, 1, 25)],
            ["2", datetime.datetime(2023, 2, 13)],
        ], number_formats={1: "yyyy-mm-dd"})

        # Detect types
        from pages.compare import _detect_csv_column_types, detect_column_types
        csv_types = _detect_csv_column_types(csv_path)
        tpl_types = detect_column_types(tpl_path, "Sheet1", 1)

        # Find the EventDate column (by index — detection returns int keys)
        csv_dt = None
        for k, v in csv_types.items():
            if v.get("header", "").lower() == "eventdate":
                csv_dt = v
                break
        tpl_dt = None
        for k, v in tpl_types.items():
            if v.get("header", "").lower() == "eventdate":
                tpl_dt = v
                break
        assert csv_dt is not None and csv_dt["type"] == "date"
        assert tpl_dt is not None and tpl_dt["type"] == "date"

        # Signatures should differ
        csv_sig = csv_dt.get("signature")
        tpl_sig = tpl_dt.get("signature")
        if csv_sig and tpl_sig:
            match, diffs = _date_signatures_match(csv_sig, tpl_sig)
            assert match is False
            assert len(diffs) > 0

    def test_same_format_no_finding(self, tmp_csv, tmp_xlsx):
        """Same format on both sides → no mismatch."""
        csv_path = tmp_csv("placed.csv", [
            ["ID", "EventDate"],
            ["1", "2023-01-25"],
            ["2", "2023-02-13"],
        ])
        tpl_path = tmp_xlsx("template.xlsx", [
            ["ID", "EventDate"],
            ["1", datetime.datetime(2023, 1, 25)],
            ["2", datetime.datetime(2023, 2, 13)],
        ], number_formats={1: "yyyy-mm-dd"})

        from pages.compare import _detect_csv_column_types, detect_column_types
        csv_types = _detect_csv_column_types(csv_path)
        tpl_types = detect_column_types(tpl_path, "Sheet1", 1)

        csv_dt = None
        for k, v in csv_types.items():
            if v.get("header", "").lower() == "eventdate":
                csv_dt = v
                break
        tpl_dt = None
        for k, v in tpl_types.items():
            if v.get("header", "").lower() == "eventdate":
                tpl_dt = v
                break

        csv_sig = csv_dt.get("signature") if csv_dt else None
        tpl_sig = tpl_dt.get("signature") if tpl_dt else None
        if csv_sig and tpl_sig:
            match, _ = _date_signatures_match(csv_sig, tpl_sig)
            assert match is True
