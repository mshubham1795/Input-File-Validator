"""Shared fixtures for date-format tests.

Mocks Streamlit and utils so ``pages.compare`` can be imported in a
plain pytest session without a running Streamlit server.
"""

import sys
import types
import os
from unittest import mock
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# Streamlit mock — must be installed *before* pages.compare is imported
# ---------------------------------------------------------------------------

_st_mock = types.ModuleType("streamlit")
class _SessionState(dict):
    """Dict that also supports attribute access (like Streamlit's session_state)."""
    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            return None  # Streamlit returns None-ish for missing keys
    def __setattr__(self, name, value):
        self[name] = value
    def __delattr__(self, name):
        try:
            del self[name]
        except KeyError:
            pass
    def __contains__(self, key):
        return dict.__contains__(self, key)

_st_mock.session_state = _SessionState()
_st_mock.cache_data = lambda *a, **kw: (lambda f: f)  # no-op decorator
_st_mock.cache_resource = lambda *a, **kw: (lambda f: f)
_st_mock.set_page_config = lambda **kw: None
_st_mock.title = lambda *a, **kw: None
_st_mock.header = lambda *a, **kw: None
_st_mock.subheader = lambda *a, **kw: None
_st_mock.write = lambda *a, **kw: None
_st_mock.markdown = lambda *a, **kw: None
_st_mock.error = lambda *a, **kw: None
_st_mock.warning = lambda *a, **kw: None
_st_mock.info = lambda *a, **kw: None
_st_mock.success = lambda *a, **kw: None
_st_mock.spinner = lambda *a, **kw: mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None)
_st_mock.expander = lambda *a, **kw: mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None)
def _mock_columns(*a, **kw):
    n = 2
    if a and hasattr(a[0], '__len__'):
        n = len(a[0])
    elif a and isinstance(a[0], int):
        n = a[0]
    return [mock.MagicMock() for _ in range(n)]
_st_mock.columns = _mock_columns
_st_mock.sidebar = mock.MagicMock()
_st_mock.button = lambda *a, **kw: False
_st_mock.text_input = lambda *a, **kw: ""
_st_mock.selectbox = lambda *a, **kw: None
_st_mock.multiselect = lambda *a, **kw: []
_st_mock.checkbox = lambda *a, **kw: False
_st_mock.radio = lambda *a, **kw: None
_st_mock.file_uploader = lambda *a, **kw: None
_st_mock.progress = lambda *a, **kw: mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None)
_st_mock.toast = lambda *a, **kw: None
_st_mock.rerun = lambda: None
_st_mock.stop = lambda: None
_st_mock.context = mock.MagicMock()
_st_mock.context.headers = {}
_st_mock.tabs = lambda *a, **kw: [mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None) for _ in range(10)]
_st_mock.container = lambda *a, **kw: mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None)
_st_mock.empty = lambda: mock.MagicMock()
_st_mock.form = lambda *a, **kw: mock.MagicMock(__enter__=lambda s: None, __exit__=lambda s, *a: None)
_st_mock.form_submit_button = lambda *a, **kw: False
_st_mock.data_editor = lambda *a, **kw: None
_st_mock.dataframe = lambda *a, **kw: None
_st_mock.table = lambda *a, **kw: None
_st_mock.metric = lambda *a, **kw: None
_st_mock.divider = lambda: None
_st_mock.download_button = lambda *a, **kw: False
_st_mock.number_input = lambda *a, **kw: 0
_st_mock.text_area = lambda *a, **kw: ""
_st_mock.toggle = lambda *a, **kw: False
_st_mock.dialog = lambda *a, **kw: (lambda f: f)
_st_mock.fragment = lambda *a, **kw: (lambda f: f)

sys.modules["streamlit"] = _st_mock

# ---------------------------------------------------------------------------
# Minimal utils mock — just enough symbols for ``from utils import ...``
# ---------------------------------------------------------------------------

_APP_DIR = Path(__file__).resolve().parent.parent

_utils_mock = types.ModuleType("utils")
_utils_mock.browse_folder = lambda *a, **kw: None
_utils_mock.scan_folder = lambda *a, **kw: []
_utils_mock.get_sheet_names_fast = lambda *a, **kw: ["Sheet1"]
_utils_mock.find_header_row = lambda *a, **kw: 1
_utils_mock.find_subject_column = lambda *a, **kw: None
_utils_mock.find_status_column = lambda *a, **kw: None
_utils_mock.get_file_category = lambda *a, **kw: "unknown"
_utils_mock.FILENAME_RULES = {}
_utils_mock.IS_SERVER = False
_utils_mock.save_uploaded_files = lambda *a, **kw: []
_utils_mock.server_folder_browser = lambda *a, **kw: None
_utils_mock.convert_windows_path_for_server = lambda p: p
sys.modules["utils"] = _utils_mock


# ---------------------------------------------------------------------------
# Import the module under test — now safe with mocks in place
# ---------------------------------------------------------------------------

# Add the app root to sys.path so ``pages.compare`` resolves
sys.path.insert(0, str(_APP_DIR))


# ---------------------------------------------------------------------------
# Temp-file fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_csv(tmp_path):
    """Return a helper that writes a CSV and returns its path."""
    def _make(name, rows):
        """rows: list of lists; first row is the header."""
        import csv
        p = tmp_path / name
        with open(p, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)
        return str(p)
    return _make


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Return a helper that writes a minimal XLSX and returns its path."""
    def _make(name, rows, number_formats=None):
        """rows: list of lists; first row is the header.
        number_formats: optional dict {col_idx: excel_code} applied to data rows.
        """
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for r_idx, row in enumerate(rows, 1):
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if number_formats and r_idx > 1 and (c_idx - 1) in number_formats:
                    cell.number_format = number_formats[c_idx - 1]
        p = tmp_path / name
        wb.save(p)
        return str(p)
    return _make
