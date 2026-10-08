"""Shared utility functions for File Data Extractor."""
import pandas as pd
from pathlib import Path
import streamlit as st
import openpyxl
import os
import re
import functools


# ---------------------------------------------------------------------------
# Environment detection
# ---------------------------------------------------------------------------
def is_server_environment() -> bool:
    """Detect if running on Posit Connect or other headless server."""
    return bool(
        os.environ.get("RSTUDIO_CONNECT_HKEY")
        or os.environ.get("CONNECT_SERVER")
        or os.name != "nt"
    )


IS_SERVER = is_server_environment()

# Default starting directory for the server-side folder browser.
SERVER_BROWSE_START = os.environ.get("IFV_BROWSE_START", "/lillyce/")

# Mapping of Windows drive letters to server mount points.
# Used to convert paths from study_config.json when running on the server.
# Format: "DRIVE_LETTER:SERVER_PATH" (configurable via env var).
_DRIVE_MAP_RAW = os.environ.get("IFV_DRIVE_MAP", "Z:/lillyce")
DRIVE_TO_SERVER: dict[str, str] = {}
for _entry in _DRIVE_MAP_RAW.split(";"):
    _entry = _entry.strip()
    if ":" in _entry:
        _drive, _mount = _entry.split(":", 1)
        DRIVE_TO_SERVER[_drive.upper()] = _mount.rstrip("/")


def convert_windows_path_for_server(path: str) -> str:
    """Convert a Windows path (e.g. Z:\\qa\\...) to a server path (/lillyce/qa/...).

    Uses the DRIVE_TO_SERVER mapping.  Returns the path unchanged if no
    mapping matches or if not running on the server.
    """
    if not IS_SERVER or not path:
        return path
    path_fwd = path.replace("\\", "/")
    if len(path_fwd) >= 2 and path_fwd[1] == ":":
        drive = path_fwd[0].upper()
        mount = DRIVE_TO_SERVER.get(drive)
        if mount:
            return mount + path_fwd[2:]  # replace "Z:" with "/lillyce"
    return path


def _cleanup_upload_dir(dir_path: str) -> None:
    """Silently remove an ifv_upload_* temp directory."""
    import shutil
    try:
        if dir_path and os.path.isdir(dir_path) and os.path.basename(dir_path).startswith("ifv_upload_"):
            shutil.rmtree(dir_path, ignore_errors=True)
    except Exception:
        pass


def _purge_stale_upload_dirs(max_age_seconds: int = 3600) -> None:
    """Remove ifv_upload_* directories older than *max_age_seconds* from the
    system temp folder.  Catches orphans left by crashed sessions.
    Limits scanning to avoid blocking on large /tmp directories."""
    import tempfile
    import time
    tmp_root = tempfile.gettempdir()
    now = time.time()
    purged = 0
    try:
        for entry in os.scandir(tmp_root):
            if not entry.name.startswith("ifv_upload_"):
                continue
            if entry.is_dir():
                try:
                    age = now - entry.stat().st_mtime
                    if age > max_age_seconds:
                        _cleanup_upload_dir(entry.path)
                        purged += 1
                        if purged >= 20:  # cap to avoid long scans
                            break
                except Exception:
                    pass
    except Exception:
        pass


# Track the most-recently created upload dir so we can clean it up later.
_active_upload_dirs: list[str] = []


def _upload_fingerprint(uploaded_files) -> str:
    """Build a fingerprint from file names and sizes to detect changes."""
    parts = sorted(f"{uf.name}:{uf.size}" for uf in uploaded_files)
    return "|".join(parts)


def save_uploaded_files(uploaded_files) -> str:
    """Save Streamlit UploadedFile objects to a temp directory.

    Returns the absolute path of the temp directory so downstream code
    (scan_folder, openpyxl reads, etc.) can work with normal file paths.

    Skips re-saving if the same set of files was already saved (detected
    via a name+size fingerprint), making Streamlit reruns instant.
    Automatically cleans up old upload directories.
    """
    import atexit
    import tempfile

    # Fast path: if we already saved exactly these files, return the
    # existing directory without touching disk.
    fp = _upload_fingerprint(uploaded_files)
    cached = st.session_state.get("_ifv_upload_cache")
    if cached and cached.get("fp") == fp:
        cached_dir = cached["dir"]
        if os.path.isdir(cached_dir):
            return cached_dir

    # 1. Remove any previous upload dir from this process
    while _active_upload_dirs:
        _cleanup_upload_dir(_active_upload_dirs.pop())

    # 2. Purge stale upload dirs left by other/crashed sessions (background,
    #    limited to avoid blocking on large /tmp directories).
    _purge_stale_upload_dirs()

    # 3. Create a fresh temp directory and save the files
    upload_dir = tempfile.mkdtemp(prefix="ifv_upload_")
    for uf in uploaded_files:
        dest = os.path.join(upload_dir, uf.name)
        with open(dest, "wb") as f:
            f.write(uf.getbuffer())

    # 4. Track it for cleanup on next upload or process exit
    _active_upload_dirs.append(upload_dir)
    atexit.register(_cleanup_upload_dir, upload_dir)

    # 5. Cache the fingerprint so reruns skip the disk write
    st.session_state["_ifv_upload_cache"] = {"fp": fp, "dir": upload_dir}

    return upload_dir


# ---------------------------------------------------------------------------
# Server-side folder browser (Posit Connect)
# ---------------------------------------------------------------------------
def _list_subdirectories(path: str) -> tuple[list[str], str | None]:
    """Return (sorted list of subdirectory names, error_message or None)."""
    try:
        dirs = []
        with os.scandir(path) as entries:
            for entry in entries:
                try:
                    if entry.is_dir(follow_symlinks=True):
                        dirs.append(entry.name)
                except OSError:
                    pass
        return sorted(dirs, key=str.lower), None
    except PermissionError:
        return [], f"Permission denied: {path}"
    except FileNotFoundError:
        return [], f"Path not found: {path}"
    except OSError as e:
        return [], f"Cannot read directory: {e}"


@st.dialog("Browse Server Folder", width="large")
def server_folder_browser(browser_id: str, target_keys: list[str],
                          start_dir: str | None = None) -> None:
    """Modal dialog that lets users navigate the server filesystem.

    Parameters
    ----------
    browser_id : str
        Unique id for this browser instance (namespaces session-state keys).
    target_keys : list[str]
        Session-state keys to set when the user confirms a selection.
    start_dir : str, optional
        Initial directory to show.  Defaults to ``SERVER_BROWSE_START``.
        Falls back to ``/`` if the path is not accessible.
    """
    if start_dir is None:
        start_dir = SERVER_BROWSE_START

    cwd_key = f"_fb_{browser_id}_cwd"
    nav_key = f"_fb_{browser_id}_nav"  # manual path input

    # Initialise current directory (fall back to / if start_dir not readable)
    if cwd_key not in st.session_state:
        if os.path.isdir(start_dir):
            st.session_state[cwd_key] = start_dir.rstrip("/") + "/"
        else:
            st.session_state[cwd_key] = "/"

    cwd = st.session_state[cwd_key]

    # --- Breadcrumb navigation ---
    parts = [p for p in cwd.split("/") if p]
    crumb_md = "📂  "
    crumb_md += "`/`"
    for part in parts:
        crumb_md += f" › `{part}`"
    st.markdown(crumb_md)

    # Breadcrumb buttons (clickable to jump to any level)
    n_crumbs = 1 + len(parts)  # root + segments
    if n_crumbs > 1:
        cols = st.columns(min(n_crumbs, 8))
        # Root button
        with cols[0]:
            if st.button("/", key=f"_fb_{browser_id}_crumb_0",
                         use_container_width=True):
                st.session_state[cwd_key] = "/"
                st.rerun()
        for i, part in enumerate(parts):
            with cols[(i + 1) % len(cols)]:
                if st.button(f"{part}/", key=f"_fb_{browser_id}_crumb_{i+1}",
                             use_container_width=True):
                    st.session_state[cwd_key] = "/" + "/".join(parts[:i+1]) + "/"
                    st.rerun()

    st.divider()

    # --- Manual path input ---
    st.caption("Type or paste a path and click **Go** to navigate directly:")
    path_col, go_col = st.columns([5, 1])
    with path_col:
        manual_path = st.text_input(
            "Path", value=cwd, key=nav_key, label_visibility="collapsed",
        )
    with go_col:
        go_clicked = st.button("Go", key=f"_fb_{browser_id}_go",
                               use_container_width=True)

    if go_clicked and manual_path:
        normalised = manual_path.rstrip("/") + "/"
        if not os.path.isdir(manual_path):
            st.error("Path does not exist.")
        else:
            st.session_state[cwd_key] = normalised
            st.rerun()

    # --- Parent folder button ---
    if cwd.rstrip("/") != "":
        if st.button("⬆  Parent folder", key=f"_fb_{browser_id}_up",
                     use_container_width=True):
            parent = os.path.dirname(cwd.rstrip("/"))
            if not parent:
                parent = "/"
            else:
                parent = parent.rstrip("/") + "/"
            st.session_state[cwd_key] = parent
            st.rerun()

    # --- Directory listing ---
    with st.spinner("Loading directories…"):
        dirs, error = _list_subdirectories(cwd)

    if error:
        st.warning(error)
        st.info("💡 **Tip:** Type a full path above and click **Go** to navigate directly to an accessible folder.")
    elif not dirs:
        st.info("This folder contains no subdirectories.")
    else:
        if len(dirs) > 200:
            st.caption(f"Showing first 200 of {len(dirs)} subdirectories.")
            dirs = dirs[:200]
        # Render directory buttons in two columns for better use of space
        col_a, col_b = st.columns(2)
        for idx, d in enumerate(dirs):
            with (col_a if idx % 2 == 0 else col_b):
                if st.button(f"📁 {d}/", key=f"_fb_{browser_id}_d_{idx}",
                             use_container_width=True):
                    st.session_state[cwd_key] = os.path.join(cwd, d) + "/"
                    st.rerun()

    # --- Select / Cancel ---
    st.divider()
    sel_col, cancel_col = st.columns(2)
    with sel_col:
        if st.button("✅ Select this folder", type="primary",
                     key=f"_fb_{browser_id}_select", use_container_width=True):
            for tkey in target_keys:
                st.session_state[tkey] = cwd.rstrip("/")
            st.session_state[f"_fb_{browser_id}_open"] = False
            # Clean up browser state
            st.session_state.pop(cwd_key, None)
            st.session_state.pop(nav_key, None)
            st.rerun()
    with cancel_col:
        if st.button("Cancel", key=f"_fb_{browser_id}_cancel",
                     use_container_width=True):
            st.session_state[f"_fb_{browser_id}_open"] = False
            st.session_state.pop(cwd_key, None)
            st.session_state.pop(nav_key, None)
            st.rerun()


def browse_folder() -> str:
    """Open a native Windows folder picker dialog and return the selected path.

    Uses a subprocess to avoid the 'main thread is not in main loop' error
    that occurs when tkinter runs inside Streamlit's threaded environment.
    """
    import subprocess
    import sys

    script = (
        "import tkinter as tk\n"
        "from tkinter import filedialog\n"
        "root = tk.Tk()\n"
        "root.withdraw()\n"
        "root.wm_attributes('-topmost', 1)\n"
        "folder = filedialog.askdirectory(title='Select Folder')\n"
        "root.destroy()\n"
        "print(folder if folder else '')\n"
    )
    try:
        result = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=120
        )
        folder = result.stdout.strip()
        if folder:
            return os.path.normpath(folder)
    except (subprocess.TimeoutExpired, Exception):
        pass
    return ""


def convert_path_to_unix(folder_path: str) -> str:
    """Convert Windows network/mapped drive path to Unix-style /lillyce path."""
    import subprocess
    folder_path = folder_path.replace("\\", "/")

    if folder_path.startswith("//"):
        return "/" + folder_path.lstrip("/")

    if len(folder_path) >= 2 and folder_path[1] == ":":
        drive_letter = folder_path[0].upper() + ":"
        try:
            result = subprocess.run(
                ["net", "use", drive_letter],
                capture_output=True, text=True, timeout=5
            )
            for line in result.stdout.splitlines():
                line_stripped = line.strip()
                if line_stripped.startswith("Remote name"):
                    unc_path = line_stripped.split(None, 2)[-1].strip()
                    unc_path = unc_path.replace("\\", "/")
                    unix_base = "/" + unc_path.lstrip("/")
                    rest = folder_path[2:]
                    return unix_base + rest
        except Exception:
            pass

    return folder_path


def scan_folder(folder_path: str) -> list[str]:
    """Scan a folder and return all Excel/CSV file names."""
    supported_extensions = {".xlsx", ".xls", ".csv"}
    try:
        files = []
        with os.scandir(folder_path) as entries:
            for entry in entries:
                if entry.is_file():
                    _, ext = os.path.splitext(entry.name)
                    if ext.lower() in supported_extensions:
                        files.append(entry.name)
        return sorted(files)
    except FileNotFoundError:
        raise FileNotFoundError(f"Folder not found: {folder_path}")
    except NotADirectoryError:
        raise NotADirectoryError(f"Path is not a folder: {folder_path}")
    except OSError:
        if not os.path.isdir(folder_path):
            if not os.path.exists(folder_path):
                raise FileNotFoundError(f"Folder not found: {folder_path}")
            raise NotADirectoryError(f"Path is not a folder: {folder_path}")
        raise


@functools.lru_cache(maxsize=256)
def get_sheet_names_fast(file_path: str) -> list[str]:
    """Get sheet names from an Excel file using the fastest method available."""
    path = Path(file_path)
    try:
        if path.suffix.lower() == ".xls":
            import xlrd
            wb = xlrd.open_workbook(file_path, on_demand=True)
            names = wb.sheet_names()
            wb.release_resources()
            return names
        else:
            import zipfile
            from xml.etree import ElementTree
            try:
                with zipfile.ZipFile(file_path, "r") as zf:
                    with zf.open("xl/workbook.xml") as wbxml:
                        tree = ElementTree.parse(wbxml)
                        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                        sheets = tree.findall(".//s:sheet", ns)
                        names = [s.get("name") for s in sheets if s.get("name")]
                        if names:
                            return names
            except Exception:
                pass
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
            sheets = wb.sheetnames
            wb.close()
            return sheets if sheets else ["Sheet1"]
    except Exception as e:
        return [f"ERROR: {str(e)[:50]}"]


@functools.lru_cache(maxsize=256)
def find_header_row(file_path: str, sheet_name: str) -> tuple[int, list]:
    """Find the row number where actual data column headers are present (1-indexed).

    Uses max_col=500 to avoid trusting the sheet's stale <dimension> tag which
    can under-report the true column count in read_only mode.
    """
    path = Path(file_path)
    try:
        if path.suffix.lower() == ".xls":
            import xlrd
            wb = xlrd.open_workbook(file_path, on_demand=True)
            ws = wb.sheet_by_name(sheet_name)
            for row_idx in range(min(30, ws.nrows)):
                row_values = [ws.cell_value(row_idx, col) for col in range(ws.ncols)]
                if _is_header_row(row_values):
                    wb.release_resources()
                    return row_idx + 1, row_values
            wb.release_resources()
            return 1, []
        else:
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
            ws = wb[sheet_name]
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, max_col=100, values_only=True), start=1):
                row_values = list(row)
                # Strip trailing None values (from the wide max_col scan)
                while row_values and row_values[-1] is None:
                    row_values.pop()
                if _is_header_row(row_values):
                    wb.close()
                    return row_idx, row_values
            wb.close()
            return 1, []
    except Exception as e:
        print(f"find_header_row failed for {file_path}/{sheet_name}: {e}")
        return 1, []


def _is_header_row(row_values: list) -> bool:
    """Determine if a row looks like a table header row."""
    non_empty = []
    for v in row_values:
        if v is not None and str(v).strip() != "":
            non_empty.append(str(v).strip())

    if len(non_empty) < 4:
        return False

    header_like = 0
    for val in non_empty:
        try:
            float(val)
            continue
        except (ValueError, TypeError):
            pass
        if len(val) < 40 and val.count(" ") < 6:
            header_like += 1

    return header_like >= 4 and header_like >= len(non_empty) * 0.5


def find_subject_column(headers: list) -> str:
    """Find the column header that represents subject/patient identifier."""
    subject_keywords = ["subject", "subjid", "subj_id", "subj id", "patient", "participant"]
    for header in headers:
        if header is None:
            continue
        header_str = str(header).strip()
        if not header_str:
            continue
        header_lower = header_str.lower()
        for keyword in subject_keywords:
            if keyword in header_lower:
                return header_str
    return ""


def find_status_column(headers: list) -> str:
    """Find the column header that represents status."""
    status_keywords = [
        "issue status", "issue_status", "form > marked", "form > mark",
        "marked for removal", "form_progress", "form progress",
        "reconciliation", "recon status", "status",
    ]
    for keyword in status_keywords:
        for header in headers:
            if header is None:
                continue
            header_str = str(header).strip()
            if not header_str:
                continue
            if keyword in header_str.lower():
                return header_str
    return ""


# --- Filename classification rules ---
FILENAME_RULES = [
    (["labfix", "lab","labfix_lab"], "Labfix_LAB", "labfix_lab_cnt", "labfix_lab_issue"),
    (["ecg", "labfix_ecg"], "Labfix_ECG", "labfix_ecg_cnt", "labfix_ecg_issue"),
    (["data issue", "data_issue", "dm issue", "dm_issue", "cra issue", "cra_issue",
      "dm data", "dm_data", "issue list", "issue_list", "cra follow", "cra_follow",
      "follow up tracker", "follow_up_tracker", "followup tracker"],
     "DM_Data_Issue_List", "dil_cnt", "data_issue_list_issue"),
    (["sae", "reconciliation", "recon"], "SAE", "sae_cnt", "sae_issue"),
    (["biometric", "biostats", "biostat", "stats"], "Biometrics", "bio_cnt", "biometrics_issue"),
    (["form progress", "form_progress", "form mark", "form_mark",
      "marked for removal", "form requiring", "form reset", "form_reset"],
     "Form_reset", "form_reset_cnt", "form_reset"),
    (["ecoa dcr", "ecoa_dcr", "dcr report", "dcr_report", "ecoa dcf", "ecoa_dcf", "dcf"],
     "ECOA_DCR", "ecoa_dcr_cnt", "ecoa_dcr_issue"),
    (["ecoa", "ecoa_issue", "ecoa issue", "ecoa dashboard", "ecoa_dashboard",
      "ecoa comment", "ecoa_comment", "ecoa recon", "ecoa_recon"],
     "ECOA_Issue", "ecoa_iss_cnt", "ecoa_issue"),
    (["pc tracker", "pc_tracker", "pc "], "PC_Tracker", "pc_cnt", "pc_issue"),
]


def classify_filename(filename: str) -> tuple[str, str, str]:
    """Classify a filename into (Files, output_dataset, output_var_name)."""
    name = Path(filename).stem
    name_lower = name.lower()

    cleaned = re.sub(r"^(?:[A-Z0-9]{2,}[-_])+", "", name, flags=re.IGNORECASE)
    if not cleaned or len(cleaned) < 3:
        cleaned = name
    cleaned_lower = cleaned.lower()

    for patterns, files_val, dataset_val, varname_val in FILENAME_RULES:
        for pattern in patterns:
            if pattern in name_lower or pattern in cleaned_lower:
                return files_val, dataset_val, varname_val

    fallback = re.sub(r"[^a-z0-9]+", "_", cleaned_lower).strip("_")
    if len(fallback) > 25:
        parts = fallback.split("_")
        fallback = "_".join(parts[:3])

    files_val = fallback.replace("_", " ").title().replace(" ", "_")
    return files_val, f"{fallback}_cnt", f"{fallback}_issue"


def get_file_category(filename: str) -> str:
    """Get just the category name (Files column) for a file."""
    files_val, _, _ = classify_filename(filename)
    return files_val
