"""Compare page — Compare files against template, then place them."""
import streamlit as st
import pandas as pd
from pathlib import Path
import openpyxl
import warnings
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
warnings.filterwarnings("ignore", message="Workbook contains no default style")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=pd.errors.DtypeWarning)
import os
import shutil
import sys
import datetime
import re
import unicodedata
import html as html_mod
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
import functools
import numpy as np
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils import (
    browse_folder, scan_folder, get_sheet_names_fast, find_header_row,
    find_subject_column, find_status_column, get_file_category, FILENAME_RULES,
    IS_SERVER, save_uploaded_files,
)

# --- Shared date format constants ---
# Mapping: format label → Python strftime format string
_DATE_FMT_LABEL_TO_STRFTIME = {
    "YYYY-MM-DD": "%Y-%m-%d",
    "YYYY-MM-DD HH:MM:SS": "%Y-%m-%d %H:%M:%S",
    "M/D/YYYY": "%-m/%-d/%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "DD-MMM-YYYY": "%d-%b-%Y",
    "DD-MMMM-YYYY": "%d-%B-%Y",
}
# Apply Windows-specific strftime fix at import time
if os.name == "nt":
    _DATE_FMT_LABEL_TO_STRFTIME = {k: v.replace("%-", "%#") for k, v in _DATE_FMT_LABEL_TO_STRFTIME.items()}

# Format labels where day comes before month (affects ambiguous date parsing)
_DAY_FIRST_LABELS = {"DD-MM-YYYY", "DD/MM/YYYY", "DD.MM.YYYY", "DD-MMM-YYYY", "DD-MMMM-YYYY"}
# Mapping: format label → valid Excel number_format code
_DATE_FMT_LABEL_TO_EXCEL = {
    "YYYY-MM-DD": "yyyy-mm-dd",
    "YYYY-MM-DD HH:MM:SS": "yyyy-mm-dd hh:mm:ss",
    "M/D/YYYY": "m/d/yyyy",
    "DD-MM-YYYY": "dd-mm-yyyy",
    "DD-MMM-YYYY": "dd-mmm-yyyy",
    "DD-MMMM-YYYY": "dd-mmmm-yyyy",
}
_DATE_FMT_LABEL_TO_STRFTIME.update({
    "YYYY/MM/DD": "%Y/%m/%d", "YYYY.MM.DD": "%Y.%m.%d", "DD.MM.YYYY": "%d.%m.%Y",
    "MM-DD-YYYY": "%m-%d-%Y", "MM/DD/YYYY": "%m/%d/%Y", "MM.DD.YYYY": "%m.%d.%Y",
    "DD/MM/YYYY": "%d/%m/%Y",
})
_DATE_FMT_LABEL_TO_EXCEL.update({
    "YYYY/MM/DD": "yyyy/mm/dd", "YYYY.MM.DD": "yyyy.mm.dd", "DD.MM.YYYY": "dd.mm.yyyy",
    "MM-DD-YYYY": "mm-dd-yyyy", "MM/DD/YYYY": "mm/dd/yyyy", "MM.DD.YYYY": "mm.dd.yyyy",
    "DD/MM/YYYY": "dd/mm/yyyy",
})
# Auto-generate the " HH:MM:SS" time variant of every base label
for _b, _p in list(_DATE_FMT_LABEL_TO_STRFTIME.items()):
    _DATE_FMT_LABEL_TO_STRFTIME.setdefault(f"{_b} HH:MM:SS", _p if "%H" in _p else _p + " %H:%M:%S")
for _b, _e in list(_DATE_FMT_LABEL_TO_EXCEL.items()):
    _DATE_FMT_LABEL_TO_EXCEL.setdefault(f"{_b} HH:MM:SS", _e if "hh" in _e.lower() else _e + " hh:mm:ss")
_TEXT_DATE_LABELS = set(_DATE_FMT_LABEL_TO_STRFTIME.keys())

# Reverse mapping: Excel number_format code → our label (for normalizing XLSX detected formats)
# Use setdefault so the FIRST (correct) entry wins and auto-expanded duplicates don't overwrite
_EXCEL_FMT_TO_LABEL = {}
for _lbl, _excel_code in _DATE_FMT_LABEL_TO_EXCEL.items():
    _EXCEL_FMT_TO_LABEL.setdefault(_excel_code, _lbl)
    _EXCEL_FMT_TO_LABEL.setdefault(_excel_code.lower(), _lbl)
# Also map common variants openpyxl might report
# IMPORTANT: Delimiters (/ vs - vs .) MUST be preserved in the label.
# "dd/mm/yyyy" -> "DD/MM/YYYY" (slash), "dd-mm-yyyy" -> "DD-MM-YYYY" (dash).
_EXCEL_FMT_TO_LABEL.update({
    "yyyy\\-mm\\-dd": "YYYY-MM-DD",
    "yyyy-mm-dd;@": "YYYY-MM-DD",
    "m/d/yy": "M/D/YYYY",
    "m/d/yy h:mm": "M/D/YYYY HH:MM:SS",           # Excel built-in format 22
    "m/d/yy hh:mm": "M/D/YYYY HH:MM:SS",
    "m/d/yy h:mm:ss": "M/D/YYYY HH:MM:SS",
    "m/d/yy hh:mm:ss": "M/D/YYYY HH:MM:SS",
    "mm/dd/yy": "M/D/YYYY",
    "mm/dd/yy h:mm": "M/D/YYYY HH:MM:SS",
    "mm/dd/yy hh:mm": "M/D/YYYY HH:MM:SS",
    "mm/dd/yy h:mm:ss": "M/D/YYYY HH:MM:SS",
    "mm/dd/yyyy": "M/D/YYYY",
    "mm/dd/yyyy h:mm": "M/D/YYYY HH:MM:SS",
    "mm/dd/yyyy hh:mm": "M/D/YYYY HH:MM:SS",
    "d/m/yyyy": "DD/MM/YYYY",
    "d/m/yy": "DD/MM/YYYY",
    "d/m/yy h:mm": "DD/MM/YYYY HH:MM:SS",
    "d/m/yyyy h:mm": "DD/MM/YYYY HH:MM:SS",
    "d/m/yyyy h:mm:ss": "DD/MM/YYYY HH:MM:SS",
    "d-mmm-yyyy": "DD-MMM-YYYY",
    "d-mmm-yy": "DD-MMM-YYYY",
    "dd/mm/yyyy": "DD/MM/YYYY",
    "dd/mm/yy": "DD/MM/YYYY",
    "dd/mm/yy h:mm": "DD/MM/YYYY HH:MM:SS",
    "dd/mm/yy hh:mm": "DD/MM/YYYY HH:MM:SS",
    "d-m-yyyy": "DD-MM-YYYY",
    "d-m-yy": "DD-MM-YYYY",
    "yyyy-mm-dd h:mm:ss": "YYYY-MM-DD HH:MM:SS",
    "yyyy-mm-dd hh:mm:ss": "YYYY-MM-DD HH:MM:SS",
    "yyyy-mm-dd hh:mm": "YYYY-MM-DD HH:MM:SS",
    "yyyy-mm-dd h:mm": "YYYY-MM-DD HH:MM:SS",
    "yyyy/mm/dd h:mm:ss": "YYYY/MM/DD HH:MM:SS",
    "yyyy/mm/dd hh:mm:ss": "YYYY/MM/DD HH:MM:SS",
    "yyyy/mm/dd hh:mm": "YYYY/MM/DD HH:MM:SS",
    "yyyy/mm/dd h:mm": "YYYY/MM/DD HH:MM:SS",
    "m/d/yyyy h:mm": "M/D/YYYY HH:MM:SS",
    "m/d/yyyy h:mm:ss": "M/D/YYYY HH:MM:SS",
    "m/d/yyyy hh:mm": "M/D/YYYY HH:MM:SS",
    "m/d/yyyy hh:mm:ss": "M/D/YYYY HH:MM:SS",
    "mm/dd/yyyy h:mm:ss": "M/D/YYYY HH:MM:SS",
    "mm/dd/yyyy hh:mm:ss": "M/D/YYYY HH:MM:SS",
    "mm/dd/yyyy hh:mm": "M/D/YYYY HH:MM:SS",
    "dd/mm/yyyy h:mm:ss": "DD/MM/YYYY HH:MM:SS",
    "dd/mm/yyyy hh:mm:ss": "DD/MM/YYYY HH:MM:SS",
    "dd/mm/yyyy hh:mm": "DD/MM/YYYY HH:MM:SS",
    "dd/mm/yyyy h:mm": "DD/MM/YYYY HH:MM:SS",
    "dd-mmm-yy": "DD-MMM-YYYY",
    "[$-en-us]m/d/yyyy": "M/D/YYYY",
    "[$-en-us]mm/dd/yyyy": "M/D/YYYY",
    "[$-en-us]m/d/yy h:mm": "M/D/YYYY HH:MM:SS",
})

# --- Canonical Date Signature System ---
# The canonical signature is the structural fingerprint of a date format.
# It preserves component order and delimiter but normalizes time precision.
# Two columns have a REAL format mismatch iff their date signatures differ.

def _canonical_date_signature(format_label):
    """Derive a canonical date-format signature from a format label.

    The signature captures:
      - Year/month/day component ORDER
      - Date SEPARATOR (- / . etc.) — delimiter differences ARE significant
      - Presence of time component (but seconds precision is NOT significant)

    Returns a tuple: (date_signature: str, has_time: bool)
      date_signature is the date-only portion (e.g., "YYYY-MM-DD", "MM/DD/YYYY")
      has_time indicates whether a time component is present

    For comparison purposes, only date_signature is used as the primary discriminant.
    Time precision (HH:MM vs HH:MM:SS) is normalized away.
    Leading-zero width (M vs MM, D vs DD) is normalized away.
    T vs space as datetime joiner is normalized away.

    Handles BOTH our internal label vocabulary (e.g. "M/D/YYYY HH:MM:SS")
    AND raw Excel number_format codes (e.g. "m/d/yy h:mm") as fallback.
    """
    if not format_label or format_label in ("General", "text-date", "unknown"):
        return (None, False)

    label = str(format_label).strip()

    # --- Step 1: If this looks like a raw Excel format, normalize it first ---
    # Raw Excel formats are lowercase with patterns like m/d/yy, dd-mm-yyyy, h:mm
    if label != label.upper() and any(c in label.lower() for c in ("yy", "mm", "dd")):
        normalized = _EXCEL_FMT_TO_LABEL.get(label.lower())
        if not normalized:
            # Strip locale prefix like [$-en-us] and try again
            stripped = re.sub(r"\[\$[^\]]*\]", "", label).strip()
            normalized = _EXCEL_FMT_TO_LABEL.get(stripped.lower())
        if normalized:
            label = normalized

    # --- Step 2: Detect time component ---
    # Match any time indicator: HH:MM, H:MM, hh:mm, h:mm (case-insensitive)
    has_time = bool(re.search(r"[Hh]{1,2}:[Mm]{2}", label)) or "HH:MM" in label

    # --- Step 3: Extract date-only portion ---
    # Strip time suffixes in our label format
    date_part = re.sub(r"\s*HH:MM(:SS)?$", "", label, flags=re.IGNORECASE).strip()
    # Strip raw Excel time suffixes (h:mm, hh:mm:ss, etc.)
    date_part = re.sub(r"\s*[Hh]{1,2}:[Mm]{2}(:[Ss]{2})?\s*(AM/PM|am/pm)?$", "", date_part).strip()
    # Handle "T" joiner
    date_part = date_part.rstrip("T").strip()

    # --- Step 4: Normalize raw Excel date patterns that survived ---
    # If date_part still looks like a raw Excel code (lowercase), normalize structurally
    if date_part and date_part[0].islower():
        # Parse raw Excel date format into our canonical form
        # Year-first patterns: yyyy-mm-dd, yyyy/mm/dd, yyyy.mm.dd
        m = re.match(r"^y{2,4}([-/.])m{1,2}\1d{1,2}$", date_part, re.IGNORECASE)
        if m:
            sep = m.group(1)
            date_part = f"YYYY{sep}MM{sep}DD"
        else:
            # Month-first: m/d/yy, mm/dd/yyyy
            m = re.match(r"^m{1,2}([-/.])d{1,2}\1y{2,4}$", date_part, re.IGNORECASE)
            if m:
                sep = m.group(1)
                date_part = f"MM{sep}DD{sep}YYYY"
            else:
                # Day-first: d/m/yyyy, dd-mm-yyyy, dd/mm/yy
                m = re.match(r"^d{1,2}([-/.])m{1,2}\1y{2,4}$", date_part, re.IGNORECASE)
                if m:
                    sep = m.group(1)
                    date_part = f"DD{sep}MM{sep}YYYY"
                else:
                    # Day-month_name-year: d-mmm-yyyy, dd-mmmm-yyyy
                    m = re.match(r"^d{1,2}([-/.])m{3,}([-/.])?y{2,4}$", date_part, re.IGNORECASE)
                    if m:
                        sep = m.group(1)
                        mlen = len(re.search(r"m+", date_part, re.IGNORECASE).group())
                        label_m = "MMMM" if mlen >= 4 else "MMM"
                        date_part = f"DD{sep}{label_m}{sep}YYYY"

    # --- Step 5: Normalize leading-zero width ---
    # "M/D/YYYY" and "MM/DD/YYYY" are structurally identical (zero-padding only)
    if date_part == "M/D/YYYY":
        date_part = "MM/DD/YYYY"
    elif date_part == "D/M/YYYY":
        date_part = "DD/MM/YYYY"

    return (date_part, has_time)


def _date_signatures_match(sig1, sig2):
    """Compare two canonical date signatures for structural equivalence.

    Returns True if the formats are structurally the same (no finding needed).
    Time precision differences (has_time vs not) are NOT flagged.
    Only the date_signature (component order + delimiter) matters.
    """
    date1, _ = sig1
    date2, _ = sig2

    # If either is None (General/unknown), can't compare — treat as match (no finding)
    if date1 is None or date2 is None:
        return True

    return date1 == date2


def _describe_signature_mismatch(p_sig, t_sig):
    """Generate a human-readable reason for a date signature mismatch.

    Returns a string describing why the signatures differ.
    """
    p_date, p_time = p_sig
    t_date, t_time = t_sig

    if p_date == t_date:
        return "Date format signatures are identical."

    # Determine the nature of the difference
    # Extract separators
    p_sep = _extract_date_separator(p_date)
    t_sep = _extract_date_separator(t_date)
    p_order = _extract_component_order(p_date)
    t_order = _extract_component_order(t_date)

    reasons = []
    if p_sep != t_sep:
        reasons.append(f'Date separator differs: "{p_sep}" vs "{t_sep}"')
    if p_order != t_order:
        reasons.append(f"Component order differs: {p_order} vs {t_order}")

    return "; ".join(reasons) if reasons else f"Format differs: {p_date} vs {t_date}"


def _extract_date_separator(date_sig):
    """Extract the date separator character from a signature like 'YYYY-MM-DD'."""
    if not date_sig:
        return ""
    for ch in date_sig:
        if ch in "-/.":
            return ch
    return ""


def _extract_component_order(date_sig):
    """Extract the component order (e.g., 'YMD', 'DMY', 'MDY') from a signature."""
    if not date_sig:
        return ""
    # Normalize to uppercase and extract component initials in order
    parts = re.split(r"[-/.]", date_sig.upper())
    order = ""
    for part in parts:
        if part.startswith("Y"):
            order += "Y"
        elif part.startswith("M"):
            order += "M"
        elif part.startswith("D"):
            order += "D"
    return order


# Step progress indicator (3 steps)
_step = 1
if st.session_state.get("cmp_done"):
    _step = 2
if st.session_state.get("cmp_fixed"):
    _step = 3
if st.session_state.get("cmp_placed_done"):
    _step = 3

# Compute progress parameters for premium UI
progress_pct = 0
if _step == 2:
    progress_pct = 50
elif _step == 3:
    progress_pct = 100

step1_bg = "linear-gradient(135deg, #0D9488 0%, #0F766E 100%)" if _step == 1 else ("#D1FAE5" if _step > 1 else "#FFFFFF")
step1_border = "#0D9488" if _step >= 1 else "#E2E8F0"
step1_color = "#FFFFFF" if _step == 1 else "#0D9488"
step1_shadow = "0 0 0 4px rgba(13, 148, 136, 0.2)" if _step == 1 else "none"
step1_fw = "700" if _step == 1 else "600"
step1_txt_color = "#FFFFFF" if _step == 1 else ("#2DD4BF" if _step > 1 else "#64748B")
step1_val = "✓" if _step > 1 else "1"

step2_bg = "linear-gradient(135deg, #0D9488 0%, #0F766E 100%)" if _step == 2 else ("#D1FAE5" if _step > 2 else "#1E293B")
step2_border = "#0D9488" if _step >= 2 else "#334155"
step2_color = "#FFFFFF" if _step == 2 else ("#0D9488" if _step > 2 else "#64748B")
step2_shadow = "0 0 0 4px rgba(13, 148, 136, 0.2)" if _step == 2 else "none"
step2_fw = "700" if _step == 2 else "500"
step2_txt_color = "#FFFFFF" if _step == 2 else ("#2DD4BF" if _step > 2 else "#64748B")
step2_val = "✓" if _step > 2 else "2"

step3_bg = "linear-gradient(135deg, #0D9488 0%, #0F766E 100%)" if _step == 3 else "#1E293B"
step3_border = "#0D9488" if _step == 3 else "#334155"
step3_color = "#FFFFFF" if _step == 3 else "#64748B"
step3_shadow = "0 0 0 4px rgba(13, 148, 136, 0.2)" if _step == 3 else "none"
step3_fw = "700" if _step == 3 else "500"
step3_txt_color = "#FFFFFF" if _step == 3 else "#64748B"
step3_val = "✓" if st.session_state.get("cmp_placed_done") else "3"

# Render vertical step progress in sidebar
with st.sidebar:
    st.markdown(f"""
    <div style="display: flex; flex-direction: column; align-items: center; padding: 2rem 0; position: relative; height: 65vh;">
      <!-- Vertical line connecting steps — positioned on the center of circles -->
      <div style="position: absolute; top: 55px; bottom: 55px; left: calc(18% + 22px); width: 3px; background-color: #1E293B; z-index: 1;">
        <div style="width: 100%; height: {progress_pct}%; background: linear-gradient(180deg, #0D9488, #2DD4BF); transition: height 0.4s ease;"></div>
      </div>

      <!-- Step 1 -->
      <div style="display: flex; align-items: center; z-index: 2; width: 100%; padding-left: 18%;">
        <div style="width: 44px; height: 44px; border-radius: 50%; background: {step1_bg}; border: 2px solid {step1_border}; color: {step1_color}; display: flex; align-items: center; justify-content: center; font-weight: 700; font-family: 'Plus Jakarta Sans', sans-serif; box-shadow: {step1_shadow}; font-size: 1.05rem; flex-shrink: 0;">{step1_val}</div>
        <span style="font-size: 0.88rem; font-weight: {step1_fw}; color: {step1_txt_color}; font-family: 'Plus Jakarta Sans', sans-serif; margin-left: 16px;">Select</span>
      </div>

      <!-- Spacer -->
      <div style="flex: 1;"></div>

      <!-- Step 2 -->
      <div style="display: flex; align-items: center; z-index: 2; width: 100%; padding-left: 18%;">
        <div style="width: 44px; height: 44px; border-radius: 50%; background: {step2_bg}; border: 2px solid {step2_border}; color: {step2_color}; display: flex; align-items: center; justify-content: center; font-weight: 700; font-family: 'Plus Jakarta Sans', sans-serif; box-shadow: {step2_shadow}; font-size: 1.05rem; flex-shrink: 0;">{step2_val}</div>
        <span style="font-size: 0.88rem; font-weight: {step2_fw}; color: {step2_txt_color}; font-family: 'Plus Jakarta Sans', sans-serif; margin-left: 16px;">Compare</span>
      </div>

      <!-- Spacer -->
      <div style="flex: 1;"></div>

      <!-- Step 3 -->
      <div style="display: flex; align-items: center; z-index: 2; width: 100%; padding-left: 18%;">
        <div style="width: 44px; height: 44px; border-radius: 50%; background: {step3_bg}; border: 2px solid {step3_border}; color: {step3_color}; display: flex; align-items: center; justify-content: center; font-weight: 700; font-family: 'Plus Jakarta Sans', sans-serif; box-shadow: {step3_shadow}; font-size: 1.05rem; flex-shrink: 0;">{step3_val}</div>
        <span style="font-size: 0.88rem; font-weight: {step3_fw}; color: {step3_txt_color}; font-family: 'Plus Jakarta Sans', sans-serif; margin-left: 16px;">Place</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

# Main page tagline
st.markdown("<h2 style='font-size:1.75rem; color:#0F172A; margin-top:-15px; margin-bottom:2.5rem; font-weight:700; font-family: Plus Jakarta Sans, sans-serif;'>Validate, compare & place study files against templates</h2>", unsafe_allow_html=True)

# --- Session State ---
for key, default in [
    ("cmp_excel_files", []), ("cmp_findings", []),
    ("cmp_placed_files", []), ("cmp_selected", []),
    ("cmp_fixed", False), ("cmp_done", False),
    ("cmp_rename_map", {}), ("cmp_placed_done", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# --- Generic Intelligent Filename Matching Engine ---
# No hard-coded dictionaries — works for ANY template/study automatically.

# Regex patterns for normalization
_RE_STUDY_PREFIX = re.compile(r"^(?:[A-Z0-9]{2,}[-_])+", re.NOFLAG)
_RE_DATE = re.compile(
    r"(?:\d{4}[-_]\d{2}[-_]\d{2})"       # 2024-01-15
    r"|(?:\d{2}[-_]\d{2}[-_]\d{4})"       # 15-01-2024
    r"|(?:\d{8})"                           # 20240115
    r"|(?:\d{1,2}\w{3}\d{2,4})"            # 15Jan2024
    r"|(?:\d{2}[-_]\d{2}[-_]\d{2,4}\s*\d{2}[-_]\d{2}[-_]\d{2})"  # timestamps
)
_RE_VERSION = re.compile(r"\b(?:v\d+\.\d+|ver\s*\d+|version\s*\d+)\b", re.IGNORECASE)
_RE_SUFFIX = re.compile(r"\b(?:copy|draft|final|rev|revised|updated|new|old|original|backup)\b", re.IGNORECASE)
_RE_EXEC_ID = re.compile(r"\b\d{6,}\b")  # Long numeric IDs (execution IDs, serial numbers)
_RE_TIMESTAMP = re.compile(r"\d{2}[_:]\d{2}[_:]\d{2}")  # HH_MM_SS or HH:MM:SS

_STOP_WORDS = frozenset({
    "the", "of", "and", "for", "in", "to", "a", "an", "on", "at", "by",
    "from", "with", "is", "are", "was", "were", "be", "been", "ist", "utc",
    "excel", "csv", "xlsx", "xls", "file", "sheet",
})


def normalize_filename(filename: str) -> str:
    """Normalize a filename by removing non-semantic elements.

    Pipeline: strip extension → strip study prefix → remove dates/timestamps →
    remove versions → remove copy/draft/final → remove execution IDs →
    lowercase → normalize separators.
    """
    name = Path(filename).stem
    # Strip study prefix — only ALL-CAPS/digit segments joined by - or _
    # Real study prefixes have multiple segments: I5T-MC-AACQ_, J2S-MC-GZME_
    # Single short segments like DIVA_, LABFIX_, SV1001_ are likely the file identity
    match = _RE_STUDY_PREFIX.match(name)
    if match:
        prefix = match.group(0)
        remainder = name[len(prefix):]
        # Count prefix segments
        _prefix_segments = re.findall(r"[A-Z0-9]{2,}", prefix)
        _n_segments = len(_prefix_segments)

        if not remainder or len(remainder) < 3:
            # Remainder too short — keep the whole name
            cleaned = name
        elif _n_segments <= 1:
            # Single-segment prefix (DIVA_, LABFIX_, SV1001_) — likely the identity, keep it
            cleaned = name
        else:
            # Multi-segment prefix (I5T-MC-AACQ_, J2S-MC-GZME_) — likely study ID
            # Check if remainder is ONLY dates/numbers — if so, prefix contains the identity
            _test_remainder = _RE_DATE.sub("", remainder)
            _test_remainder = _RE_TIMESTAMP.sub("", _test_remainder)
            _test_remainder = _RE_EXEC_ID.sub("", _test_remainder)
            _test_remainder = re.sub(r"[_\-\s\d]+", "", _test_remainder).strip()
            if not _test_remainder:
                # Remainder has no semantic content — keep the whole name
                cleaned = name
            else:
                cleaned = remainder
    else:
        cleaned = name
    # Remove dates, timestamps, versions, suffixes, execution IDs
    cleaned = _RE_DATE.sub(" ", cleaned)
    cleaned = _RE_TIMESTAMP.sub(" ", cleaned)
    cleaned = _RE_VERSION.sub(" ", cleaned)
    cleaned = _RE_SUFFIX.sub(" ", cleaned)
    cleaned = _RE_EXEC_ID.sub(" ", cleaned)
    # Normalize case and separators
    cleaned = cleaned.lower()
    cleaned = re.sub(r"[_\-\s]+", " ", cleaned).strip()
    # Split numbers glued to words: "visit3" → "visit 3"
    cleaned = re.sub(r"([a-z]{2,})(\d)", r"\1 \2", cleaned)
    cleaned = re.sub(r"(\d)([a-z]{2,})", r"\1 \2", cleaned)
    # Remove parenthetical content: "(88)", "(1)"
    cleaned = re.sub(r"\([^)]*\)", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _tokenize(normalized: str) -> list:
    """Tokenize a normalized filename into meaningful tokens."""
    tokens = normalized.split()
    result = []
    for t in tokens:
        # Strip leading/trailing punctuation
        t = t.strip(".-_,;:()")
        if len(t) < 2:
            continue
        if t in _STOP_WORDS:
            continue
        # Skip pure numbers (residual after date/ID removal)
        if t.isdigit():
            continue
        # Skip tokens that are mostly digits (e.g., ".656", "T091320")
        if sum(c.isdigit() for c in t) > len(t) * 0.6:
            continue
        result.append(t)
    return result


def _compute_token_weights(template_files: list) -> dict:
    """Compute IDF-like weights from template filenames.

    Tokens in many templates → low weight (common/generic).
    Tokens in few templates → high weight (distinguishing).
    """
    from collections import Counter
    token_doc_freq = Counter()
    n_docs = len(template_files)

    for tf in template_files:
        norm = normalize_filename(tf)
        tokens = set(_tokenize(norm))
        for t in tokens:
            token_doc_freq[t] += 1

    weights = {}
    for token, freq in token_doc_freq.items():
        # IDF-inspired: rare tokens get higher weight
        weights[token] = 1.0 / (1.0 + freq)

    return weights


def _lcs_ratio(a: list, b: list) -> float:
    """Longest common subsequence ratio (order-preserving similarity)."""
    if not a or not b:
        return 0.0
    m, n = len(a), len(b)
    # Optimized: only need previous row
    prev = [0] * (n + 1)
    for i in range(1, m + 1):
        curr = [0] * (n + 1)
        for j in range(1, n + 1):
            if a[i - 1] == b[j - 1]:
                curr[j] = prev[j - 1] + 1
            else:
                curr[j] = max(prev[j], curr[j - 1])
        prev = curr
    lcs_len = prev[n]
    return (2.0 * lcs_len) / (m + n)


def _similarity_score(placed_tokens: list, template_tokens: list, weights: dict) -> float:
    """Compute multi-signal similarity score (0.0–1.0) between two token lists."""
    if not placed_tokens and not template_tokens:
        return 1.0
    if not placed_tokens or not template_tokens:
        return 0.0

    p_set = set(placed_tokens)
    t_set = set(template_tokens)
    intersection = p_set & t_set
    union = p_set | t_set

    def _weight(token):
        return weights.get(token, 0.5)

    # Signal 1: Weighted Jaccard (45%)
    w_inter = sum(_weight(t) for t in intersection)
    w_union = sum(_weight(t) for t in union)
    weighted_jaccard = w_inter / w_union if w_union > 0 else 0.0

    # Signal 2: Containment — how much of the smaller set is in the larger (30%)
    smaller = p_set if len(p_set) <= len(t_set) else t_set
    w_smaller = sum(_weight(t) for t in smaller)
    containment = w_inter / w_smaller if w_smaller > 0 else 0.0

    # Signal 3: Sequence similarity — LCS ratio (15%)
    sequence = _lcs_ratio(placed_tokens, template_tokens)

    # Signal 4: Length penalty (10%)
    len_a, len_b = len(placed_tokens), len(template_tokens)
    length_sim = 1.0 - abs(len_a - len_b) / max(len_a, len_b, 1)

    # Weighted combination
    score = (0.45 * weighted_jaccard +
             0.30 * containment +
             0.15 * sequence +
             0.10 * length_sim)

    return score


def match_files(placed_files, template_files):
    """Match placed files to template files using intelligent similarity scoring.

    Fully generic — works for ANY filenames without hard-coded dictionaries.
    Uses token frequency analysis, weighted Jaccard, containment, and sequence similarity.

    Returns: dict mapping placed_filename -> template_filename or None
    """
    if not template_files:
        return {pf: None for pf in placed_files}

    # Pre-compute: normalize and tokenize all files
    tpl_data = {}
    for tf in template_files:
        norm = normalize_filename(tf)
        tokens = _tokenize(norm)
        tpl_data[tf] = {"norm": norm, "tokens": tokens}

    placed_data = {}
    for pf in placed_files:
        norm = normalize_filename(pf)
        tokens = _tokenize(norm)
        placed_data[pf] = {"norm": norm, "tokens": tokens}

    # Compute token weights from template corpus
    weights = _compute_token_weights(template_files)

    # --- Pass 1: Exact normalized match (after normalization, identical = perfect match) ---
    matches = {}
    unmatched_placed = []

    tpl_norms = {tf: tpl_data[tf]["norm"] for tf in template_files}
    tpl_norm_lookup = {}
    for tf, norm in tpl_norms.items():
        tpl_norm_lookup.setdefault(norm, tf)  # First template with this norm wins

    for pf in placed_files:
        pf_norm = placed_data[pf]["norm"]
        if pf_norm in tpl_norm_lookup:
            matches[pf] = tpl_norm_lookup[pf_norm]
        else:
            unmatched_placed.append(pf)

    # --- Pass 2: Similarity scoring for unmatched files ---
    for pf in unmatched_placed:
        p_tokens = placed_data[pf]["tokens"]
        best_match = None
        best_score = 0.0

        for tf in template_files:
            t_tokens = tpl_data[tf]["tokens"]
            score = _similarity_score(p_tokens, t_tokens, weights)
            if score > best_score:
                best_score = score
                best_match = tf

        # Check raw token overlap (unweighted Jaccard) to prevent false matches
        # between files that only share generic tokens (e.g. "progress").
        # Require at least 50% of unique tokens to be shared.
        raw_jaccard_ok = False
        if best_match:
            p_set = set(p_tokens)
            t_set = set(tpl_data[best_match]["tokens"])
            union = p_set | t_set
            intersection = p_set & t_set
            raw_jaccard = len(intersection) / len(union) if union else 0.0
            raw_jaccard_ok = raw_jaccard >= 0.5

        # Threshold: weighted score >= 0.35 AND raw token overlap >= 50%
        # Both conditions prevent matching files that share only generic tokens
        # (e.g. Subject_Progress vs Form_Progress share only "progress" = 33% overlap)
        if best_score >= 0.35 and best_match and raw_jaccard_ok:
            matches[pf] = best_match
        else:
            # Fallback: category-based matching — files that classify to the same
            # FILENAME_RULES group (e.g. ECOA DCF/ECOA DCR/DCR_Report are all "ECOA_DCR")
            pf_category = get_file_category(pf)
            category_match = None
            for tf in template_files:
                if tf not in matches.values():
                    if get_file_category(tf) == pf_category:
                        category_match = tf
                        break
            if category_match:
                matches[pf] = category_match
            else:
                matches[pf] = None

    return matches


# --- Comparison Logic ---

@functools.lru_cache(maxsize=256)
def get_data_sheet(file_path):
    """Get the sheet containing data headers."""
    sheets = get_sheet_names_fast(file_path)
    if not sheets or sheets[0].startswith("ERROR"):
        return ""
    for sheet in sheets:
        _, headers = find_header_row(file_path, sheet)
        if headers and len([h for h in headers if h is not None and str(h).strip()]) >= 4:
            return sheet
    return sheets[0] if sheets else ""


def get_header_columns(file_path, sheet_name):
    """Get (header_row_number, list_of_all_column_values_including_None)."""
    row_num, headers = find_header_row(file_path, sheet_name)
    # Return all headers (including None for empty cells) for accurate column counting
    return row_num, headers


def get_non_empty_columns(headers):
    """Get list of non-empty column name strings from headers (preserving spaces)."""
    return [str(h) for h in headers if h is not None and str(h).strip()]


@functools.lru_cache(maxsize=128)
def detect_column_types(file_path, sheet_name, hdr_row, max_sample=15):
    """Detect data types for each column by sampling rows below the header.

    Uses openpyxl cell.data_type for reliable type detection:
      'n' = numeric, 's' = string/text, 'd' = date, 'b' = boolean
    Also checks number_format for date detection on numeric cells.
    Captures the actual number_format string for date format comparison.
    """
    types = {}
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
        ws = wb[sheet_name]
        headers = []
        samples = {}

        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=hdr_row + max_sample, max_col=100), 1):
            if row_idx == hdr_row:
                headers = [(cell.value, i) for i, cell in enumerate(row) if cell.value is not None and str(cell.value).strip()]
            elif row_idx > hdr_row:
                for i, cell in enumerate(row):
                    col_i = i
                    if cell.value is not None and str(cell.value).strip():
                        if col_i not in samples:
                            samples[col_i] = []
                        samples[col_i].append({
                            "value": cell.value,
                            "format": cell.number_format,
                            "dtype": cell.data_type,
                        })

        wb.close()

        for hdr_val, col_i in headers:
            if hdr_val is None or not str(hdr_val).strip():
                continue
            col_samples = samples.get(col_i, [])
            if not col_samples:
                types[col_i] = {"header": str(hdr_val).strip(), "type": "unknown", "format": "General"}
                continue

            date_count = num_count = text_count = 0
            fmt = "General"
            detected_formats = []

            for s in col_samples:
                val = s["value"]
                nfmt = s["format"]
                cell_dtype = s["dtype"]

                def _normalize_excel_fmt(raw_fmt):
                    """Normalize an Excel number_format to our internal label."""
                    if not raw_fmt or raw_fmt == "General":
                        return None
                    key = raw_fmt.lower().strip()
                    # Direct lookup
                    result = _EXCEL_FMT_TO_LABEL.get(key)
                    if result:
                        return result
                    # Strip locale prefix like [$-en-US] or [$-x-sysdate]
                    stripped = re.sub(r"\[\$[^\]]*\]", "", key).strip()
                    result = _EXCEL_FMT_TO_LABEL.get(stripped)
                    if result:
                        return result
                    # Strip trailing ";@" or ";..." (alternate format sections)
                    stripped2 = stripped.split(";")[0].strip()
                    result = _EXCEL_FMT_TO_LABEL.get(stripped2)
                    if result:
                        return result
                    # Strip escaped chars (backslash before separators)
                    stripped3 = stripped2.replace("\\-", "-").replace("\\/", "/").replace("\\.", ".")
                    result = _EXCEL_FMT_TO_LABEL.get(stripped3)
                    return result

                # Check Python type FIRST — most reliable indicator regardless of cell_dtype
                if isinstance(val, (datetime.datetime, datetime.date)):
                    date_count += 1
                    if nfmt and nfmt != "General":
                        normalized = _normalize_excel_fmt(nfmt)
                        if normalized:
                            detected_formats.append(normalized)
                        else:
                            # Fallback: check if it's a datetime (has time) vs date-only
                            if isinstance(val, datetime.datetime) and (val.hour or val.minute or val.second):
                                detected_formats.append("YYYY-MM-DD HH:MM:SS")
                            else:
                                detected_formats.append("YYYY-MM-DD")
                    else:
                        if isinstance(val, datetime.datetime) and (val.hour or val.minute or val.second):
                            detected_formats.append("YYYY-MM-DD HH:MM:SS")
                        else:
                            detected_formats.append("YYYY-MM-DD")
                elif cell_dtype == "d":
                    date_count += 1
                    if nfmt and nfmt != "General":
                        normalized = _normalize_excel_fmt(nfmt)
                        if normalized:
                            detected_formats.append(normalized)
                        else:
                            detected_formats.append("YYYY-MM-DD")
                    else:
                        detected_formats.append("YYYY-MM-DD")
                elif cell_dtype == "n":
                    # Numeric — but check if it's a date stored as serial number
                    if nfmt and any(x in nfmt.lower() for x in ["yy", "mm", "dd", "m/d", "d-", "d/", "mmm"]):
                        date_count += 1
                        normalized = _normalize_excel_fmt(nfmt)
                        detected_formats.append(normalized if normalized else nfmt)
                    else:
                        num_count += 1
                elif cell_dtype == "s" or isinstance(val, str):
                    val_str = str(val).strip()
                    date_text_fmt = _detect_date_text_format(val_str)

                    if date_text_fmt:
                        date_count += 1
                        detected_formats.append(date_text_fmt)
                    else:
                        text_count += 1
                else:
                    if isinstance(val, (int, float)):
                        num_count += 1
                    else:
                        text_count += 1

            total = date_count + num_count + text_count
            if total == 0:
                dtype_result = "unknown"
            elif date_count > 0 and date_count >= num_count and date_count >= text_count:
                dtype_result = "date"
            elif date_count > 0 and num_count > 0 and text_count == 0:
                # Mix of date cells and numeric serial numbers → treat as date
                dtype_result = "date"
            elif num_count > 0 and num_count > text_count:
                dtype_result = "numeric"
            else:
                dtype_result = "text"

            # Determine the most common date format
            if dtype_result == "date" and detected_formats:
                # Pick most frequent format
                fmt_counts = Counter(detected_formats)
                fmt = fmt_counts.most_common(1)[0][0]
                # Final normalization: if format is still an Excel code, try to convert
                if fmt not in _TEXT_DATE_LABELS and fmt != "General":
                    fmt = _EXCEL_FMT_TO_LABEL.get(fmt.lower(), fmt)
                # Format consistency tracking
                consistency = fmt_counts.most_common(1)[0][1] / len(detected_formats) if detected_formats else 0.0
                types[col_i] = {"header": str(hdr_val).strip(), "type": dtype_result, "format": fmt,
                                "format_consistency": round(consistency, 2),
                                "format_variants": len(fmt_counts)}
            elif dtype_result != "date":
                fmt = "General"
                types[col_i] = {"header": str(hdr_val).strip(), "type": dtype_result, "format": fmt}
            else:
                types[col_i] = {"header": str(hdr_val).strip(), "type": dtype_result, "format": fmt}

    except Exception as e:
        print(f"detect_column_types failed for {file_path}/{sheet_name}: {e}")
    return types



def _get_csv_headers(file_path):
    """Get headers from a CSV file. Returns (1, list_of_headers)."""
    try:
        df = pd.read_csv(file_path, nrows=0)
        return 1, list(df.columns)
    except Exception:
        return 1, []


def _is_numeric_str(s):
    """Check if a string represents a numeric value."""
    try:
        float(s.replace(",", ""))
        return True
    except (ValueError, TypeError):
        return False


@functools.lru_cache(maxsize=128)
def _detect_csv_column_types(file_path, max_sample=15):
    """Detect data types for CSV columns using the same format-label logic as
    the template (openpyxl text-cell) detection.

    Uses the SAME regex patterns as _detect_types_uncached so that identical
    text-date formats in both files produce the same label — ensuring only REAL
    format differences trigger a finding.
    """

    types = {}
    try:
        # CRITICAL: Read as dtype=str to prevent pandas from auto-parsing
        # date columns (which would corrupt the original text format and make
        # format detection return wrong labels like "YYYY-MM-DD HH:MM:SS"
        # instead of the actual text format like "M/D/YYYY").
        df = pd.read_csv(file_path, nrows=50, dtype=str, low_memory=False)
        if df.empty:
            return {}

        for col_i, col_name in enumerate(df.columns):
            try:
                if not str(col_name).strip():
                    continue
                col_data = df[col_name].dropna()
                if col_data.empty:
                    types[col_i] = {"header": str(col_name).strip(), "type": "unknown", "format": "General"}
                    continue

                # --- For text columns: classify each value as date/numeric/text ---
                sample = col_data.head(max_sample).astype(str)
                date_count = 0
                num_count = 0
                text_count = 0
                detected_formats = []

                for val in sample:
                    val_str = str(val).strip()
                    if not val_str:
                        continue

                    date_text_fmt = _detect_date_text_format(val_str)
                    if date_text_fmt:
                        date_count += 1
                        detected_formats.append(date_text_fmt)
                    elif _is_numeric_str(val_str):
                        num_count += 1
                    else:
                        text_count += 1

                total = date_count + num_count + text_count
                if total == 0:
                    types[col_i] = {"header": str(col_name).strip(), "type": "unknown", "format": "General"}
                elif date_count > 0 and date_count >= num_count and date_count >= text_count:
                    # Date column — determine dominant format from detected patterns
                    fmt_counts = Counter(detected_formats)
                    fmt = fmt_counts.most_common(1)[0][0] if detected_formats else "General"
                    # Format consistency: ratio of dominant format to total detected
                    consistency = fmt_counts.most_common(1)[0][1] / len(detected_formats) if detected_formats else 0.0
                    types[col_i] = {"header": str(col_name).strip(), "type": "date", "format": fmt,
                                    "format_consistency": round(consistency, 2),
                                    "format_variants": len(fmt_counts)}
                elif num_count > 0 and num_count > text_count:
                    types[col_i] = {"header": str(col_name).strip(), "type": "numeric", "format": "General"}
                else:
                    # If regex didn't detect a recognizable date format, classify as text.
                    # The pd.to_datetime fallback is too liberal and causes false positives
                    # (e.g., assigning "YYYY-MM-DD HH:MM:SS" to non-standard date strings
                    # which then mismatch against the template's "YYYY-MM-DD" label).
                    types[col_i] = {"header": str(col_name).strip(), "type": "text", "format": "General"}
            except Exception:
                # Per-column exception: classify as text and continue with remaining columns
                types[col_i] = {"header": str(col_name).strip(), "type": "text", "format": "General"}
    except Exception:
        pass
    return types


def _detect_date_text_format(val_str):
    """Detect a date's format label from a text string.

    Precise about the DATE structure (component order + separator) so genuinely
    different date formats produce different labels, but LENIENT about the time
    portion: the date<->time joiner (T or space), a trailing 'Z'/offset timezone,
    and seconds precision are all ignored, so datetime cosmetic variants such as
    '2025-11-25T13:58Z' and '2025-11-25 13:58:00' still map to the same label.
    """
    s = str(val_str).strip()
    if not s:
        return None

    # Does the value carry a clock time? (joiner/seconds/timezone are irrelevant)
    has_time = bool(re.search(r"\d{1,2}:\d{2}", s))
    # Isolate the date portion (everything before the T/space that precedes a time)
    date_part = re.split(r"[T ]\d{1,2}:\d{2}", s, maxsplit=1)[0].strip() if has_time else s

    def _with_time(base):
        return f"{base} HH:MM:SS" if has_time else base

    # --- ISO-style, year first: YYYY<sep>MM<sep>DD (separator is significant) ---
    m = re.match(r"^\d{4}([-/.])\d{1,2}\1\d{1,2}$", date_part)
    if m:
        sep = m.group(1)
        return _with_time(f"YYYY{sep}MM{sep}DD")

    # --- Numeric day/month/year with a separator: resolve ORDER from the values ---
    m = re.match(r"^(\d{1,2})([-/.])(\d{1,2})\2\d{2,4}$", date_part)
    if m:
        f1, sep, f2 = int(m.group(1)), m.group(2), int(m.group(3))
        if f2 > 12 and f1 <= 12:
            first = "MM"                       # 2nd field must be a day -> month-first
        elif f1 > 12 and f2 <= 12:
            first = "DD"                       # 1st field must be a day -> day-first
        else:
            # ambiguous (both <= 12): slash defaults month-first (US), dash/dot day-first
            first = "MM" if sep == "/" else "DD"
        if sep == "/" and first == "MM":
            return _with_time("M/D/YYYY")      # reuse existing month-first slash label
        second = "DD" if first == "MM" else "MM"
        return _with_time(f"{first}{sep}{second}{sep}YYYY")

    # --- Month-name variants (all collapse to DD-MMM-YYYY / DD-MMMM-YYYY) ---
    if re.match(r"^\d{1,2}[-/ ]\w{3}[-/ ]\d{2,4}$", date_part):
        return _with_time("DD-MMM-YYYY")
    if re.match(r"^\d{1,2}[-/]\w{4,}[-/]\d{2,4}$", date_part):
        return _with_time("DD-MMMM-YYYY")
    if re.match(r"^\d{1,2}[A-Za-z]{3,}\d{2,4}$", date_part):        # 24Nov2025
        return _with_time("DD-MMM-YYYY")
    if re.match(r"^[A-Za-z]{3,}\s+\d{1,2},?\s+\d{4}$", date_part):  # Nov 24, 2025
        return _with_time("DD-MMM-YYYY")

    return None


def _parse_date_with_format(value, format_label):
    """Parse a raw date value using the detected format label as parsing context.

    Handles: datetime objects, Excel serial numbers, and text strings.
    Returns a normalized datetime.datetime or None if parsing fails.
    """
    if value is None:
        return None

    # Already a datetime object — return directly
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)

    # Excel serial number (float/int stored as date)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            dt = (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(value))).to_pydatetime()
            return dt
        except Exception:
            return None

    # Text string — parse based on format_label
    val_str = str(value).strip()
    if not val_str:
        return None

    # Strip trailing timezone indicators for parsing.
    # IMPORTANT: only strip a +HH:MM / -HHMM offset when the value actually carries
    # a time (contains ":"). Otherwise the regex eats the "-YYYY" year of a dash date
    # like "11-24-2025" -> "11-24", corrupting the parse.
    val_clean = re.sub(r"[Zz]$", "", val_str).strip()
    if ":" in val_clean:
        val_clean = re.sub(r"[+-]\d{2}:?\d{2}$", "", val_clean).strip()

    # Try strptime patterns that match the format label
    _STRPTIME_PATTERNS = {
        "YYYY-MM-DD": ["%Y-%m-%d", "%Y/%m/%d"],
        "YYYY-MM-DD HH:MM:SS": ["%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
                                  "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M",
                                  "%Y/%m/%d %H:%M:%S"],
        "M/D/YYYY": ["%m/%d/%Y", "%m/%d/%y"],
        "MM/DD/YYYY": ["%m/%d/%Y", "%m/%d/%y"],
        "MM-DD-YYYY": ["%m-%d-%Y", "%m-%d-%y"],
        "MM.DD.YYYY": ["%m.%d.%Y", "%m.%d.%y"],
        "DD-MM-YYYY": ["%d-%m-%Y", "%d-%m-%y", "%d/%m/%Y", "%d/%m/%y"],
        "DD/MM/YYYY": ["%d/%m/%Y", "%d/%m/%y"],
        "DD.MM.YYYY": ["%d.%m.%Y", "%d.%m.%y"],
        "DD-MMM-YYYY": ["%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d%b%Y", "%b %d, %Y", "%b %d %Y"],
        "DD-MMMM-YYYY": ["%d-%B-%Y", "%d %B %Y", "%B %d, %Y"],
    }
    patterns = _STRPTIME_PATTERNS.get(format_label, [])
    for pattern in patterns:
        try:
            return datetime.datetime.strptime(val_clean, pattern)
        except ValueError:
            continue

    # Fallback: use pd.to_datetime with format-aware dayfirst
    dayfirst = format_label in _DAY_FIRST_LABELS
    try:
        result = pd.to_datetime(val_clean, dayfirst=dayfirst, errors="coerce")
        if pd.notna(result):
            return result.to_pydatetime()
    except Exception:
        pass

    return None


# Mapping from format label to strptime INPUT patterns for vectorized parsing
# Each entry can be a single format string or a list (tried in order for fallback)
_VECTORIZED_PARSE_FORMATS = {
    "YYYY-MM-DD": "%Y-%m-%d",
    "YYYY-MM-DD HH:MM:SS": "%Y-%m-%d %H:%M:%S",
    "M/D/YYYY": "%m/%d/%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "MM-DD-YYYY": "%m-%d-%Y",
    "MM.DD.YYYY": "%m.%d.%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "DD/MM/YYYY": "%d/%m/%Y",
    "DD.MM.YYYY": "%d.%m.%Y",
    "DD-MMM-YYYY": "%d-%b-%Y",
    "DD-MMMM-YYYY": "%d-%B-%Y",
    "YYYY/MM/DD": "%Y/%m/%d",
    "YYYY.MM.DD": "%Y.%m.%d",
    "M/D/YYYY HH:MM:SS": "%m/%d/%Y %H:%M:%S",
    "MM/DD/YYYY HH:MM:SS": "%m/%d/%Y %H:%M:%S",
    "MM-DD-YYYY HH:MM:SS": "%m-%d-%Y %H:%M:%S",
    "DD-MM-YYYY HH:MM:SS": "%d-%m-%Y %H:%M:%S",
    "DD/MM/YYYY HH:MM:SS": "%d/%m/%Y %H:%M:%S",
    "DD-MMM-YYYY HH:MM:SS": "%d-%b-%Y %H:%M:%S",
    "YYYY/MM/DD HH:MM:SS": "%Y/%m/%d %H:%M:%S",
    "YYYY.MM.DD HH:MM:SS": "%Y.%m.%d %H:%M:%S",
}
# Fallback patterns without seconds (for "2023-04-12T20:34Z" → "2023-04-12 20:34")
_VECTORIZED_PARSE_FORMATS_NO_SEC = {
    "YYYY-MM-DD HH:MM:SS": "%Y-%m-%d %H:%M",
    "M/D/YYYY HH:MM:SS": "%m/%d/%Y %H:%M",
    "MM/DD/YYYY HH:MM:SS": "%m/%d/%Y %H:%M",
    "MM-DD-YYYY HH:MM:SS": "%m-%d-%Y %H:%M",
    "DD-MM-YYYY HH:MM:SS": "%d-%m-%Y %H:%M",
    "DD/MM/YYYY HH:MM:SS": "%d/%m/%Y %H:%M",
    "DD-MMM-YYYY HH:MM:SS": "%d-%b-%Y %H:%M",
    "YYYY/MM/DD HH:MM:SS": "%Y/%m/%d %H:%M",
    "YYYY.MM.DD HH:MM:SS": "%Y.%m.%d %H:%M",
}


def _vectorized_date_reformat(df, col_name, src_fmt, target_py_fmt):
    """Reformat a date column FAST using vectorized pd.to_datetime.

    50-100x faster than per-row .apply() for large files.
    """
    import time as _vdr_time
    _vdr_t0 = _vdr_time.perf_counter()

    series = df[col_name]

    # Check if series contains string data (works with both object and StringDtype)
    is_str_series = pd.api.types.is_string_dtype(series)

    # For string date data, clean ISO markers FIRST for robust parsing
    # This handles "2023-04-12T20:34Z" → "2023-04-12 20:34" before format matching
    if is_str_series and src_fmt and "HH:MM:SS" in src_fmt:
        # Pre-clean: strip Z/timezone, replace T with space (vectorized, fast)
        cleaned = series.str.replace(r"[Zz]$", "", regex=True)
        cleaned = cleaned.str.replace(r"[+-]\d{2}:?\d{2}$", "", regex=True)
        cleaned = cleaned.str.replace("T", " ", regex=False).str.strip()
    else:
        cleaned = series

    _vdr_t1 = _vdr_time.perf_counter()

    # Get the strptime pattern for the source format
    parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt)
    if not parse_fmt:
        parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt.replace(" HH:MM:SS", ""))

    parsed = None
    if parse_fmt:
        # FAST PATH: vectorized parsing with known format
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            parsed = pd.to_datetime(cleaned, format=parse_fmt, errors="coerce")

        # If most values failed (no seconds in data), try without-seconds pattern
        n_valid = series.notna().sum()
        n_parsed = parsed.notna().sum() if parsed is not None else 0
        if n_valid > 0 and n_parsed < n_valid * 0.3:
            nosec_fmt = _VECTORIZED_PARSE_FORMATS_NO_SEC.get(src_fmt)
            if nosec_fmt:
                parsed2 = pd.to_datetime(cleaned, format=nosec_fmt, errors="coerce")
                if parsed2.notna().sum() > n_parsed:
                    parsed = parsed2

    if parsed is None or parsed.notna().sum() == 0:
        # FALLBACK: let pandas infer the format
        dayfirst = src_fmt in _DAY_FIRST_LABELS
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            try:
                parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=dayfirst, format="mixed")
            except TypeError:
                parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=dayfirst)

    _vdr_t2 = _vdr_time.perf_counter()

    # Reformat successfully parsed values (only non-empty source cells)
    if parsed is not None:
        mask = parsed.notna() & series.notna()
        if is_str_series:
            mask = mask & (series.str.strip() != "")
        if mask.any():
            df.loc[mask, col_name] = parsed[mask].dt.strftime(target_py_fmt)

    _vdr_t3 = _vdr_time.perf_counter()
    print(f"  _vectorized_date_reformat:")
    print(f"    column = {col_name}")
    print(f"    rows = {len(series)}")
    print(f"    non_null = {series.notna().sum()}")
    print(f"    source_format = {src_fmt}")
    print(f"    target_format = {target_py_fmt}")
    print(f"    parse_fmt_used = {parse_fmt}")
    print(f"    clean_time = {_vdr_t1 - _vdr_t0:.3f} sec")
    print(f"    parse_time = {_vdr_t2 - _vdr_t1:.3f} sec")
    print(f"    strftime_time = {_vdr_t3 - _vdr_t2:.3f} sec")
    print(f"    total = {_vdr_t3 - _vdr_t0:.3f} sec")


def _read_column_sample(file_path, sheet_name, hdr_row, col_idx, is_csv, max_sample=10):
    """Read raw cell values from a column for date comparison.

    Returns a list of raw values (str for CSV/text cells, datetime for date cells,
    float for numeric/serial cells). Skips None/empty values.
    """
    values = []

    if is_csv:
        try:
            df = pd.read_csv(file_path, nrows=max_sample + 5, dtype=str, low_memory=False)
            if col_idx < len(df.columns):
                col_name = df.columns[col_idx]
                for val in df[col_name].head(max_sample):
                    if pd.notna(val) and str(val).strip():
                        values.append(val)
        except Exception:
            pass
    else:
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
            ws = wb[sheet_name]
            count = 0
            for row in ws.iter_rows(
                min_row=hdr_row + 1,
                max_row=hdr_row + max_sample + 5,
                min_col=col_idx + 1,
                max_col=col_idx + 1
            ):
                cell = row[0]
                if cell.value is not None and str(cell.value).strip():
                    values.append(cell.value)
                    count += 1
                    if count >= max_sample:
                        break
            wb.close()
        except Exception:
            pass

    return values


def _classify_date_cell_storage(file_path, sheet_name, hdr_row, col_idx, is_csv, max_sample=15):
    """Determine how dates are physically stored in a column.

    Returns:
        "native_date" — cells contain datetime objects or numeric serials with date format
        "text_date"   — cells contain string representations of dates
        "mixed"       — both storage types present
        "unknown"     — could not determine
    """
    if is_csv:
        return "text_date"  # CSV always stores dates as text

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
        ws = wb[sheet_name]

        native_count = 0
        text_count = 0

        for row in ws.iter_rows(
            min_row=hdr_row + 1,
            max_row=hdr_row + max_sample,
            min_col=col_idx + 1,
            max_col=col_idx + 1
        ):
            cell = row[0]
            if cell.value is None:
                continue
            # Check Python type FIRST — most reliable indicator
            if isinstance(cell.value, (datetime.datetime, datetime.date)):
                native_count += 1
            elif cell.data_type == "d":
                native_count += 1
            elif cell.data_type == "n":
                nfmt = cell.number_format or ""
                if any(x in nfmt.lower() for x in ["yy", "mm", "dd", "m/d", "d-", "d/", "mmm"]):
                    native_count += 1
                # else: plain number, not counted (correctly not a date cell)
            elif cell.data_type == "s" or isinstance(cell.value, str):
                if _detect_date_text_format(str(cell.value).strip()):
                    text_count += 1

        wb.close()

        if native_count > 0 and text_count > 0:
            return "mixed"
        elif native_count > 0:
            return "native_date"
        elif text_count > 0:
            return "text_date"
        else:
            return "unknown"
    except Exception:
        return "unknown"


def _compare_date_columns(placed_path, template_path,
                          placed_sheet, template_sheet,
                          placed_hdr_row, tpl_hdr_row,
                          placed_col_idx, tpl_col_idx,
                          p_fmt, t_fmt,
                          is_placed_csv, is_template_csv,
                          max_sample=10):
    """Structurally compare date columns between placed and template.

    Uses CANONICAL SIGNATURES for the primary decision — does NOT compare
    actual row-by-row date values (the template holds different data from the
    placed file, making value comparison meaningless for structural decisions).

    Returns a dict with:
        classification: "no_mismatch" | "format_mismatch" | "data_type"
        details: Human-readable explanation
        placed_storage: "native_date" | "text_date" | "mixed" | "unknown"
        template_storage: same
        placed_signature: (date_part, has_time)
        template_signature: (date_part, has_time)
        mismatch_reason: Explanation of why signatures differ (if they do)
    """
    result = {
        "classification": "no_mismatch",
        "details": "",
        "placed_storage": "unknown",
        "template_storage": "unknown",
        "placed_signature": (None, False),
        "template_signature": (None, False),
        "mismatch_reason": "",
    }

    # Step 1: Derive canonical signatures
    p_sig = _canonical_date_signature(p_fmt)
    t_sig = _canonical_date_signature(t_fmt)
    result["placed_signature"] = p_sig
    result["template_signature"] = t_sig

    # Step 2: Classify physical storage types
    p_storage = _classify_date_cell_storage(
        placed_path, placed_sheet, placed_hdr_row, placed_col_idx, is_placed_csv)
    t_storage = _classify_date_cell_storage(
        template_path, template_sheet, tpl_hdr_row, tpl_col_idx, is_template_csv)
    result["placed_storage"] = p_storage
    result["template_storage"] = t_storage

    # Step 3: Compare canonical signatures (structural format comparison)
    if _date_signatures_match(p_sig, t_sig):
        result["classification"] = "no_mismatch"
        result["details"] = "Date format signatures are identical."
        result["mismatch_reason"] = "Date format signatures are identical."
        return result

    # Signatures differ — determine the nature of the mismatch
    reason = _describe_signature_mismatch(p_sig, t_sig)
    result["classification"] = "format_mismatch"
    result["details"] = f"Format mismatch: {p_fmt} vs {t_fmt}"
    result["mismatch_reason"] = reason

    return result


def _detect_types_uncached(file_path, sheet_name, hdr_row, max_sample=15):
    """Run the same detection logic as detect_column_types without LRU cache."""
    types = {}
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False, data_only=True)
        ws = wb[sheet_name]
        headers = []
        samples = {}

        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=hdr_row + max_sample, max_col=100), 1):
            if row_idx == hdr_row:
                headers = [(cell.value, i) for i, cell in enumerate(row) if cell.value is not None and str(cell.value).strip()]
            elif row_idx > hdr_row:
                for i, cell in enumerate(row):
                    col_i = i
                    if cell.value is not None and str(cell.value).strip():
                        if col_i not in samples:
                            samples[col_i] = []
                        samples[col_i].append({
                            "value": cell.value,
                            "format": cell.number_format,
                            "dtype": cell.data_type,
                        })

        wb.close()

        for hdr_val, col_i in headers:
            if hdr_val is None or not str(hdr_val).strip():
                continue
            col_samples = samples.get(col_i, [])
            if not col_samples:
                types[col_i] = {"header": str(hdr_val).strip(), "type": "unknown", "format": "General"}
                continue

            date_count = num_count = text_count = 0
            fmt = "General"
            detected_formats = []

            for s in col_samples:
                val = s["value"]
                nfmt = s["format"]
                cell_dtype = s["dtype"]

                if cell_dtype == "d":
                    date_count += 1
                    if nfmt and nfmt != "General":
                        detected_formats.append(nfmt)
                    else:
                        detected_formats.append("YYYY-MM-DD")
                elif cell_dtype == "n":
                    if nfmt and any(x in nfmt.lower() for x in ["yy", "mm", "dd", "m/d", "d-", "d/", "mmm"]):
                        date_count += 1
                        detected_formats.append(nfmt)
                    else:
                        num_count += 1
                elif cell_dtype == "s" or isinstance(val, str):
                    val_str = str(val).strip()
                    date_text_fmt = _detect_date_text_format(val_str)

                    if date_text_fmt:
                        date_count += 1
                        detected_formats.append(date_text_fmt)
                    else:
                        text_count += 1
                else:
                    if isinstance(val, datetime.datetime):
                        date_count += 1
                        detected_formats.append("YYYY-MM-DD")
                    elif isinstance(val, (int, float)):
                        num_count += 1
                    else:
                        text_count += 1

            total = date_count + num_count + text_count
            if total == 0:
                dtype_result = "unknown"
            elif date_count > 0 and date_count >= num_count and date_count >= text_count:
                dtype_result = "date"
                if detected_formats:
                    from collections import Counter
                    fmt = Counter(detected_formats).most_common(1)[0][0]
                else:
                    fmt = "General"
            elif num_count > 0 and num_count > text_count:
                dtype_result = "numeric"
                fmt = "General"
            else:
                dtype_result = "text"
                fmt = "General"

            types[col_i] = {"header": str(hdr_val).strip(), "type": dtype_result, "format": fmt}
    except Exception:
        pass
    return types


def compare_file_pair(placed_path, template_path, placed_name, template_name):
    """Compare a placed file against its template."""
    findings = []

    # Check 0: File type/extension mismatch (.xlsx vs .xls vs .csv)
    placed_ext = Path(placed_name).suffix.lower()
    template_ext = Path(template_name).suffix.lower()
    if placed_ext != template_ext:
        findings.append({
            "File": placed_name, "Check": "File Type Mismatch",
            "Placed Value": placed_ext, "Template Value": template_ext,
            "fix_type": "convert_file_type",
            "fix_data": {"from_ext": placed_ext, "to_ext": template_ext},
        })

    # Check 1: Filename mismatch (skip if only difference is the extension - already caught by Check 0)
    placed_stem = Path(placed_name).stem
    template_stem = Path(template_name).stem
    if placed_name != template_name:
        # Only flag if stems differ (not just extension difference)
        if placed_ext != template_ext and placed_stem.lower() == template_stem.lower():
            pass  # Extension-only difference, already flagged as File Type Mismatch
        else:
            findings.append({
                "File": placed_name, "Check": "Filename Mismatch",
                "Placed Value": placed_name, "Template Value": template_name,
                "fix_type": "rename_file", "fix_data": {"old_name": placed_name, "new_name": template_name},
            })

    # Determine if files are CSV or Excel for header/type detection
    is_placed_csv = placed_ext == ".csv"
    is_template_csv = template_ext == ".csv"

    # Wrap deeper analysis in try/except so early findings (filename/extension)
    # are always returned even if file I/O fails.
    try:
        _deeper_analysis(findings, placed_path, template_path, placed_name,
                         is_placed_csv, is_template_csv, placed_ext)
    except Exception as e:
        print(f"[compare] _deeper_analysis failed for {placed_name}: {type(e).__name__}: {e}")

    return findings


# --- Module-level column normalization and alias data (avoid recreating per call) ---
_RE_COL_SEPARATORS = re.compile(r"[>|:;]+")
_RE_COL_WHITESPACE = re.compile(r"\s+")


def _norm_col(c):
    """Normalize column name for comparison purposes.

    Normalizes by: trimming whitespace, collapsing multiple spaces,
    replacing underscores/hyphens with spaces, removing separator punctuation
    (>, |, :, ;), normalizing Unicode, and removing non-breaking spaces.
    """
    n = str(c).strip()
    n = unicodedata.normalize("NFKC", n)
    n = n.replace(" ", " ")
    n = n.replace("_", " ").replace("-", " ")
    n = _RE_COL_SEPARATORS.sub(" ", n)
    n = _RE_COL_WHITESPACE.sub(" ", n).strip()
    return n


_COL_ALIASES = {
    "subject": {"subject", "subject id", "subjid", "subj id", "subj_id", "subject number", "patient number", "patient"},
    "status": {"status", "issue status", "issue_status", "patient status", "form status"},
    "site": {"site", "site number", "site id", "siteid", "site_id"},
    "visit": {"visit", "visit number", "visit id", "visitid"},
    "investigator": {"investigator", "inv", "inv name", "investigator name"},
}

_alias_reverse = {}
for _canonical, _aliases in _COL_ALIASES.items():
    for _alias in _aliases:
        _alias_reverse[_alias] = _canonical


def _deeper_analysis(findings, placed_path, template_path, placed_name,
                     is_placed_csv, is_template_csv, placed_ext):
    placed_sheet = ""
    template_sheet = ""
    if not is_placed_csv:
        placed_sheet = get_data_sheet(placed_path)
    if not is_template_csv:
        template_sheet = get_data_sheet(template_path)

    # Enrich convert_file_type finding with template sheet name (for CSV→XLSX naming)
    if template_sheet:
        for f in findings:
            if f.get("fix_type") == "convert_file_type" and f.get("fix_data"):
                f["fix_data"]["template_sheet"] = template_sheet
                break

    # Check 2: Sheet name (only applicable if both are Excel)
    if placed_sheet and template_sheet and not is_placed_csv and not is_template_csv:
        if placed_sheet != template_sheet:
            findings.append({
                "File": placed_name, "Check": "Sheet Name",
                "Placed Value": repr(placed_sheet), "Template Value": repr(template_sheet),
                "fix_type": "rename_sheet", "fix_data": {"sheet_old": placed_sheet, "sheet_new": template_sheet},
            })
        elif placed_sheet != placed_sheet.strip() and template_sheet == template_sheet.strip():
            # Only flag spacing issue if template is clean but placed has extra spaces
            findings.append({
                "File": placed_name, "Check": "Sheet Name (spaces)",
                "Placed Value": repr(placed_sheet), "Template Value": repr(template_sheet),
                "fix_type": "rename_sheet", "fix_data": {"sheet_old": placed_sheet, "sheet_new": template_sheet},
            })

    # Get headers (full row including None)
    if is_placed_csv:
        placed_hdr_row, placed_headers = _get_csv_headers(placed_path)
    else:
        placed_hdr_row, placed_headers = get_header_columns(placed_path, placed_sheet) if placed_sheet else (1, [])

    if is_template_csv:
        tpl_hdr_row, tpl_headers = _get_csv_headers(template_path)
    else:
        tpl_hdr_row, tpl_headers = get_header_columns(template_path, template_sheet) if template_sheet else (1, [])

    placed_cols = get_non_empty_columns(placed_headers)
    tpl_cols = get_non_empty_columns(tpl_headers)

    # === COLUMN MAPPING PIPELINE ===
    # Build a complete mapping of placed columns → template columns BEFORE generating findings.
    # A column can only have ONE status: Matched, Mismatch, Extra, or Missing.

    # Normalize all columns for matching
    placed_norm = [_norm_col(c) for c in placed_cols]
    tpl_norm = [_norm_col(c) for c in tpl_cols]

    # --- Phase 1: Exact match (normalized names are identical, case-insensitive) ---
    matched_pairs = {}  # {placed_col_original: template_col_original}
    unmatched_placed_indices = set(range(len(placed_cols)))
    unmatched_tpl_indices = set(range(len(tpl_cols)))

    # Build a lookup from normalized lowercase template name → index for O(1) matching
    tpl_norm_lookup = {}
    for ti in range(len(tpl_cols)):
        key = tpl_norm[ti].lower()
        if key not in tpl_norm_lookup:  # first occurrence wins
            tpl_norm_lookup[key] = ti

    for pi in list(unmatched_placed_indices):
        key = placed_norm[pi].lower()
        ti = tpl_norm_lookup.get(key)
        if ti is not None and ti in unmatched_tpl_indices:
            matched_pairs[placed_cols[pi]] = tpl_cols[ti]
            unmatched_placed_indices.discard(pi)
            unmatched_tpl_indices.discard(ti)

    # --- Phase 2: Alias match (both map to the same semantic group) ---
    # Build template alias group lookup for O(1) matching
    tpl_alias_groups = {}  # group_key -> ti
    for ti in unmatched_tpl_indices:
        t_group = _alias_reverse.get(tpl_norm[ti].lower())
        if t_group and t_group not in tpl_alias_groups:
            tpl_alias_groups[t_group] = ti

    for pi in list(unmatched_placed_indices):
        p_group = _alias_reverse.get(placed_norm[pi].lower())
        if not p_group:
            continue
        ti = tpl_alias_groups.get(p_group)
        if ti is not None and ti in unmatched_tpl_indices:
            matched_pairs[placed_cols[pi]] = tpl_cols[ti]
            unmatched_placed_indices.discard(pi)
            unmatched_tpl_indices.discard(ti)

    # --- Phase 3: Fuzzy/semantic match (one contains the other as substring, case-insensitive) ---
    for pi in list(unmatched_placed_indices):
        for ti in list(unmatched_tpl_indices):
            pn = placed_norm[pi].lower()
            tn = tpl_norm[ti].lower()
            # Substring containment (significant overlap)
            if len(tn) >= 4 and tn in pn:
                matched_pairs[placed_cols[pi]] = tpl_cols[ti]
                unmatched_placed_indices.discard(pi)
                unmatched_tpl_indices.discard(ti)
                break
            elif len(pn) >= 4 and pn in tn:
                matched_pairs[placed_cols[pi]] = tpl_cols[ti]
                unmatched_placed_indices.discard(pi)
                unmatched_tpl_indices.discard(ti)
                break

    # === FREEZE MAPPING — no more matches after this point ===

    # Safety net: ensure subject/status columns identified by keyword search
    # are in the mapping even if normalization didn't catch them.
    # A column cannot simultaneously be Matched, Extra, and Missing.
    placed_subj = find_subject_column(placed_cols)
    template_subj = find_subject_column(tpl_cols)
    if placed_subj and template_subj and placed_subj not in matched_pairs:
        matched_pairs[placed_subj] = template_subj
        pi_subj = placed_cols.index(placed_subj)
        ti_subj = tpl_cols.index(template_subj)
        unmatched_placed_indices.discard(pi_subj)
        unmatched_tpl_indices.discard(ti_subj)

    placed_status = find_status_column(placed_cols)
    template_status = find_status_column(tpl_cols)
    if placed_status and template_status and placed_status not in matched_pairs:
        matched_pairs[placed_status] = template_status
        pi_stat = placed_cols.index(placed_status)
        ti_stat = tpl_cols.index(template_status)
        unmatched_placed_indices.discard(pi_stat)
        unmatched_tpl_indices.discard(ti_stat)

    # Compute extra/missing AFTER all matching is finalized
    # Extra columns = placed columns that have NO match in template
    extra_cols = [placed_cols[i] for i in unmatched_placed_indices]
    # Missing columns = template columns that have NO match in placed
    missing_cols = [tpl_cols[i] for i in unmatched_tpl_indices]

    # === GENERATE FINDINGS FROM MAPPING ===

    # Check 3: Subject column mismatch
    if placed_subj and template_subj:
        if placed_subj != template_subj:
            findings.append({
                "File": placed_name, "Check": "Subject Column",
                "Placed Value": repr(placed_subj), "Template Value": repr(template_subj),
                "fix_type": "rename_column", "fix_data": {"sheet": placed_sheet, "old_col": placed_subj, "new_col": template_subj},
            })

    # Check 4: Status column mismatch
    if placed_status and template_status:
        if placed_status != template_status:
            findings.append({
                "File": placed_name, "Check": "Status Column",
                "Placed Value": repr(placed_status), "Template Value": repr(template_status),
                "fix_type": "rename_column", "fix_data": {"sheet": placed_sheet, "old_col": placed_status, "new_col": template_status},
            })

    # Check 5: Extra columns (only truly unmatched columns)
    if extra_cols:
        findings.append({
            "File": placed_name, "Check": "Extra Columns",
            "Placed Value": f"{len(placed_cols)} cols (extra: {', '.join(extra_cols[:3])}{'...' if len(extra_cols) > 3 else ''})",
            "Template Value": f"{len(tpl_cols)} columns",
            "fix_type": "remove_extra_columns",
            "fix_data": {"sheet": placed_sheet, "hdr_row": placed_hdr_row,
                         "keep_count": len(tpl_cols), "total_count": len(placed_cols),
                         "extra_col_names": extra_cols,
                         "is_csv": is_placed_csv},
        })
    if missing_cols and not extra_cols:
        findings.append({
            "File": placed_name, "Check": "Fewer Columns",
            "Placed Value": f"{len(placed_cols)} columns ({len(missing_cols)} fewer)",
            "Template Value": f"{len(tpl_cols)} columns",
            "fix_type": None, "fix_data": None,
        })

    # Check 6: Data type mismatches (works for both CSV and Excel)
    # Uses matched_pairs mapping to compare corresponding columns by their position/index
    # Note: detect_column_types cache is cleared from the main thread before the
    # ThreadPoolExecutor starts (see _run_compare block), so no need to clear here.

    if is_placed_csv:
        placed_types = _detect_csv_column_types(placed_path)
    elif placed_sheet:
        placed_types = detect_column_types(placed_path, placed_sheet, placed_hdr_row)
    else:
        placed_types = {}

    if is_template_csv:
        tpl_types = _detect_csv_column_types(template_path)
    elif template_sheet:
        tpl_types = detect_column_types(template_path, template_sheet, tpl_hdr_row)
    else:
        tpl_types = {}

    if placed_types and tpl_types:
        # Build template type lookup by BOTH exact name and normalized name
        tpl_type_by_name = {}
        tpl_col_idx_by_name = {}  # header_lower → column index in template
        for tpl_col_i, info in tpl_types.items():
            hdr = info["header"].strip().lower()
            tpl_type_by_name[hdr] = info
            tpl_col_idx_by_name[hdr] = tpl_col_i
            # Also index by normalized form (strip punctuation, replace separators)
            hdr_norm = re.sub(r"[?!.]+$", "", re.sub(r"[_\-]+", " ", hdr)).strip()
            tpl_type_by_name[hdr_norm] = info
            tpl_col_idx_by_name[hdr_norm] = tpl_col_i

        # Build a reverse lookup from placed column name -> matched template column name
        placed_to_tpl_name = {}
        for p_col, t_col in matched_pairs.items():
            placed_to_tpl_name[p_col.strip().lower()] = t_col.strip().lower()

        for col_i, p_info in placed_types.items():
          try:
            p_name = p_info["header"].strip().lower()

            # Find the corresponding template column via matched_pairs first
            t_name = placed_to_tpl_name.get(p_name, p_name)
            t_info = tpl_type_by_name.get(t_name)

            # Fallback: try normalized form
            if not t_info:
                p_name_norm = re.sub(r"[?!.]+$", "", re.sub(r"[_\-]+", " ", p_name)).strip()
                t_name_norm = re.sub(r"[?!.]+$", "", re.sub(r"[_\-]+", " ", t_name)).strip()
                t_info = tpl_type_by_name.get(t_name_norm) or tpl_type_by_name.get(p_name_norm)

            if not t_info:
                continue

            p_type, t_type = p_info["type"], t_info["type"]

            if p_type == "unknown" or t_type == "unknown":
                continue

            if p_type != t_type:
                findings.append({
                    "File": placed_name, "Check": f"Data Type ({p_info['header']})",
                    "Placed Value": f"{p_type}", "Template Value": f"{t_type}",
                    "fix_type": "fix_data_type",
                    "fix_data": {"sheet": placed_sheet, "col_idx": col_i, "hdr_row": placed_hdr_row,
                                 "from_type": p_type, "to_type": t_type, "target_format": t_info["format"],
                                 "source_format": p_info["format"],
                                 "template_path": template_path, "tpl_col_idx": tpl_col_idx_by_name.get(t_name, col_i),
                                 "tpl_hdr_row": tpl_hdr_row,
                                 "is_csv": is_placed_csv},
                })
            elif p_type == "date" and t_type == "date":
                # Date format comparison using CANONICAL SIGNATURES.
                # A date column is a real mismatch ONLY when the structural
                # FORMAT differs (component order + delimiter). Actual cell
                # VALUES are never compared: the template is a structural
                # reference and holds different data, so row-by-row comparison
                # would always produce false positives.
                p_fmt = p_info["format"]
                t_fmt = t_info["format"]

                # Skip legacy/unknown format labels
                if p_fmt == "text-date" or t_fmt == "text-date":
                    continue

                tpl_col_actual = tpl_col_idx_by_name.get(t_name, col_i)

                # --- Canonical signature comparison ---
                # Derives structural fingerprint: component order + delimiter.
                # Time precision (HH:MM vs HH:MM:SS, or time present vs absent)
                # is NOT treated as a structural mismatch.
                p_sig = _canonical_date_signature(p_fmt)
                t_sig = _canonical_date_signature(t_fmt)

                if _date_signatures_match(p_sig, t_sig):
                    continue  # Same structural format → no finding

                # Real format mismatch found — flag it
                check_label = f"Date Format ({p_info['header']})"
                p_storage = "text_date" if is_placed_csv else "native_date"
                t_storage = "text_date" if is_template_csv else "native_date"

                findings.append({
                    "File": placed_name, "Check": check_label,
                    "Placed Value": p_fmt, "Template Value": t_fmt,
                    "fix_type": "fix_data_type",
                    "fix_data": {"sheet": placed_sheet, "col_idx": col_i, "hdr_row": placed_hdr_row,
                                 "from_type": "date", "to_type": "date", "target_format": t_fmt,
                                 "source_format": p_fmt,
                                 "template_path": template_path, "tpl_col_idx": tpl_col_actual,
                                 "tpl_hdr_row": tpl_hdr_row,
                                 "is_csv": is_placed_csv,
                                 "date_classification": "format_mismatch",
                                 "placed_storage": p_storage,
                                 "template_storage": t_storage},
                })
          except Exception as e:
            print(f"[compare] Check 6 column error ({p_info.get('header', '?')}): {type(e).__name__}: {e}")
            continue  # Skip this column on error, process remaining columns

        # Enrich convert_file_type finding with matched template column types.
        # This allows the fast path to serialize columns using the TEMPLATE type
        # (not just the CSV-detected type) — satisfying the requirement that
        # serialization decisions are based on both CSV AND template agreement.
        if is_placed_csv:
            _matched_col_types = {}  # {csv_col_idx: template_type_str}
            for col_i, p_info in placed_types.items():
                p_name = p_info["header"].strip().lower()
                t_name = placed_to_tpl_name.get(p_name, p_name)
                t_info = tpl_type_by_name.get(t_name)
                if not t_info:
                    p_name_norm = re.sub(r"[?!.]+$", "", re.sub(r"[_\-]+", " ", p_name)).strip()
                    t_name_norm = re.sub(r"[?!.]+$", "", re.sub(r"[_\-]+", " ", t_name)).strip()
                    t_info = tpl_type_by_name.get(t_name_norm) or tpl_type_by_name.get(p_name_norm)
                if t_info:
                    _matched_col_types[col_i] = t_info["type"]

            if _matched_col_types:
                for f in findings:
                    if f.get("fix_type") == "convert_file_type" and f.get("fix_data"):
                        f["fix_data"]["matched_col_types"] = _matched_col_types
                        break

    # Check 7: Column header formatting — for matched columns, check if exact names differ
    # Use the frozen mapping to identify columns that matched semantically but differ in formatting
    for placed_col, tpl_col in matched_pairs.items():
        # Skip subject/status (already reported in checks 3 & 4)
        if placed_col == placed_subj or placed_col == placed_status:
            continue
        # If the exact strings differ, it's a formatting/spacing issue
        if placed_col != tpl_col:
            findings.append({
                "File": placed_name, "Check": "Column Spaces",
                "Placed Value": repr(placed_col), "Template Value": repr(tpl_col),
                "fix_type": "rename_column", "fix_data": {"sheet": placed_sheet, "old_col": placed_col, "new_col": tpl_col,
                                                            "is_csv": is_placed_csv},
            })

    return findings


def _excel_convert(src_path, dst_path, sheet_name=None):
    """Convert between Excel/CSV formats using Excel COM automation (instant).

    Uses PowerShell to invoke Excel COM — works on any Windows machine with Excel installed.
    Supports: CSV↔XLSX, XLS↔XLSX, CSV↔XLS, and any combination.
    Optionally renames the first sheet to `sheet_name` after conversion.
    Returns True on success, False on failure.
    """
    import subprocess

    src_abs = os.path.abspath(src_path).replace("'", "''")
    dst_abs = os.path.abspath(dst_path).replace("'", "''")
    dst_ext = Path(dst_path).suffix.lower()

    # Excel file format constants
    fmt_map = {".xlsx": 51, ".xls": 56, ".csv": 6}
    file_format = fmt_map.get(dst_ext, 51)

    # Build PowerShell script
    rename_cmd = ""
    if sheet_name and dst_ext in (".xlsx", ".xls"):
        safe_sheet = sheet_name.replace("'", "''")
        rename_cmd = f"$wb.Sheets.Item(1).Name = '{safe_sheet}'"

    ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Open('{src_abs}')
    {rename_cmd}
    $wb.SaveAs('{dst_abs}', {file_format})
    $wb.Close($false)
    Write-Output 'OK'
}} catch {{
    Write-Output "FAIL:$($_.Exception.Message)"
}} finally {{
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=120
        )
        return result.stdout.strip().startswith("OK")
    except Exception:
        return False


def _excel_convert_batch(conversions):
    """Batch-convert multiple files using a SINGLE Excel COM instance.

    Args:
        conversions: list of (src_path, dst_path, sheet_name_or_None) tuples

    Returns:
        list of booleans (True/False for each conversion success)
    """
    import subprocess

    if not conversions:
        return []

    # Build a single PowerShell script that opens Excel once and converts all files
    file_ops = []
    for i, (src, dst, sheet) in enumerate(conversions):
        src_abs = os.path.abspath(src).replace("'", "''")
        dst_abs = os.path.abspath(dst).replace("'", "''")
        dst_ext = Path(dst).suffix.lower()
        fmt_map = {".xlsx": 51, ".xls": 56, ".csv": 6}
        file_format = fmt_map.get(dst_ext, 51)

        rename_cmd = ""
        if sheet and dst_ext in (".xlsx", ".xls"):
            safe_sheet = sheet.replace("'", "''")
            rename_cmd = f"        $wb.Sheets.Item(1).Name = '{safe_sheet}'"

        file_ops.append(f"""
    try {{
        $wb = $excel.Workbooks.Open('{src_abs}')
{rename_cmd}
        $wb.SaveAs('{dst_abs}', {file_format})
        $wb.Close($false)
        Write-Output 'OK:{i}'
    }} catch {{
        Write-Output 'FAIL:{i}'
    }}""")

    ops_block = "\n".join(file_ops)
    ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
{ops_block}
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=300
        )
        output = result.stdout.strip()
        # Parse results
        results = [False] * len(conversions)
        for line in output.split("\n"):
            line = line.strip()
            if line.startswith("OK:"):
                idx = int(line.split(":")[1])
                results[idx] = True
        return results
    except Exception:
        return [False] * len(conversions)


def _excel_fix_workbook(file_path, operations):
    """Apply workbook operations via Excel COM. Returns list of booleans (per-op success)."""
    import subprocess

    file_abs = os.path.abspath(file_path).replace("'", "''")

    # Build PowerShell commands for each operation with per-op success tracking
    op_cmds = []
    for i, op_tuple in enumerate(operations):
        op_type = op_tuple[0]
        arg1 = op_tuple[1]
        arg2 = op_tuple[2]
        # Optional 5th element (sheet name for rename_column)
        extra = op_tuple[4] if len(op_tuple) > 4 else ""
        if op_type == "rename_sheet":
            old_s = str(arg1).replace("'", "''")
            new_s = str(arg2).replace("'", "''")
            op_cmds.append(f"""
    try {{
        $s = $wb.Sheets.Item('{old_s}')
        $s.Name = '{new_s}'
        Write-Output 'OP_OK:{i}'
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
        elif op_type == "rename_column":
            old_col = str(arg1).replace("'", "''")
            new_col = str(arg2).replace("'", "''")
            sheet_ref = f"$wb.Sheets.Item('{extra.replace(chr(39), chr(39)+chr(39))}')" if extra else "$wb.Sheets.Item(1)"
            op_cmds.append(f"""
    try {{
        $ws = {sheet_ref}
        $lastCol = $ws.UsedRange.Columns.Count + $ws.UsedRange.Column - 1
        $renamed = $false
        for ($r = 1; $r -le 30; $r++) {{
            for ($c = 1; $c -le $lastCol; $c++) {{
                $cv = $ws.Cells.Item($r, $c).Value2
                if ($cv -ne $null) {{
                    $cvStr = $cv.ToString().Trim()
                    if ($cvStr -eq '{old_col}' -or $cvStr -ieq '{old_col}') {{
                        $ws.Cells.Item($r, $c).Value2 = '{new_col}'
                        $renamed = $true
                        break
                    }}
                }}
            }}
            if ($renamed) {{ break }}
        }}
        if ($renamed) {{ Write-Output 'OP_OK:{i}' }} else {{ Write-Output 'OP_FAIL:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
        elif op_type == "remove_extra_columns":
            extra_names = arg1  # list of column names
            if extra_names:
                names_arr = ",".join([f"'{n.replace(chr(39), chr(39)+chr(39))}'" for n in extra_names])
                op_cmds.append(f"""
    try {{
        $ws = $wb.Sheets.Item(1)
        $extraNames = @({names_arr})
        $lastCol = $ws.UsedRange.Columns.Count + $ws.UsedRange.Column - 1
        $deleted = 0
        for ($c = $lastCol; $c -ge 1; $c--) {{
            $hdr = $ws.Cells.Item(1, $c).Value2
            if ($hdr -ne $null -and ($extraNames -contains $hdr.ToString().Trim())) {{
                $ws.Columns.Item($c).Delete() | Out-Null
                $deleted++
            }}
        }}
        if ($deleted -gt 0) {{ Write-Output 'OP_OK:{i}' }} else {{ Write-Output 'OP_FAIL:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
            else:
                op_cmds.append(f"    Write-Output 'OP_FAIL:{i}'")
        elif op_type == "fix_data_type":
            fix_data = arg1
            col_idx = fix_data.get("col_idx", 0)
            hdr_row = fix_data.get("hdr_row", 1)
            target_fmt = fix_data.get("target_format", "")
            to_type = fix_data.get("to_type", "")
            from_type = fix_data.get("from_type", "")
            sheet_name = fix_data.get("sheet", "")
            src_fmt = fix_data.get("source_format", "")
            if to_type == "date" and target_fmt and target_fmt not in ("General", "text-date"):
                excel_col = col_idx + 1
                safe_fmt = target_fmt.replace("'", "''")
                sheet_ref = f"$wb.Sheets.Item('{sheet_name.replace(chr(39), chr(39)+chr(39))}')" if sheet_name else "$wb.Sheets.Item(1)"
                # Cells are text if from_type is text, or source format is pure text (@)
                _cells_are_text = (from_type == "text") or (src_fmt == "@")
                if _cells_are_text:
                    # Text→Date: must convert text values to actual date serial numbers, then apply format
                    # Strategy: set NumberFormat first, then use TextToColumns to force re-parse, OR
                    # cell-by-cell conversion using Value (not Value2) for proper date handling
                    op_cmds.append(f"""
    try {{
        $ws = {sheet_ref}
        $lastRow = $ws.Cells.Item($ws.Rows.Count, {excel_col}).End(-4162).Row
        $startRow = {hdr_row} + 1
        $converted = 0
        if ($lastRow -ge $startRow) {{
            $rng = $ws.Range($ws.Cells.Item($startRow, {excel_col}), $ws.Cells.Item($lastRow, {excel_col}))
            # First try TextToColumns — this is Excel's native text-to-date conversion
            try {{
                $rng.TextToColumns($rng, 1, 1, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false)
            }} catch {{ }}
            # Now set the NumberFormat
            $rng.NumberFormat = '{safe_fmt}'
            # Verify at least some cells converted by checking if first data cell is numeric (date serial)
            $testVal = $ws.Cells.Item($startRow, {excel_col}).Value2
            if ($testVal -is [double] -or $testVal -is [int]) {{
                $converted = 1
            }} else {{
                # Fallback: cell-by-cell DateTime parse
                for ($r = $startRow; $r -le [Math]::Min($lastRow, $startRow + 999); $r++) {{
                    $cell = $ws.Cells.Item($r, {excel_col})
                    $val = $cell.Value2
                    if ($val -ne $null -and $val.ToString().Trim() -ne '') {{
                        if ($val -is [double]) {{ $converted++; continue }}
                        try {{
                            $dateVal = [DateTime]::Parse($val.ToString())
                            $cell.Value = $dateVal
                            $converted++
                        }} catch {{ }}
                    }}
                }}
                # Re-apply NumberFormat after conversion
                $rng.NumberFormat = '{safe_fmt}'
            }}
            if ($converted -gt 0) {{ Write-Output 'OP_OK:{i}' }} else {{ Write-Output 'OP_FAIL:{i}' }}
        }} else {{ Write-Output 'OP_FAIL:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
                else:
                    # Date→Date format change: cells already have date serial numbers, just change NumberFormat
                    op_cmds.append(f"""
    try {{
        $ws = {sheet_ref}
        $lastRow = $ws.Cells.Item($ws.Rows.Count, {excel_col}).End(-4162).Row
        $startRow = {hdr_row} + 1
        if ($lastRow -ge $startRow) {{
            $rng = $ws.Range($ws.Cells.Item($startRow, {excel_col}), $ws.Cells.Item($lastRow, {excel_col}))
            $rng.NumberFormat = '{safe_fmt}'
            Write-Output 'OP_OK:{i}'
        }} else {{ Write-Output 'OP_FAIL:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
            elif to_type == "text":
                # Numeric/Date → Text: bulk array approach (fast — only 2 COM calls per column)
                # 1. Read all values at once into array
                # 2. Set NumberFormat to "@"
                # 3. Convert values to strings in-memory
                # 4. Write entire array back at once
                excel_col = col_idx + 1
                sheet_ref = f"$wb.Sheets.Item('{sheet_name.replace(chr(39), chr(39)+chr(39))}')" if sheet_name else "$wb.Sheets.Item(1)"
                op_cmds.append(f"""
    try {{
        $ws = {sheet_ref}
        $lastRow = $ws.Cells.Item($ws.Rows.Count, {excel_col}).End(-4162).Row
        $startRow = {hdr_row} + 1
        if ($lastRow -ge $startRow) {{
            $rng = $ws.Range($ws.Cells.Item($startRow, {excel_col}), $ws.Cells.Item($lastRow, {excel_col}))
            $values = $rng.Value2
            $rng.NumberFormat = "@"
            $rowCount = $lastRow - $startRow + 1
            if ($rowCount -eq 1) {{
                if ($values -ne $null) {{ $rng.Value2 = $values.ToString() }}
            }} else {{
                for ($ri = 1; $ri -le $rowCount; $ri++) {{
                    if ($values[$ri, 1] -ne $null) {{
                        $values[$ri, 1] = $values[$ri, 1].ToString()
                    }}
                }}
                $rng.Value2 = $values
            }}
            Write-Output 'OP_OK:{i}'
        }} else {{ Write-Output 'OP_OK:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
            elif to_type == "numeric":
                # Text → Numeric: bulk array approach (fast)
                excel_col = col_idx + 1
                sheet_ref = f"$wb.Sheets.Item('{sheet_name.replace(chr(39), chr(39)+chr(39))}')" if sheet_name else "$wb.Sheets.Item(1)"
                op_cmds.append(f"""
    try {{
        $ws = {sheet_ref}
        $lastRow = $ws.Cells.Item($ws.Rows.Count, {excel_col}).End(-4162).Row
        $startRow = {hdr_row} + 1
        if ($lastRow -ge $startRow) {{
            $rng = $ws.Range($ws.Cells.Item($startRow, {excel_col}), $ws.Cells.Item($lastRow, {excel_col}))
            $values = $rng.Value2
            $rng.NumberFormat = "General"
            $rowCount = $lastRow - $startRow + 1
            if ($rowCount -eq 1) {{
                if ($values -ne $null -and $values -is [string]) {{
                    try {{ $rng.Value2 = [double]$values }} catch {{ }}
                }}
            }} else {{
                for ($ri = 1; $ri -le $rowCount; $ri++) {{
                    if ($values[$ri, 1] -ne $null -and $values[$ri, 1] -is [string]) {{
                        try {{ $values[$ri, 1] = [double]$values[$ri, 1] }} catch {{ }}
                    }}
                }}
                $rng.Value2 = $values
            }}
            Write-Output 'OP_OK:{i}'
        }} else {{ Write-Output 'OP_OK:{i}' }}
    }} catch {{ Write-Output 'OP_FAIL:{i}' }}""")
            else:
                op_cmds.append(f"    Write-Output 'OP_FAIL:{i}'")

    if not op_cmds:
        return [False] * len(operations)

    ops_block = "\n".join(op_cmds)
    ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Open('{file_abs}')
{ops_block}
    $wb.Save()
    $wb.Close($false)
}} catch {{
    Write-Output "SCRIPT_FAIL:$($_.Exception.Message)"
}} finally {{
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=180
        )
        output = result.stdout.strip()
        # Parse per-operation results
        op_results = [False] * len(operations)
        for line in output.split("\n"):
            line = line.strip()
            if line.startswith("OP_OK:"):
                idx = int(line.split(":")[1])
                if idx < len(op_results):
                    op_results[idx] = True
        return op_results
    except Exception:
        return [False] * len(operations)


def _fix_excel_data_types_openpyxl(file_path, data_type_fixes):
    """Fix data type mismatches in an Excel file using openpyxl + pandas.

    This avoids COM automation entirely — directly reads/writes cell values and types.
    Returns a list of booleans indicating per-fix success.

    PERFORMANCE: Uses BULK read → vectorized pandas transform → bulk write instead of
    per-cell iteration. For 11K rows, this is 50-100x faster than per-cell loops.

    Conversion logic:
      numeric → text:  Store values as strings with number_format "@"
      text → numeric:  Use pd.to_numeric(errors='coerce') and store as numbers
      text → date:     Use pd.to_datetime(errors='coerce') and store as datetime
      date → date:     Change number_format OR reformat as text (bulk vectorized)
      numeric → date:  Treat numeric as Excel serial, convert to datetime
    """
    results = [False] * len(data_type_fixes)

    try:
        wb = openpyxl.load_workbook(file_path, keep_links=False)
    except Exception:
        return results

    modified = False

    for fix_idx, finding in enumerate(data_type_fixes):
        fix_data = finding["fix_data"]
        sheet_name = fix_data.get("sheet", "")
        col_idx = fix_data.get("col_idx", 0)
        hdr_row = fix_data.get("hdr_row", 1)
        from_type = fix_data.get("from_type", "")
        to_type = fix_data.get("to_type", "")
        target_fmt = fix_data.get("target_format", "")

        # Get the worksheet
        try:
            if sheet_name and sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.active
        except Exception:
            continue

        # Determine data range: start after header row, use the column (1-based in openpyxl)
        excel_col = col_idx + 1  # openpyxl uses 1-based columns
        start_row = hdr_row + 1
        max_row = ws.max_row

        if max_row is None or max_row < start_row:
            results[fix_idx] = True  # No data to fix
            continue

        try:
            # --- BULK READ: get all cell values from the column in one pass ---
            cells = []
            for row_tuple in ws.iter_rows(min_row=start_row, max_row=max_row,
                                          min_col=excel_col, max_col=excel_col):
                cells.append(row_tuple[0])
            raw_values = [c.value for c in cells]

            if to_type == "text":
                # --- Numeric/Date → Text (bulk) ---
                for cell in cells:
                    if cell.value is not None:
                        val = cell.value
                        if isinstance(val, float):
                            cell.value = str(int(val)) if val == int(val) else str(val)
                        elif isinstance(val, datetime.datetime):
                            cell.value = val.strftime("%Y-%m-%d")
                        else:
                            cell.value = str(val)
                        cell.number_format = "@"
                        cell.data_type = "s"
                results[fix_idx] = True
                modified = True

            elif to_type == "numeric":
                # --- Text → Numeric (bulk via pandas) ---
                converted = pd.to_numeric(pd.Series(raw_values), errors="coerce")
                for i, cell in enumerate(cells):
                    if pd.notna(converted.iloc[i]):
                        num_val = converted.iloc[i]
                        cell.value = int(num_val) if num_val == int(num_val) else float(num_val)
                        cell.number_format = "General"
                results[fix_idx] = True
                modified = True

            elif to_type == "date" and from_type == "date":
                # --- Date → Date (format change — VECTORIZED) ---
                _tpl_storage = fix_data.get("template_storage", "")
                if _tpl_storage == "text_date":
                    is_template_text_date = True
                elif _tpl_storage in ("native_date", "mixed"):
                    is_template_text_date = False
                else:
                    is_template_text_date = target_fmt in _TEXT_DATE_LABELS

                if is_template_text_date:
                    # Template stores dates as TEXT — convert to formatted text strings
                    py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                    src_fmt = fix_data.get("source_format", "")

                    # BULK VECTORIZED: separate datetime objects from text strings
                    # and handle each group efficiently
                    str_vals = []
                    dt_indices = []   # indices where value is already datetime
                    str_indices = []  # indices where value is a text string to parse
                    for i, val in enumerate(raw_values):
                        if val is None:
                            str_vals.append("")
                        elif isinstance(val, datetime.datetime):
                            dt_indices.append(i)
                            str_vals.append("")
                        elif isinstance(val, datetime.date):
                            dt_indices.append(i)
                            str_vals.append("")
                        elif isinstance(val, (int, float)) and not isinstance(val, bool):
                            # Excel serial number
                            dt_indices.append(i)
                            str_vals.append("")
                        else:
                            str_indices.append(i)
                            str_vals.append(str(val).strip())

                    # Handle datetime objects (native dates) — direct strftime (fast)
                    for i in dt_indices:
                        val = raw_values[i]
                        if isinstance(val, (int, float)):
                            try:
                                dt = (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))).to_pydatetime()
                                cells[i].value = dt.strftime(py_fmt)
                            except Exception:
                                continue
                        elif isinstance(val, datetime.date) and not isinstance(val, datetime.datetime):
                            dt = datetime.datetime(val.year, val.month, val.day)
                            cells[i].value = dt.strftime(py_fmt)
                        else:
                            cells[i].value = val.strftime(py_fmt)
                        cells[i].number_format = "@"
                        cells[i].data_type = "s"

                    # Handle text strings — VECTORIZED with pd.to_datetime
                    if str_indices:
                        text_series = pd.Series([str_vals[i] for i in str_indices])
                        # Strip timezone indicators (Z, +offset) for parsing
                        cleaned = text_series.str.replace(r"[Zz]$", "", regex=True).str.strip()
                        cleaned = cleaned.str.replace(r"[+-]\d{2}:?\d{2}$", "", regex=True).str.strip()
                        # Replace T separator with space for strptime compatibility
                        cleaned = cleaned.str.replace("T", " ", regex=False)

                        # Try vectorized parsing with known format
                        parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt)
                        if not parse_fmt:
                            parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt.replace(" HH:MM:SS", ""))
                            # If base format found but source has time, append time pattern
                            if parse_fmt and "HH:MM:SS" in src_fmt:
                                parse_fmt = parse_fmt + " %H:%M:%S"

                        parsed = None
                        if parse_fmt:
                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore")
                                parsed = pd.to_datetime(cleaned, format=parse_fmt, errors="coerce")
                        # Fallback: try ISO format (handles "2023-04-12 20:34" after T→space)
                        if parsed is None or parsed.isna().all():
                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore")
                                try:
                                    parsed = pd.to_datetime(cleaned, format="ISO8601", errors="coerce")
                                except (ValueError, TypeError):
                                    _dayfirst = src_fmt in _DAY_FIRST_LABELS
                                    try:
                                        parsed = pd.to_datetime(cleaned, errors="coerce",
                                                                dayfirst=_dayfirst, format="mixed")
                                    except TypeError:
                                        parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=_dayfirst)

                        # Write formatted results back
                        if parsed is not None:
                            formatted = parsed.dt.strftime(py_fmt)
                            for j, idx in enumerate(str_indices):
                                if pd.notna(parsed.iloc[j]) and str_vals[idx]:
                                    cells[idx].value = formatted.iloc[j]
                                    cells[idx].number_format = "@"
                                    cells[idx].data_type = "s"
                else:
                    # Native date format change — just update number_format (fast)
                    excel_fmt = _DATE_FMT_LABEL_TO_EXCEL.get(target_fmt, target_fmt)
                    for cell in cells:
                        if cell.value is not None:
                            cell.number_format = excel_fmt
                results[fix_idx] = True
                modified = True

            elif to_type == "date":
                # --- Text/Numeric → Date (bulk via pandas) ---
                src_fmt = fix_data.get("source_format", "")
                _dayfirst = src_fmt in _DAY_FIRST_LABELS

                # Clean text values for parsing
                text_series = pd.Series([str(v).strip() if v is not None else "" for v in raw_values])
                cleaned = text_series.str.replace(r"[Zz]$", "", regex=True).str.strip()
                cleaned = cleaned.str.replace(r"[+-]\d{2}:?\d{2}$", "", regex=True).str.strip()
                cleaned = cleaned.str.replace("T", " ", regex=False)

                # Try vectorized parse with known format first
                parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt)
                if not parse_fmt:
                    parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt.replace(" HH:MM:SS", ""))
                    if parse_fmt and "HH:MM:SS" in src_fmt:
                        parse_fmt = parse_fmt + " %H:%M:%S"

                converted = None
                if parse_fmt:
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore")
                        converted = pd.to_datetime(cleaned, format=parse_fmt, errors="coerce")
                if converted is None or converted.isna().all():
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore")
                        converted = pd.to_datetime(cleaned, errors="coerce", dayfirst=_dayfirst)

                excel_fmt = target_fmt if target_fmt and target_fmt not in ("General", "text-date") else "YYYY-MM-DD"
                excel_fmt = _DATE_FMT_LABEL_TO_EXCEL.get(excel_fmt, excel_fmt)

                for i, cell in enumerate(cells):
                    if pd.notna(converted.iloc[i]):
                        cell.value = converted.iloc[i].to_pydatetime()
                        cell.number_format = excel_fmt
                results[fix_idx] = True
                modified = True

            else:
                # Unsupported conversion — skip
                continue

        except Exception:
            # If one column fails, continue with others
            continue

    # Save the workbook if any changes were made
    if modified:
        try:
            wb.save(file_path)
        except Exception:
            # Save failed — mark all as unsuccessful
            return [False] * len(data_type_fixes)

    try:
        wb.close()
    except Exception:
        pass

    return results


def _xlsx_col_letter(col_idx):
    """Convert 1-based column index to Excel column letter (A, B, ..., Z, AA, AB, ...)."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return ''.join(reversed(result))


def _write_df_to_xlsx_fast(df, file_path, sheet_name="Sheet1", date_formats=None):
    """Write a DataFrame to XLSX using vectorized XML generation.

    ~8x faster than openpyxl write_only mode by:
    - Building cell XML column-by-column using numpy vectorized string ops
    - Binary-tree reduction to join columns (O(log N) passes instead of O(N))
    - Writing the ZIP package directly (no per-cell Python object creation)
    - Using inline strings (avoids shared strings table overhead for high-cardinality data)

    Args:
        df: DataFrame to write
        file_path: output .xlsx path
        sheet_name: Excel sheet name
        date_formats: optional dict {col_name: excel_format_code} for native date columns

    Falls back to openpyxl write_only if the vectorized path fails.
    """
    rows_count = len(df)
    cols_count = len(df.columns)

    if rows_count == 0 or cols_count == 0:
        # Edge case: empty DataFrame — use openpyxl for simplicity
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title=sheet_name)
        ws.append(list(df.columns))
        wb.save(file_path)
        wb.close()
        return

    try:
        _write_df_to_xlsx_vectorized(df, file_path, sheet_name, date_formats=date_formats)
    except Exception:
        # Fallback: openpyxl write_only mode
        df_clean = df.where(df.notna(), None)
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title=sheet_name)
        ws.append(list(df_clean.columns))
        for row in df_clean.values.tolist():
            ws.append(row)
        wb.save(file_path)
        wb.close()


def _write_df_to_xlsx_vectorized(df, file_path, sheet_name="Sheet1", date_formats=None):
    """Core vectorized XLSX writer. Builds XML using numpy char operations.

    Args:
        date_formats: optional dict {col_name: excel_format_code} for native date styling.
    """
    rows_count = len(df)
    cols_count = len(df.columns)

    # Pre-compute column letters
    col_letters = [_xlsx_col_letter(i + 1) for i in range(cols_count)]

    # Row number strings (data rows start at row 2; row 1 is the header)
    row_nums_np = np.arange(2, rows_count + 2).astype('U10')

    # Track date styles: map col_name → style_index (1-based, 0 is "General")
    date_formats = date_formats or {}
    # Build unique format list: each unique Excel format code gets its own style index
    _unique_fmts = list(dict.fromkeys(date_formats.values()))  # preserve order, deduplicate
    _fmt_to_style = {fmt: idx + 1 for idx, fmt in enumerate(_unique_fmts)}
    # Map column name → style index string
    _col_style = {col_name: str(_fmt_to_style[fmt]) for col_name, fmt in date_formats.items()}
    has_dates = bool(_unique_fmts)

    # Build cell XML for each column using vectorized operations
    col_arrays = []
    for ci in range(cols_count):
        col = df.iloc[:, ci]
        letter = col_letters[ci]
        refs = np.char.add(letter, row_nums_np)
        col_name = df.columns[ci]

        dtype = col.dtype
        if dtype in (np.int64, np.int32, np.int16, np.int8, np.uint64, np.uint32, np.uint16, np.uint8):
            # Integer column: <c r="A2"><v>123</v></c>
            vals_str = col.values.astype('U20')
            arr = np.char.add(
                np.char.add('<c r="', refs),
                np.char.add('"><v>', np.char.add(vals_str, '</v></c>'))
            )
            col_arrays.append(arr)
        elif hasattr(dtype, 'kind') and dtype.kind == 'M':
            # Datetime64 column: convert to Excel serial number with style reference
            style_idx = _col_style.get(col_name, '1')
            mask = col.notna().values
            # Excel epoch: 1899-12-30 (Excel's day-0 is 1900-01-00 which is 1899-12-30)
            excel_epoch = pd.Timestamp('1899-12-30')
            # Compute serial numbers: days + fraction for time
            delta = (col - excel_epoch)
            serial = delta.dt.total_seconds() / 86400.0
            # Format as string (enough precision for time)
            serial_vals = serial.values
            vals_str = np.where(mask, np.char.mod('%.10g', serial_vals), '')
            # Build XML: <c r="A2" s="1"><v>45678.5</v></c>
            full = np.char.add(
                np.char.add('<c r="', refs),
                np.char.add('" s="' + style_idx + '"><v>', np.char.add(vals_str, '</v></c>'))
            )
            arr = np.where(mask, full, '')
            col_arrays.append(arr)
        elif dtype in (np.float64, np.float32):
            # Float column with NaN handling
            mask = np.isfinite(col.values)
            vals_str = col.values.astype('U20')
            full = np.char.add(
                np.char.add('<c r="', refs),
                np.char.add('"><v>', np.char.add(vals_str, '</v></c>'))
            )
            arr = np.where(mask, full, '')
            col_arrays.append(arr)
        else:
            # Object/string column: inline string format
            mask = col.notna().values
            s = col.fillna('').astype(str)
            # XML escape using pandas vectorized str methods
            s = s.str.replace('&', '&amp;', regex=False)
            s = s.str.replace('<', '&lt;', regex=False)
            s = s.str.replace('>', '&gt;', regex=False)
            s_np = s.values.astype('U')
            full = np.char.add(
                np.char.add('<c r="', refs),
                np.char.add('" t="inlineStr"><is><t>', np.char.add(s_np, '</t></is></c>'))
            )
            arr = np.where(mask, full, '')
            col_arrays.append(arr)

    # Binary-tree reduction to join all columns (O(log N) passes)
    arrays = list(col_arrays)
    while len(arrays) > 1:
        new = []
        for i in range(0, len(arrays), 2):
            if i + 1 < len(arrays):
                new.append(np.char.add(arrays[i], arrays[i + 1]))
            else:
                new.append(arrays[i])
        arrays = new
    row_content = arrays[0]

    # Wrap each row in <row> tags
    row_xmls = np.char.add(
        np.char.add('<row r="', np.char.add(row_nums_np, '">')),
        np.char.add(row_content, '</row>')
    )

    # Join all rows into the sheet body
    body = ''.join(row_xmls)

    # Build header row XML
    hdr_cells = ''.join(
        '<c r="{letter}1" t="inlineStr"><is><t>{val}</t></is></c>'.format(
            letter=col_letters[ci],
            val=str(df.columns[ci]).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        )
        for ci in range(cols_count)
    )

    # Assemble complete sheet XML
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetData><row r="1">' + hdr_cells + '</row>' + body + '</sheetData></worksheet>'
    ).encode('utf-8')

    # Escape sheet name for XML attribute
    sn_esc = (sheet_name.replace('&', '&amp;').replace('<', '&lt;')
              .replace('>', '&gt;').replace('"', '&quot;'))

    # Build styles.xml if date columns exist
    styles_xml = None
    if has_dates:
        # Build numFmts (custom format IDs start at 164)
        numfmt_entries = ''.join(
            f'<numFmt numFmtId="{164 + i}" formatCode="{fmt}"/>'
            for i, fmt in enumerate(_unique_fmts)
        )
        # Build cellXfs: style 0 = General, then one per date format
        xf_entries = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        for i, fmt in enumerate(_unique_fmts):
            xf_entries += f'<xf numFmtId="{164 + i}" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'

        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<numFmts count="{len(_unique_fmts)}">{numfmt_entries}</numFmts>'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{1 + len(_unique_fmts)}">{xf_entries}</cellXfs>'
            '</styleSheet>'
        )

    # Build minimal XLSX package parts
    styles_ct = (
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        if has_dates else ''
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        + styles_ct +
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{sn_esc}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    styles_rel = (
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        if has_dates else ''
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        + styles_rel +
        '</Relationships>'
    )

    # Write ZIP file (compresslevel=1 for speed; still achieves good compression)
    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        if styles_xml:
            zf.writestr('xl/styles.xml', styles_xml)


def _cast_df_columns_to_detected_types(df, csv_types, fixed_col_indices):
    """Convert string columns to their detected types (numeric/date) for proper XLSX output.

    When a CSV is read with dtype=str, all columns are text. This function converts
    columns to their actual data types so the output XLSX has correct cell types:
    - "numeric" → pd.to_numeric (written as Excel numbers)
    - "date" → pd.to_datetime (written as native Excel dates)
    - "text"/"unknown" → left as-is (written as inline strings)

    Columns in `fixed_col_indices` are SKIPPED because they already had explicit
    fix_data_type fixes applied.

    Args:
        df: DataFrame (read from CSV, fixes already applied)
        csv_types: dict from _detect_csv_column_types {col_idx: {"type":..., "format":..., "header":...}}
        fixed_col_indices: set of column indices that already had fix_data_type applied

    Returns:
        (df, date_fmt_map) where date_fmt_map = {col_name: excel_format_code} for date columns
    """
    date_fmt_map = {}  # col_name -> Excel format code (e.g., "yyyy-mm-dd")

    for col_idx, info in csv_types.items():
        if col_idx in fixed_col_indices:
            continue
        if col_idx >= len(df.columns):
            continue

        col_name = df.columns[col_idx]
        col_type = info["type"]
        col_fmt = info.get("format", "")

        if col_type == "numeric":
            # Convert string "123" to actual numeric
            try:
                df[col_name] = pd.to_numeric(df[col_name], errors="coerce")
            except Exception:
                pass  # Leave as string if conversion fails

        elif col_type == "date":
            # Convert text date strings to datetime for native Excel date storage
            parse_fmt = _VECTORIZED_PARSE_FORMATS.get(col_fmt)
            if not parse_fmt:
                # Try without time suffix
                parse_fmt = _VECTORIZED_PARSE_FORMATS.get(col_fmt.replace(" HH:MM:SS", ""))

            if parse_fmt:
                try:
                    # Clean common artifacts before parsing
                    series = df[col_name].copy()
                    mask = series.notna() & (series.astype(str).str.strip() != "")
                    if mask.any():
                        clean = series[mask].astype(str).str.strip()
                        # Normalize ISO-8601 variants: "2025-11-12T22:07Z" → "2025-11-12 22:07"
                        if clean.str.contains('T', na=False).any():
                            clean = clean.str.replace('T', ' ', regex=False)
                            clean = clean.str.replace('Z', '', regex=False)
                            # Remove timezone offset like +05:30 or -04:00
                            clean = clean.str.replace(r'[+-]\d{2}:\d{2}$', '', regex=True)
                        # Try with full format first (may include time)
                        parsed = pd.to_datetime(clean, format=parse_fmt, errors="coerce")
                        # Fallback: try without seconds if many NaT
                        nat_count = parsed.isna().sum()
                        if nat_count > len(parsed) * 0.5:
                            nosec_fmt = _VECTORIZED_PARSE_FORMATS_NO_SEC.get(col_fmt)
                            if nosec_fmt:
                                parsed = pd.to_datetime(clean, format=nosec_fmt, errors="coerce")
                        # Last fallback: let pandas infer (handles mixed formats)
                        if parsed.isna().sum() > len(parsed) * 0.5:
                            parsed = pd.to_datetime(clean, errors="coerce", dayfirst=col_fmt in _DAY_FIRST_LABELS)
                        # Only convert if at least half parsed successfully
                        if parsed.notna().sum() >= len(parsed) * 0.5:
                            # Build full-length datetime column (NaT for empty/missing rows)
                            full_dt = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
                            full_dt.iloc[mask.values.nonzero()[0]] = parsed.values
                            df[col_name] = full_dt
                            # Determine Excel format code
                            excel_fmt = _DATE_FMT_LABEL_TO_EXCEL.get(col_fmt, "yyyy-mm-dd")
                            date_fmt_map[col_name] = excel_fmt
                except Exception:
                    pass  # Leave as string if conversion fails

    return df, date_fmt_map

def apply_fixes(findings, output_folder):
    """Apply fixable findings grouped by file.

    Fix order per file:
      1. convert_file_type (CSV→XLSX, XLS→XLSX, etc.) — batch via Excel COM
      2. rename_sheet
      3. rename_column, fix_data_type, remove_extra_columns (workbook ops)
      4. rename_file — must happen LAST

    FAST PATH: When a CSV needs conversion to XLSX AND has data type fixes,
    everything is done in one pass with pandas (read CSV → fix → write XLSX).
    This avoids Excel COM + openpyxl entirely (10-30x faster for large files).

    SINGLE-PASS XLSX: All XLSX fixes (data types, rename sheet/column, remove
    columns) are applied in ONE openpyxl open→fix→save cycle. No Excel COM
    needed for workbook-level operations.

    Returns (fixed_count, rename_map) where rename_map tracks old→new filenames.
    """
    import time as _time
    _t0 = _time.perf_counter()
    from collections import defaultdict
    grouped = defaultdict(list)
    for f in findings:
        if f.get("fix_type"):
            grouped[f["File"]].append(f)

    # Define execution order
    FIX_ORDER = {
        "convert_file_type": 0,
        "rename_sheet": 1,
        "rename_column": 2,
        "fix_data_type": 3,
        "remove_extra_columns": 4,
        "rename_file": 5,
    }

    # === FAST PATH: CSV→XLSX conversion + fixes in one pandas pass ===
    # Identify CSV files that need conversion AND have workbook-level fixes.
    # For these, we skip COM entirely and do: read CSV → apply fixes → write XLSX.
    _fast_converted = {}  # original_name -> new_name (files handled by fast path)

    for original_name, file_findings in grouped.items():
        file_path = os.path.join(output_folder, original_name)
        if not os.path.exists(file_path):
            continue
        cur_ext = Path(original_name).suffix.lower()
        if cur_ext != ".csv":
            continue

        # Check if this CSV needs conversion to xlsx
        conv_finding = None
        for f in file_findings:
            if f["fix_type"] == "convert_file_type" and f["fix_data"].get("to_ext", "") == ".xlsx":
                conv_finding = f
                break
        if not conv_finding:
            continue

        # Gather data type + column rename + remove_extra fixes
        wb_fixes = [f for f in file_findings if f["fix_type"] in ("fix_data_type", "rename_column", "remove_extra_columns")]
        if not wb_fixes and not conv_finding:
            continue

        # FAST PATH: read CSV, apply all fixes, write directly as XLSX
        try:
            import time as _time
            _fp_t0 = _time.perf_counter()

            # --- PROFILING: Read CSV ---
            _prof_read_t0 = _time.perf_counter()
            df = pd.read_csv(file_path, dtype=str, low_memory=False)
            _prof_read_elapsed = _time.perf_counter() - _prof_read_t0
            _file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            print(f"\n[CSV FAST PROFILE]")
            print(f"  file = {original_name}")
            print(f"  rows = {len(df)}")
            print(f"  columns = {len(df.columns)}")
            print(f"  size_mb = {_file_size_mb:.1f}")
            print(f"  read_csv = {_prof_read_elapsed:.2f} sec")

            csv_modified = False
            fixed_in_fast = 0
            _fix_num = 0
            _date_fmt_map = {}  # col_name -> Excel format code for native dates

            # Build a name-based template-type map while column indices still
            # match the original CSV layout (before renames / drops shift them).
            _matched_col_types_raw = conv_finding["fix_data"].get("matched_col_types", {})
            _tpl_type_by_name = {}  # lowercase col name → template type str
            for _mci, _mt in _matched_col_types_raw.items():
                if _mci < len(df.columns):
                    _tpl_type_by_name[df.columns[_mci].strip().lower()] = _mt

            # --- Serialization plan: ONE authoritative type per column --------
            # Authority order, highest first:
            #   1. an explicit fix_data_type finding - applied in the loop below
            #   2. the TEMPLATE type, when conclusive - it agrees with the placed
            #      file anyway, otherwise check 6 would have raised a finding
            #   3. the PLACED CSV's OWN detected type - preserved as-is
            # Rule 3 is the fix: previously a column whose template counterpart
            # was blank in the sampled rows ("unknown") got no finding AND no
            # cast, so it fell through to inlineStr and arrived in Excel as text.
            _placed_types_raw = _detect_csv_column_types(file_path)
            _plan = {}  # lowercase current col name -> {"type":..., "format":...}
            for _pci, _pinfo in _placed_types_raw.items():
                if _pci >= len(df.columns):
                    continue
                _t = _matched_col_types_raw.get(_pci)
                if not _t or _t == "unknown":
                    _t = _pinfo.get("type", "unknown")
                _plan[df.columns[_pci].strip().lower()] = {
                    "type": _t,
                    "format": _pinfo.get("format", "General"),
                }
            _fixed_col_names = set()  # columns owned by an explicit fix

            for finding in sorted(wb_fixes, key=lambda f: FIX_ORDER.get(f["fix_type"], 99)):
                fix_type = finding["fix_type"]
                fix_data = finding["fix_data"]
                _fix_num += 1
                _fix_t0 = _time.perf_counter()
                try:
                    if fix_type == "rename_column":
                        old_col = fix_data["old_col"]
                        new_col = fix_data["new_col"]
                        if old_col in df.columns:
                            df = df.rename(columns={old_col: new_col})
                            # Keep _tpl_type_by_name in sync with the rename
                            _old_key = old_col.strip().lower()
                            if _old_key in _tpl_type_by_name:
                                _tpl_type_by_name[new_col.strip().lower()] = _tpl_type_by_name.pop(_old_key)
                            if _old_key in _plan:
                                _plan[new_col.strip().lower()] = _plan.pop(_old_key)
                            csv_modified = True
                            fixed_in_fast += 1
                            finding["_fixed"] = True
                        _fix_elapsed = _time.perf_counter() - _fix_t0
                        print(f"  fix #{_fix_num}: type=rename_column  column={old_col!r}  elapsed={_fix_elapsed:.2f} sec")
                    elif fix_type == "fix_data_type":
                        col_idx = fix_data["col_idx"]
                        from_type = fix_data.get("from_type", "")
                        to_type = fix_data["to_type"]
                        target_fmt = fix_data.get("target_format", "")
                        if col_idx < len(df.columns):
                            col_name = df.columns[col_idx]
                            if from_type == "date" and to_type == "date":
                                py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                                src_fmt = fix_data.get("source_format", "")
                                _vectorized_date_reformat(df, col_name, src_fmt, py_fmt)
                                # Convert reformatted strings to datetime for native Excel date storage
                                try:
                                    df[col_name] = pd.to_datetime(df[col_name], format=py_fmt, errors="coerce")
                                    _date_fmt_map[col_name] = _DATE_FMT_LABEL_TO_EXCEL.get(target_fmt, "yyyy-mm-dd")
                                except Exception:
                                    pass  # Leave as string if conversion fails
                            elif to_type == "numeric":
                                df[col_name] = pd.to_numeric(df[col_name], errors="coerce")
                            elif to_type == "date":
                                target_py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                                src_fmt = fix_data.get("source_format", "")
                                _vectorized_date_reformat(df, col_name, src_fmt, target_py_fmt)
                                # Convert to datetime for native Excel date storage
                                try:
                                    df[col_name] = pd.to_datetime(df[col_name], format=target_py_fmt, errors="coerce")
                                    _date_fmt_map[col_name] = _DATE_FMT_LABEL_TO_EXCEL.get(target_fmt, "yyyy-mm-dd")
                                except Exception:
                                    pass
                            elif to_type == "text":
                                df[col_name] = df[col_name].astype(str).replace("nan", "")
                            _fixed_col_names.add(str(col_name).strip().lower())
                            csv_modified = True
                            fixed_in_fast += 1
                            finding["_fixed"] = True
                        _fix_elapsed = _time.perf_counter() - _fix_t0
                        print(f"  fix #{_fix_num}: type=fix_data_type  column={col_name if col_idx < len(df.columns) else '?'}  from={from_type}  to={to_type}  src_fmt={fix_data.get('source_format','')}  tgt_fmt={target_fmt}  elapsed={_fix_elapsed:.2f} sec")
                    elif fix_type == "remove_extra_columns":
                        _extra_names = fix_data.get("extra_col_names", [])
                        if _extra_names:
                            _to_drop = [c for c in df.columns if c.strip() in _extra_names]
                            if _to_drop:
                                df = df.drop(columns=_to_drop)
                                csv_modified = True
                                fixed_in_fast += 1
                                finding["_fixed"] = True
                        else:
                            keep_count = fix_data["keep_count"]
                            if len(df.columns) > keep_count:
                                df = df.iloc[:, :keep_count]
                                csv_modified = True
                                fixed_in_fast += 1
                                finding["_fixed"] = True
                        _fix_elapsed = _time.perf_counter() - _fix_t0
                        print(f"  fix #{_fix_num}: type=remove_extra_columns  elapsed={_fix_elapsed:.2f} sec")
                except Exception:
                    _fix_elapsed = _time.perf_counter() - _fix_t0
                    print(f"  fix #{_fix_num}: type={fix_type}  FAILED  elapsed={_fix_elapsed:.2f} sec")
                    continue

            # Write as XLSX directly using vectorized XML generation (fastest method)
            new_name = Path(original_name).stem + ".xlsx"
            new_path = os.path.join(output_folder, new_name)
            sheet_name = conv_finding["fix_data"].get("template_sheet", "Sheet1") or "Sheet1"

            # --- Apply the serialization plan --------------------------------
            # Each column is written with the type resolved above. A column that
            # an explicit fix already converted is left alone; a column planned
            # as text/unknown stays text. Nothing is changed without a reason.
            _cast_num, _cast_dt = [], []
            for _cn in list(df.columns):
                _key = str(_cn).strip().lower()
                if _key in _fixed_col_names:
                    continue                      # an explicit fix owns it
                if (pd.api.types.is_numeric_dtype(df[_cn])
                        or pd.api.types.is_datetime64_any_dtype(df[_cn])):
                    continue                      # already typed
                _p = _plan.get(_key)
                if not _p:
                    continue                      # no plan -> leave untouched
                if _p["type"] == "numeric":
                    _s = df[_cn].dropna().head(200)
                    if _s.empty:
                        continue
                    if pd.to_numeric(_s, errors="coerce").notna().mean() >= 0.8:
                        df[_cn] = pd.to_numeric(df[_cn], errors="coerce")
                        _cast_num.append(_cn)
                elif _p["type"] == "date":
                    _one = {list(df.columns).index(_cn): {
                        "type": "date", "format": _p["format"], "header": _cn}}
                    df, _dm = _cast_df_columns_to_detected_types(df, _one, set())
                    if _dm:
                        _date_fmt_map.update(_dm)
                        _cast_dt.append(_cn)
            print(f"  [plan] numeric columns     = {_cast_num}")
            print(f"  [plan] native-date columns = {_cast_dt}")

            # --- PROFILING: Write XLSX ---
            _wr_t0 = _time.perf_counter()
            _write_df_to_xlsx_fast(df, new_path, sheet_name, date_formats=_date_fmt_map)
            _wr_t1 = _time.perf_counter()

            print(f"  _write_df_to_xlsx_fast (vectorized):")
            print(f"    total write           = {_wr_t1 - _wr_t0:.2f} sec")

            _fp_elapsed = _time.perf_counter() - _fp_t0
            print(f"  TOTAL FAST PATH = {_fp_elapsed:.2f} sec")
            print(f"[apply_fixes] fast_path({original_name}): {_fp_elapsed:.2f}s | {fixed_in_fast} fixes + convert")

            # Remove original CSV
            try:
                os.remove(file_path)
            except Exception:
                pass

            conv_finding["_fixed"] = True
            _fast_converted[original_name] = new_name

        except Exception:
            # Fast path failed — fall through to normal COM conversion
            continue

    # === Phase 1: Batch remaining file conversions via Excel COM ===
    _conversion_jobs = []  # (src, dst, sheet, finding, original_name)
    _name_after_convert = {}  # original_name -> new_name after conversion
    # Include fast-path converted files in the name mapping
    _name_after_convert.update(_fast_converted)

    for original_name, file_findings in grouped.items():
        if original_name in _fast_converted:
            continue  # Already handled by fast path
        file_path = os.path.join(output_folder, original_name)
        if not os.path.exists(file_path):
            continue
        for finding in file_findings:
            if finding["fix_type"] == "convert_file_type":
                fix_data = finding["fix_data"]
                to_ext = fix_data["to_ext"]
                new_name = Path(original_name).stem + to_ext
                new_path = os.path.join(output_folder, new_name)
                sheet = fix_data.get("template_sheet")
                _conversion_jobs.append((file_path, new_path, sheet, finding, original_name, new_name))
                break  # Only one conversion per file

    # Run all conversions in a single Excel COM session
    if _conversion_jobs:
        batch_args = [(src, dst, sheet) for src, dst, sheet, _, _, _ in _conversion_jobs]
        results = _excel_convert_batch(batch_args)
        for i, success in enumerate(results):
            src, dst, sheet, finding, orig_name, new_name = _conversion_jobs[i]
            if success:
                try:
                    os.remove(src)
                except Exception:
                    pass
                _name_after_convert[orig_name] = new_name
                finding["_fixed"] = True

    # === Phase 2 & 3: Workbook ops + renames (parallel across files) ===
    def _fix_one_file(item):
        """Fix workbook ops + rename for a single file."""
        import time as _time
        _ft0 = _time.perf_counter()
        original_name, file_findings = item
        current_name = _name_after_convert.get(original_name, original_name)
        file_path = os.path.join(output_folder, current_name)
        if not os.path.exists(file_path):
            return (0, original_name, None)

        local_count = 0
        file_findings.sort(key=lambda f: FIX_ORDER.get(f["fix_type"], 99))

        # Skip findings already handled by fast path
        rename_fixes = [f for f in file_findings if f["fix_type"] == "rename_file" and not f.get("_fixed")]
        wb_fixes = [f for f in file_findings if f["fix_type"] in ("rename_sheet", "rename_column", "fix_data_type", "remove_extra_columns") and not f.get("_fixed")]

        # Count fixes already done by fast path
        local_count += sum(1 for f in file_findings if f.get("_fixed") and f["fix_type"] != "convert_file_type")

        # Step 2: Batch all workbook operations (open once, apply all, save once)
        if wb_fixes:
            file_path = os.path.join(output_folder, current_name)
            cur_ext = Path(current_name).suffix.lower()

            if cur_ext == ".csv":
                # CSV: use pandas (read once, apply all column/type fixes, write once)
                try:
                    df = pd.read_csv(file_path, dtype=str, low_memory=False)
                    csv_modified = False
                    for finding in wb_fixes:
                        fix_type = finding["fix_type"]
                        fix_data = finding["fix_data"]
                        _b4 = local_count
                        try:
                            if fix_type == "rename_column":
                                old_col = fix_data["old_col"]
                                new_col = fix_data["new_col"]
                                if old_col in df.columns:
                                    df = df.rename(columns={old_col: new_col})
                                    csv_modified = True
                                    local_count += 1
                            elif fix_type == "fix_data_type":
                                col_idx = fix_data["col_idx"]
                                from_type = fix_data.get("from_type", "")
                                to_type = fix_data["to_type"]
                                target_fmt = fix_data.get("target_format", "")
                                if col_idx < len(df.columns):
                                    col_name = df.columns[col_idx]
                                    if from_type == "date" and to_type == "date":
                                        # Date format change (e.g. M/D/YYYY → YYYY-MM-DD)
                                        # FAST vectorized approach: use pd.to_datetime with known format
                                        py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                                        src_fmt = fix_data.get("source_format", "")
                                        _vectorized_date_reformat(df, col_name, src_fmt, py_fmt)
                                    elif to_type == "numeric":
                                        df[col_name] = pd.to_numeric(df[col_name], errors="coerce")
                                    elif to_type == "date":
                                        target_py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                                        src_fmt = fix_data.get("source_format", "")
                                        _vectorized_date_reformat(df, col_name, src_fmt, target_py_fmt)
                                    elif to_type == "text":
                                        df[col_name] = df[col_name].astype(str).replace("nan", "")
                                    csv_modified = True
                                    local_count += 1
                            elif fix_type == "remove_extra_columns":
                                _extra_names = fix_data.get("extra_col_names", [])
                                if _extra_names:
                                    # Drop extra columns by name
                                    _to_drop = [c for c in df.columns if c.strip() in _extra_names]
                                    if _to_drop:
                                        df = df.drop(columns=_to_drop)
                                        csv_modified = True
                                        local_count += 1
                                else:
                                    # Fallback: keep first N columns
                                    keep_count = fix_data["keep_count"]
                                    if len(df.columns) > keep_count:
                                        df = df.iloc[:, :keep_count]
                                        csv_modified = True
                                        local_count += 1
                        except Exception:
                            continue
                        if local_count > _b4:
                            finding["_fixed"] = True
                    if csv_modified:
                        df.to_csv(file_path, index=False)
                except Exception:
                    pass

            elif cur_ext in (".xlsx", ".xls") and os.path.exists(file_path):
                # === SINGLE-PASS: All fixes in ONE openpyxl session ===
                # Opens workbook ONCE, applies ALL fix types, saves ONCE.
                # No Excel COM needed for rename_sheet/rename_column/remove_extra_columns.
                try:
                    _xl_t0 = _time.perf_counter()
                    wb = openpyxl.load_workbook(file_path, keep_links=False)
                    wb_modified = False

                    for finding in wb_fixes:
                        fix_type = finding["fix_type"]
                        fix_data = finding["fix_data"]
                        try:
                            if fix_type == "rename_sheet":
                                old_name = fix_data["sheet_old"]
                                new_name = fix_data["sheet_new"]
                                if old_name in wb.sheetnames:
                                    wb[old_name].title = new_name
                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                            elif fix_type == "rename_column":
                                old_col = fix_data.get("old_col", "")
                                new_col = fix_data.get("new_col", "")
                                sheet_name = fix_data.get("sheet", "")
                                ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
                                # Scan header rows (first 30 rows) for the old column name
                                renamed = False
                                for row in ws.iter_rows(min_row=1, max_row=30, max_col=ws.max_column or 100):
                                    for cell in row:
                                        if cell.value is not None and str(cell.value).strip() == old_col:
                                            cell.value = new_col
                                            renamed = True
                                            break
                                    if renamed:
                                        break
                                if renamed:
                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                            elif fix_type == "remove_extra_columns":
                                extra_names = fix_data.get("extra_col_names", [])
                                ws = wb.active
                                if extra_names:
                                    # Find and delete columns by header name (reverse order to preserve indices)
                                    cols_to_delete = []
                                    for col_idx_1b in range(1, (ws.max_column or 100) + 1):
                                        hdr = ws.cell(row=1, column=col_idx_1b).value
                                        if hdr is not None and str(hdr).strip() in extra_names:
                                            cols_to_delete.append(col_idx_1b)
                                    # Delete from right to left
                                    for col_idx_1b in sorted(cols_to_delete, reverse=True):
                                        ws.delete_cols(col_idx_1b)
                                    if cols_to_delete:
                                        wb_modified = True
                                        local_count += 1
                                        finding["_fixed"] = True
                                else:
                                    keep_count = fix_data.get("keep_count", 0)
                                    if keep_count and ws.max_column and ws.max_column > keep_count:
                                        ws.delete_cols(keep_count + 1, ws.max_column - keep_count)
                                        wb_modified = True
                                        local_count += 1
                                        finding["_fixed"] = True

                            elif fix_type == "fix_data_type":
                                # Data type fix — inline in same session
                                col_idx = fix_data.get("col_idx", 0)
                                hdr_row = fix_data.get("hdr_row", 1)
                                from_type = fix_data.get("from_type", "")
                                to_type = fix_data.get("to_type", "")
                                target_fmt = fix_data.get("target_format", "")
                                src_fmt = fix_data.get("source_format", "")
                                sheet_name = fix_data.get("sheet", "")

                                ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
                                excel_col = col_idx + 1
                                start_row = hdr_row + 1
                                max_row = ws.max_row or start_row

                                if max_row < start_row:
                                    local_count += 1
                                    finding["_fixed"] = True
                                    continue

                                # Bulk read column cells
                                cells = []
                                for row_tuple in ws.iter_rows(min_row=start_row, max_row=max_row,
                                                              min_col=excel_col, max_col=excel_col):
                                    cells.append(row_tuple[0])
                                raw_values = [c.value for c in cells]

                                if to_type == "text":
                                    for cell in cells:
                                        if cell.value is not None:
                                            val = cell.value
                                            if isinstance(val, float):
                                                cell.value = str(int(val)) if val == int(val) else str(val)
                                            elif isinstance(val, datetime.datetime):
                                                cell.value = val.strftime("%Y-%m-%d")
                                            else:
                                                cell.value = str(val)
                                            cell.number_format = "@"
                                            cell.data_type = "s"
                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                                elif to_type == "numeric":
                                    converted = pd.to_numeric(pd.Series(raw_values), errors="coerce")
                                    for i, cell in enumerate(cells):
                                        if pd.notna(converted.iloc[i]):
                                            num_val = converted.iloc[i]
                                            cell.value = int(num_val) if num_val == int(num_val) else float(num_val)
                                            cell.number_format = "General"
                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                                elif to_type == "date" and from_type == "date":
                                    # Date→Date format change
                                    _tpl_storage = fix_data.get("template_storage", "")
                                    is_tpl_text = _tpl_storage == "text_date" or (
                                        _tpl_storage not in ("native_date", "mixed") and target_fmt in _TEXT_DATE_LABELS)

                                    if is_tpl_text:
                                        # Template stores dates as text — reformat to text strings
                                        py_fmt = _DATE_FMT_LABEL_TO_STRFTIME.get(target_fmt, "%Y-%m-%d")
                                        str_vals = []
                                        dt_indices = []
                                        str_indices = []
                                        for i, val in enumerate(raw_values):
                                            if val is None:
                                                str_vals.append("")
                                            elif isinstance(val, (datetime.datetime, datetime.date)):
                                                dt_indices.append(i)
                                                str_vals.append("")
                                            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                                                dt_indices.append(i)
                                                str_vals.append("")
                                            else:
                                                str_indices.append(i)
                                                str_vals.append(str(val).strip())

                                        for i in dt_indices:
                                            val = raw_values[i]
                                            if isinstance(val, (int, float)):
                                                try:
                                                    dt = (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))).to_pydatetime()
                                                    cells[i].value = dt.strftime(py_fmt)
                                                except Exception:
                                                    continue
                                            elif isinstance(val, datetime.date) and not isinstance(val, datetime.datetime):
                                                dt = datetime.datetime(val.year, val.month, val.day)
                                                cells[i].value = dt.strftime(py_fmt)
                                            else:
                                                cells[i].value = val.strftime(py_fmt)
                                            cells[i].number_format = "@"
                                            cells[i].data_type = "s"

                                        if str_indices:
                                            text_series = pd.Series([str_vals[i] for i in str_indices])
                                            cleaned = text_series.str.replace(r"[Zz]$", "", regex=True).str.strip()
                                            cleaned = cleaned.str.replace(r"[+-]\d{2}:?\d{2}$", "", regex=True).str.strip()
                                            cleaned = cleaned.str.replace("T", " ", regex=False)

                                            parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt)
                                            if not parse_fmt:
                                                parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt.replace(" HH:MM:SS", ""))
                                                if parse_fmt and "HH:MM:SS" in src_fmt:
                                                    parse_fmt = parse_fmt + " %H:%M:%S"

                                            parsed = None
                                            if parse_fmt:
                                                with warnings.catch_warnings():
                                                    warnings.simplefilter("ignore")
                                                    parsed = pd.to_datetime(cleaned, format=parse_fmt, errors="coerce")
                                            if parsed is None or parsed.isna().all():
                                                with warnings.catch_warnings():
                                                    warnings.simplefilter("ignore")
                                                    try:
                                                        parsed = pd.to_datetime(cleaned, format="ISO8601", errors="coerce")
                                                    except (ValueError, TypeError):
                                                        _df = src_fmt in _DAY_FIRST_LABELS
                                                        try:
                                                            parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=_df, format="mixed")
                                                        except TypeError:
                                                            parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=_df)

                                            if parsed is not None:
                                                formatted = parsed.dt.strftime(py_fmt)
                                                for j, idx in enumerate(str_indices):
                                                    if pd.notna(parsed.iloc[j]) and str_vals[idx]:
                                                        cells[idx].value = formatted.iloc[j]
                                                        cells[idx].number_format = "@"
                                                        cells[idx].data_type = "s"
                                    else:
                                        # Native date — just update number_format
                                        excel_fmt = _DATE_FMT_LABEL_TO_EXCEL.get(target_fmt, target_fmt)
                                        for cell in cells:
                                            if cell.value is not None:
                                                cell.number_format = excel_fmt

                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                                elif to_type == "date":
                                    # Text/Numeric → Date
                                    _dayfirst = src_fmt in _DAY_FIRST_LABELS
                                    text_series = pd.Series([str(v).strip() if v is not None else "" for v in raw_values])
                                    cleaned = text_series.str.replace(r"[Zz]$", "", regex=True).str.strip()
                                    cleaned = cleaned.str.replace(r"[+-]\d{2}:?\d{2}$", "", regex=True).str.strip()
                                    cleaned = cleaned.str.replace("T", " ", regex=False)

                                    parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt)
                                    if not parse_fmt:
                                        parse_fmt = _VECTORIZED_PARSE_FORMATS.get(src_fmt.replace(" HH:MM:SS", ""))
                                        if parse_fmt and "HH:MM:SS" in src_fmt:
                                            parse_fmt = parse_fmt + " %H:%M:%S"

                                    converted = None
                                    if parse_fmt:
                                        with warnings.catch_warnings():
                                            warnings.simplefilter("ignore")
                                            converted = pd.to_datetime(cleaned, format=parse_fmt, errors="coerce")
                                    if converted is None or converted.isna().all():
                                        with warnings.catch_warnings():
                                            warnings.simplefilter("ignore")
                                            converted = pd.to_datetime(cleaned, errors="coerce", dayfirst=_dayfirst)

                                    excel_fmt_code = target_fmt if target_fmt and target_fmt not in ("General", "text-date") else "YYYY-MM-DD"
                                    excel_fmt_code = _DATE_FMT_LABEL_TO_EXCEL.get(excel_fmt_code, excel_fmt_code)

                                    for i, cell in enumerate(cells):
                                        if pd.notna(converted.iloc[i]):
                                            cell.value = converted.iloc[i].to_pydatetime()
                                            cell.number_format = excel_fmt_code
                                    wb_modified = True
                                    local_count += 1
                                    finding["_fixed"] = True

                        except Exception:
                            continue

                    # SAVE ONCE after all fixes applied
                    if wb_modified:
                        wb.save(file_path)
                    wb.close()
                    _xl_elapsed = _time.perf_counter() - _xl_t0
                    print(f"[apply_fixes] xlsx_pass({current_name}): {_xl_elapsed:.2f}s | {sum(1 for f in wb_fixes if f.get('_fixed'))} fixes")
                except Exception:
                    pass

        # Step 3: Handle file rename (must be last)
        for finding in rename_fixes:
            fix_data = finding["fix_data"]
            file_path = os.path.join(output_folder, current_name)
            if not os.path.exists(file_path):
                continue
            _b4 = local_count
            try:
                new_name = fix_data["new_name"]
                current_ext = Path(current_name).suffix.lower()
                target_ext = Path(new_name).suffix.lower()
                if current_ext != target_ext:
                    new_name = Path(new_name).stem + current_ext
                new_path = os.path.join(output_folder, new_name)
                if os.path.normcase(file_path) != os.path.normcase(new_path):
                    if os.path.exists(new_path) and os.path.normcase(file_path) != os.path.normcase(new_path):
                        os.remove(new_path)
                    os.rename(file_path, new_path)
                    current_name = new_name
                    local_count += 1
            except Exception:
                continue
            if local_count > _b4:
                finding["_fixed"] = True

        # Return result for this file
        final_name = current_name if current_name != original_name else None
        _ft_elapsed = _time.perf_counter() - _ft0
        if _ft_elapsed > 0.5:  # Only log slow files
            print(f"[apply_fixes] file({original_name}): {_ft_elapsed:.2f}s | fixes={local_count}")
        return (local_count, original_name, final_name)

    # --- Parallel execution across files ---
    # Count successful conversions (both fast-path and COM)
    fixed_count = len([j for j in _conversion_jobs if _name_after_convert.get(j[4])])  # COM conversions
    fixed_count += len(_fast_converted)  # Fast-path conversions
    rename_map = {}

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_fix_one_file, item) for item in grouped.items()]
        for fut in as_completed(futs):
            try:
                count, orig, new_name = fut.result()
                fixed_count += count
                if new_name:
                    rename_map[orig] = new_name
                elif orig in _name_after_convert:
                    # File was converted but not renamed further
                    rename_map[orig] = _name_after_convert[orig]
            except Exception:
                pass

    _elapsed = _time.perf_counter() - _t0
    print(f"[apply_fixes] TOTAL: {_elapsed:.2f}s | files={len(grouped)} | fixes_applied={fixed_count}")
    return fixed_count, rename_map


# ===== Study Config =====
import json

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "study_config.json")


def load_study_config():
    """Load study configuration from JSON file."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_study_config(config):
    """Save study configuration to JSON file."""
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)


def extract_study_name(filenames):
    """Extract study name from filenames. Looks for pattern xxx_MC_XXXX or xxx-MC-XXXX."""
    for f in filenames:
        match = re.search(r'[-_]mc[-_]([a-z0-9]{4})', f.lower())
        if match:
            return match.group(1).upper()
    return ""


# ===== UI =====

# --- Section 1: Select Files to be Placed ---
_sel_count = len(st.session_state.get("cmp_selected", []))

with st.container(border=True):
    _header_col1, _header_col2 = st.columns([4, 1.2])
    with _header_col1:
        st.markdown("<h3 style='margin-top:0;'>📥 Select Files to be Placed</h3>", unsafe_allow_html=True)
    with _header_col2:
        if _sel_count > 0:
            st.markdown(f"<div style='text-align:right;'><span style='background:#E6F7F3;color:#0D9488;font-size:0.8rem;font-weight:600;padding:0.25rem 0.75rem;border-radius:99px;border:1px solid #A7F3D0;'>{_sel_count} selected</span></div>", unsafe_allow_html=True)

    if not IS_SERVER:
        if st.button("Browse source folder", key="cmp_browse_src"):
            selected = browse_folder()
            if selected:
                st.session_state["cmp_src_input"] = selected
                st.session_state["cmp_trigger_scan"] = True
                st.session_state.cmp_done = False
                st.session_state.cmp_findings = []
                st.session_state.cmp_fixed = False
                st.session_state["cmp_placed_done"] = False
                st.session_state["cmp_rename_map"] = {}
                st.session_state.pop("cmp_place_result", None)
                st.rerun()

    src_folder = st.text_input("Source folder path", placeholder=r"e.g., Z:\qa\study\data\raw\shared\input\cpt", key="cmp_src_input")

    # --- Server fallback: file uploader for source files ---
    if IS_SERVER:
        st.info("💡 Enter a network path accessible from the server, **or** upload files directly below.")
        uploaded_src = st.file_uploader(
            "Upload source Excel/CSV files",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="cmp_upload_src",
        )
        if uploaded_src:
            upload_dir = save_uploaded_files(uploaded_src)
            st.session_state["cmp_src_input"] = upload_dir
            st.session_state["cmp_trigger_scan"] = True
            st.session_state.cmp_done = False
            src_folder = upload_dir

    if st.session_state.pop("cmp_trigger_scan", False) and src_folder:
        try:
            files = scan_folder(src_folder)
            if files:
                st.session_state.cmp_excel_files = files
                st.session_state.cmp_src_folder = src_folder
                st.session_state.cmp_done = False
        except Exception:
            pass

    if st.session_state.cmp_excel_files:
        st.markdown("<div style='height:0.5rem;'></div>", unsafe_allow_html=True)
        col_sa, col_da, _ = st.columns([1.2, 1.2, 4])
        with col_sa:
            if st.button("Select all", key="cmp_sel_all"):
                for f in st.session_state.cmp_excel_files:
                    st.session_state[f"cmp_chk_{f}"] = True
                st.rerun()
        with col_da:
            if st.button("Deselect all", key="cmp_desel_all"):
                for f in st.session_state.cmp_excel_files:
                    st.session_state[f"cmp_chk_{f}"] = False
                st.rerun()

        st.markdown("<div style='height:0.8rem;'></div>", unsafe_allow_html=True)

        selected_files = []
        _file_cols = st.columns(2)
        for idx, f in enumerate(st.session_state.cmp_excel_files):
            with _file_cols[idx % 2]:
                _checked = st.checkbox(f, key=f"cmp_chk_{f}")
                if _checked:
                    selected_files.append(f)
        # Only update cmp_selected from checkboxes BEFORE comparison is done.
        # After cmp_done, the list may have been updated by Fix All (renames).
        if not st.session_state.get("cmp_done"):
            st.session_state.cmp_selected = selected_files

        if selected_files and not st.session_state.get("cmp_done"):
            st.markdown("<div style='height:0.5rem;'></div>", unsafe_allow_html=True)
            if st.button("Done — proceed to compare", type="primary", key="cmp_done_btn", use_container_width=True):
                study = extract_study_name(selected_files)
                st.session_state.cmp_study_name = study
                st.session_state.cmp_done = True
                config = load_study_config()
                if study and study in config:
                    st.session_state["cmp_tpl_input"] = config[study].get("template_path", "")
                    st.session_state["cmp_out_input"] = config[study].get("output_path", "")
                st.rerun()


# --- Section 2: Study, Paths & Compare (appears after Done) ---
if st.session_state.get("cmp_done") and st.session_state.get("cmp_selected"):
    st.write("")
    with st.container(border=True):
        st.markdown("<h3 style='margin-top:0;'>🔄 Study, Paths & Compare</h3>", unsafe_allow_html=True)

        config = load_study_config()
        detected_study = st.session_state.get("cmp_study_name", "")

        # Build dropdown options
        all_studies = list(config.keys())
        if detected_study:
            matching = [s for s in all_studies if detected_study.upper() in s.upper() or s.upper() in detected_study.upper()]
            non_matching = [s for s in all_studies if s not in matching]
            dropdown_options = matching + non_matching + ["+ New Study (browse manually)"]
            default_idx = 0 if matching else len(dropdown_options) - 1
        else:
            dropdown_options = all_studies + ["+ New Study (browse manually)"]
            default_idx = len(dropdown_options) - 1

        # Ensure session state value is valid for current options list
        current_ss_val = st.session_state.get("cmp_study_select")
        if current_ss_val is None or current_ss_val not in dropdown_options:
            st.session_state["cmp_study_select"] = dropdown_options[default_idx]

        is_new_study = True
        _study_col1, _study_col2 = st.columns([2, 1])
        with _study_col1:
            selected_option = st.selectbox(
                "Select Study",
                options=dropdown_options,
                key="cmp_study_select",
            )
        is_new_study = not (selected_option and selected_option != "+ New Study (browse manually)" and selected_option in config)
        if is_new_study:
            with _study_col2:
                study_name = st.text_input("Study Name", value=detected_study, key="cmp_study_name_input",
                                           placeholder="e.g., GZQD")

        if not is_new_study:
            # Existing study — auto-fill, no browse needed
            study_data = config[selected_option]
            folder_path = study_data.get("template_path", "")
            out_path = study_data.get("output_path", "") or folder_path

            # Always set paths from config — this is the source of truth for the selected study
            st.session_state["cmp_tpl_input"] = folder_path
            st.session_state["cmp_out_input"] = out_path

            # Sync advanced inputs to match the selected study (prevents stale values from a previous study)
            st.session_state["cmp_adv_tpl"] = folder_path
            st.session_state["cmp_adv_out"] = out_path

            st.markdown(f"**Template Path:** `{folder_path}`")

            with st.expander("Advanced: use separate template & output folders", expanded=False):
                st.caption("By default, the same folder is used for comparison and placing files.")
                adv_tpl = st.text_input("Template folder (if different)", key="cmp_adv_tpl")
                adv_out = st.text_input("Output folder (if different)", key="cmp_adv_out")
                # Only override if user manually changed the value away from this study's default
                if adv_tpl and adv_tpl != folder_path:
                    st.session_state["cmp_tpl_input"] = adv_tpl
                if adv_out and adv_out != out_path:
                    st.session_state["cmp_out_input"] = adv_out

            # Compare Now button
            if st.button("Compare now", type="primary", key="cmp_compare_btn", use_container_width=True):
                st.session_state.cmp_findings = []
                st.session_state.cmp_fixed = False
                st.session_state["cmp_placed_done"] = False
                st.session_state["cmp_rename_map"] = {}
                st.session_state.pop("cmp_place_result", None)
                st.session_state["_run_compare"] = True
                st.rerun()

        else:
            # New study — browse button only (hidden on server)
            if not IS_SERVER:
                if st.button("Browse folder (template & output)", key="cmp_browse_folder"):
                    sel = browse_folder()
                    if sel:
                        st.session_state["cmp_tpl_input"] = sel
                        st.session_state["cmp_out_input"] = sel
                        st.session_state.cmp_findings = []
                        st.session_state.cmp_fixed = False
                        st.rerun()

            _path_col2, _ = st.columns([2, 1])
            with _path_col2:
                main_folder = st.text_input(
                    "Folder path (used for comparison & placing)",
                    placeholder=r"e.g., Z:\qa\study\data\raw\shared\input\cpt\Template",
                    key="cmp_tpl_input"
                )
            if main_folder and not st.session_state.get("cmp_out_input"):
                st.session_state["cmp_out_input"] = main_folder

            with st.expander("Advanced: Use separate Output folder", expanded=False):
                st.caption("By default, files are placed in the same folder used for comparison.")
                if not IS_SERVER:
                    if st.button("Browse output folder", key="cmp_browse_out"):
                        sel = browse_folder()
                        if sel:
                            st.session_state["cmp_out_input"] = sel
                            st.rerun()
                st.text_input("Output folder (if different)", key="cmp_out_input")

            # Save & Compare button
            if main_folder:
                if st.button("Save & compare", type="primary", key="cmp_save_compare_btn", use_container_width=True):
                    if study_name:
                        out_folder = st.session_state.get("cmp_out_input", main_folder)
                        config[study_name.upper()] = {
                            "template_path": main_folder,
                            "output_path": out_folder,
                        }
                        save_study_config(config)
                    st.session_state.cmp_findings = []
                    st.session_state.cmp_fixed = False
                    st.session_state["cmp_placed_done"] = False
                    st.session_state["cmp_rename_map"] = {}
                    st.session_state.pop("cmp_place_result", None)
                    st.session_state["_run_compare"] = True
                    st.rerun()

    # --- Run comparison (triggered by either button) ---
    if st.session_state.pop("_run_compare", False):
        # Clear old findings immediately
        st.session_state.cmp_findings = []
        st.session_state.cmp_fixed = False

        # Only clear source-file caches (templates rarely change and benefit from caching)
        # detect_column_types cache is keyed by file_path so template entries remain valid
        # We only need to clear if source files may have been modified since last run
        detect_column_types.cache_clear()
        _detect_csv_column_types.cache_clear()
        get_data_sheet.cache_clear()

        tpl_folder = st.session_state.get("cmp_tpl_input", "")
        if not tpl_folder:
            st.warning("Please select a folder path.")
        else:
            try:
                # Show progress immediately
                progress = st.progress(0, text="Scanning folders...")
                src_path = st.session_state.get("cmp_src_folder", "")
                template_files = scan_folder(tpl_folder)
                selected = st.session_state.cmp_selected
                matches = match_files(selected, template_files)

                progress.progress(0.1, text="Comparing files...")

                all_findings = []
                items = list(matches.items())

                # Compare directly from source/template paths (no temp copy needed)
                def _cmp(item):
                    pn, tn = item
                    if tn is None:
                        return [{"File": pn, "Check": "New File (inform DA)",
                                 "Placed Value": pn,
                                 "Template Value": "This file has been placed for first time, please inform DA",
                                 "fix_type": None, "fix_data": None}]
                    return compare_file_pair(
                        os.path.join(src_path, pn),
                        os.path.join(tpl_folder, tn), pn, tn)

                with ThreadPoolExecutor(max_workers=12) as ex:
                    futs = {ex.submit(_cmp, it): it for it in items}
                    for i, fut in enumerate(as_completed(futs), 1):
                        try:
                            res = fut.result()
                            if res:
                                all_findings.extend(res)
                        except Exception:
                            pass
                        progress.progress(0.1 + 0.8 * i / len(items))

                # Check for missing files
                matched_templates = set(v for v in matches.values() if v is not None)
                for tf in template_files:
                    if tf not in matched_templates:
                        all_findings.append({
                            "File": tf, "Check": "Missing File",
                            "Placed Value": "Not Selected- Please select the file if required",
                            "Template Value": f"Please place/select '{tf}' also",
                            "fix_type": None, "fix_data": None,
                        })

                progress.empty()

                st.session_state.cmp_findings = all_findings
                st.session_state.cmp_fixed = False
                if all_findings:
                    st.warning(f"⚠️ Found {len(all_findings)} discrepancy(ies)")
                else:
                    st.success("✅ All files match! No discrepancies.")
                    st.session_state.cmp_fixed = True
            except Exception as e:
                st.error(f"Error: {e}")


# --- Findings (shown inline after compare, part of step 2) ---
if st.session_state.cmp_findings and st.session_state.get("cmp_done"):
    st.write("")
    st.markdown("<h3 style='margin-top:0;'>🔍 Discrepancies Found</h3>", unsafe_allow_html=True)

    html_rows = []
    for f in st.session_state.cmp_findings:
        is_fixable = f.get("fix_type") is not None
        check_name = f["Check"]

        # Determine status badge
        if check_name == "Missing File":
            badge_html = "<span style='background-color:#FEE2E2; color:#EF4444; font-size:0.75rem; font-weight:600; padding:0.2rem 0.65rem; border-radius:99px; border:1px solid #FCA5A5;'>Missing</span>"
        elif check_name.startswith("New File"):
            badge_html = "<span style='background-color:#E0F2FE; color:#0284C7; font-size:0.75rem; font-weight:600; padding:0.2rem 0.65rem; border-radius:99px; border:1px solid #7DD3FC;'>New File</span>"
        elif is_fixable and f.get("_fixed"):
            badge_html = "<span style='background-color:#D1FAE5; color:#047857; font-size:0.75rem; font-weight:600; padding:0.2rem 0.65rem; border-radius:99px; border:1px solid #6EE7B7;'>Fixed</span>"
        elif is_fixable:
            badge_html = "<span style='background-color:#D1FAE5; color:#059669; font-size:0.75rem; font-weight:600; padding:0.2rem 0.65rem; border-radius:99px; border:1px solid #6EE7B7;'>Auto-Fixable</span>"
        else:
            badge_html = "<span style='background-color:#FEF3C7; color:#D97706; font-size:0.75rem; font-weight:600; padding:0.2rem 0.65rem; border-radius:99px; border:1px solid #FCD34D;'>Manual Fix</span>"

        # HTML-escape all dynamic cell values to prevent broken rendering
        esc_file = html_mod.escape(str(f["File"]))
        esc_check = html_mod.escape(str(check_name))
        esc_placed = html_mod.escape(str(f["Placed Value"])).replace("\n", "<br>")
        esc_template = html_mod.escape(str(f["Template Value"])).replace("\n", "<br>")

        html_rows.append(
            f'<tr style="border-bottom: 1px solid #E2E8F0; font-size: 0.88rem;">'
            f'<td style="padding: 12px 14px; color: #0F172A; font-weight: 600;">{esc_file}</td>'
            f'<td style="padding: 12px 14px; color: #334155; font-weight: 500;">{esc_check}</td>'
            f'<td style="padding: 12px 14px; color: #475569; font-family: \'Courier New\', Courier, monospace;">{esc_placed}</td>'
            f'<td style="padding: 12px 14px; color: #475569; font-family: \'Courier New\', Courier, monospace;">{esc_template}</td>'
            f'<td style="padding: 12px 14px; text-align: center;">{badge_html}</td>'
            f'</tr>'
        )
        
    table_body = "".join(html_rows)
    table_html = (
        '<div style="overflow-x: auto; border: 1px solid #E2E8F0; border-radius: 12px; box-shadow: 0 4px 12px 0 rgba(0,0,0,0.03); background: white; margin-bottom: 1rem;">'
        '<table style="width: 100%; border-collapse: collapse; text-align: left;">'
        '<thead>'
        '<tr style="background-color: #F8FAFC; border-bottom: 1.5px solid #E2E8F0; font-family: \'Plus Jakarta Sans\', sans-serif; font-size: 0.85rem; font-weight: 700; color: #475569;">'
        '<th style="padding: 14px 14px;">File Name</th>'
        '<th style="padding: 14px 14px;">Issue Type</th>'
        '<th style="padding: 14px 14px;">Placed Value</th>'
        '<th style="padding: 14px 14px;">Template Value</th>'
        '<th style="padding: 14px 14px; text-align: center;">Status</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{table_body}</tbody>'
        '</table>'
        '</div>'
    )
    st.markdown(table_html, unsafe_allow_html=True)

    fixable = [f for f in st.session_state.cmp_findings if f.get("fix_type")]
    non_fixable = len(st.session_state.cmp_findings) - len(fixable)

    if fixable and not st.session_state.get("cmp_fixed"):
        st.markdown(f"<p style='font-size:0.88rem;color:#374151;'><b>{len(fixable)}</b> can be auto-fixed. <b>{non_fixable}</b> require manual review.</p>", unsafe_allow_html=True)
        if st.button(f"Fix all ({len(fixable)} issues)", type="primary", key="cmp_fix_btn"):
            st.session_state["_run_fix"] = True
            st.rerun()
    elif not fixable and not st.session_state.get("cmp_fixed"):
        st.info("No auto-fixable issues. Manual review required for the above. You can still proceed to place files.")

# --- Run Fix All (triggered by button, runs on rerun for immediate progress bar) ---
if st.session_state.pop("_run_fix", False):
    src_path = st.session_state.get("cmp_src_folder", "")
    if src_path:
        fix_progress = st.progress(0, text="Fixing issues...")
        count, rename_map = apply_fixes(st.session_state.cmp_findings, src_path)
        fix_progress.progress(100, text="Done!")
        fix_progress.empty()

        # Store rename map so placement step uses correct filenames
        st.session_state["cmp_rename_map"] = rename_map

        # Update selected files list to reflect renames
        updated_selected = []
        for f in st.session_state.cmp_selected:
            updated_selected.append(rename_map.get(f, f))
        st.session_state.cmp_selected = updated_selected

        # Clear metadata caches since files were modified
        get_sheet_names_fast.cache_clear()
        find_header_row.cache_clear()
        detect_column_types.cache_clear()
        _detect_csv_column_types.cache_clear()
        get_data_sheet.cache_clear()

        # Keep findings visible for cross-validation (do NOT clear cmp_findings)
        st.session_state.cmp_fixed = True
        st.session_state.cmp_fix_message = f"Fixed {count} discrepancy(ies) in source! Re-run Compare to verify."
        st.rerun()
    else:
        st.error("Source folder not set.")

# Show fix success message
if st.session_state.get("cmp_fix_message"):
    st.success(f"✅ {st.session_state.pop('cmp_fix_message')}")


# --- Step 3: Place Files ---
# Show ONLY when:
#   - Comparison has been run (cmp_findings exists or cmp_fixed is set)
#   - AND either: all fixes applied, OR no auto-fixable issues exist (only Missing/informational)
_findings = st.session_state.get("cmp_findings", [])
_comparison_ran = st.session_state.get("cmp_fixed", False) or len(_findings) > 0
_has_fixable = any(f.get("fix_type") for f in _findings)
_can_place = st.session_state.get("cmp_fixed", False) or (not _has_fixable and _comparison_ran)
if st.session_state.get("cmp_done") and _can_place and _comparison_ran:

    @st.fragment
    def _place_files_fragment():
        """Self-contained fragment: only THIS section re-runs on button click.
        The rest of the page stays untouched — no greying out."""

        with st.container(border=True):
            st.markdown("<h3 style='margin-top:0;'>🚀 Place Files</h3>", unsafe_allow_html=True)
            out_folder = st.session_state.get("cmp_out_input", "")
            if out_folder:
                st.markdown(f"**Target Path:** `{out_folder}`")

            # Show result if available
            _place_result = st.session_state.get("cmp_place_result")
            if _place_result:
                if _place_result.get("error"):
                    st.error(_place_result["error"])
                else:
                    if _place_result.get("placed"):
                        st.success(f"✅ Placed {len(_place_result['placed'])} file(s) to: `{_place_result['out_folder']}`")
                    if _place_result.get("failed"):
                        st.warning(f"⚠️ Could not place {len(_place_result['failed'])} file(s): {', '.join(_place_result['failed'])}")

            # Button
            _already_placed = st.session_state.get("cmp_placed_done", False)
            if not _already_placed:
                if st.button("Place selected files", type="primary", key="cmp_place_btn", use_container_width=True):
                    # Do the work inline within the fragment rerun
                    out_folder = st.session_state.get("cmp_out_input", "")
                    if not out_folder:
                        st.session_state["cmp_place_result"] = {"error": "No output folder set."}
                    else:
                        try:
                            os.makedirs(out_folder, exist_ok=True)
                            src = st.session_state.get("cmp_src_folder", "")
                            selected = st.session_state.cmp_selected
                            rename_map = st.session_state.get("cmp_rename_map", {})

                            # Also gather originally selected files from checkboxes as fallback
                            _original_selected = [f for f in st.session_state.get("cmp_excel_files", [])
                                                  if st.session_state.get(f"cmp_chk_{f}", False)]
                            # Merge: use cmp_selected (may have renames) + any originals not covered
                            _all_to_place = list(selected)
                            for f in _original_selected:
                                # Check if this original file or its renamed version is already in the list
                                renamed_ver = rename_map.get(f, f)
                                if renamed_ver not in _all_to_place and f not in _all_to_place:
                                    _all_to_place.append(renamed_ver)

                            placed, failed = [], []
                            try:
                                src_files_on_disk = os.listdir(src)
                            except Exception:
                                src_files_on_disk = []
                            # Build case-insensitive lookup: lowercase → actual name on disk
                            _src_lookup = {name.lower(): name for name in src_files_on_disk}

                            def _find_and_copy(f):
                                dst_file = os.path.join(out_folder, f)
                                # Case-insensitive file lookup
                                actual_name = _src_lookup.get(f.lower())
                                if actual_name:
                                    shutil.copy2(os.path.join(src, actual_name), dst_file)
                                    return (f, None)
                                # Check rename_map: file may have been renamed by Fix All
                                for old_name, new_name in rename_map.items():
                                    if new_name.lower() == f.lower():
                                        actual = _src_lookup.get(new_name.lower())
                                        if actual:
                                            shutil.copy2(os.path.join(src, actual), dst_file)
                                            return (f, None)
                                    if old_name.lower() == f.lower():
                                        actual = _src_lookup.get(new_name.lower())
                                        if actual:
                                            alt_dst = os.path.join(out_folder, new_name)
                                            shutil.copy2(os.path.join(src, actual), alt_dst)
                                            return (new_name, None)
                                # Fallback: try alternate extensions
                                stem = Path(f).stem.lower()
                                for ext in [".xlsx", ".xls", ".csv"]:
                                    alt_key = stem + ext
                                    actual = _src_lookup.get(alt_key)
                                    if actual:
                                        alt_dst = os.path.join(out_folder, actual)
                                        shutil.copy2(os.path.join(src, actual), alt_dst)
                                        return (actual, None)
                                return (None, f)

                            # Progress bar below the button
                            _prog = st.progress(0, text="Placing files...")
                            with ThreadPoolExecutor(max_workers=8) as ex:
                                futs = {ex.submit(_find_and_copy, f): f for f in _all_to_place}
                                total = len(futs)
                                for i, fut in enumerate(as_completed(futs), 1):
                                    try:
                                        ok, fail = fut.result()
                                        if ok:
                                            placed.append(ok)
                                        elif fail:
                                            failed.append(fail)
                                    except Exception:
                                        failed.append(futs[fut])
                                    _prog.progress(int(i / total * 100), text=f"Placing files... ({i}/{total})")
                            _prog.empty()

                            st.session_state.cmp_placed_files = placed
                            st.session_state["cmp_place_result"] = {"placed": placed, "failed": failed, "out_folder": out_folder}
                            st.session_state["cmp_placed_done"] = True
                        except Exception as e:
                            st.session_state["cmp_place_result"] = {"error": str(e)}
                    st.rerun()

    st.write("")
    _place_files_fragment()

    # --- Server fallback: download placed files as zip ---
    if IS_SERVER and st.session_state.get("cmp_placed_done"):
        _pr = st.session_state.get("cmp_place_result", {})
        _out = _pr.get("out_folder", "")
        _placed = _pr.get("placed", [])
        if _out and _placed and os.path.isdir(_out):
            import io
            import zipfile as _zf
            buf = io.BytesIO()
            with _zf.ZipFile(buf, "w", _zf.ZIP_DEFLATED) as zf:
                for fname in _placed:
                    fpath = os.path.join(_out, fname)
                    if os.path.isfile(fpath):
                        zf.write(fpath, fname)
            buf.seek(0)
            st.download_button(
                "⬇️ Download placed files (.zip)",
                data=buf,
                file_name="placed_files.zip",
                mime="application/zip",
                use_container_width=True,
            )
