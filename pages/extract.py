"""Extract page — Generate CPT_files.xlsx from selected Excel files."""
import streamlit as st
import pandas as pd
from pathlib import Path
import sys
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# Add parent dir to path so we can import utils
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils import (
    browse_folder, scan_folder, get_sheet_names_fast, find_header_row,
    find_subject_column, find_status_column, classify_filename,
    convert_path_to_unix, IS_SERVER, save_uploaded_files,
)

st.title("CPT metadata")
st.caption("Generate structured metadata from Excel files for CPT tracking.")

# --- Session State ---
if "ext_excel_files" not in st.session_state:
    st.session_state.ext_excel_files = []
if "ext_metadata_df" not in st.session_state:
    st.session_state.ext_metadata_df = None
if "ext_sheet_cache" not in st.session_state:
    st.session_state.ext_sheet_cache = {}


def generate_metadata(folder_path, files_and_sheets):
    display_path = convert_path_to_unix(folder_path)
    rows = []
    for filename, sheet_name in files_and_sheets.items():
        file_path = str(Path(folder_path) / filename)
        header_row, headers = find_header_row(file_path, sheet_name)
        subject_var = find_subject_column(headers)
        status_var = find_status_column(headers)
        files_name, output_dataset, output_var_name = classify_filename(filename)
        rows.append({
            "Files": files_name,
            "Path": display_path,
            "Filename_with_extension": filename,
            "sheetname": sheet_name,
            "range": f"A{header_row}",
            "subject_var": subject_var,
            "status_var": status_var,
            "output_dataset": output_dataset,
            "output_var_name": output_var_name,
        })
    return pd.DataFrame(rows)


def save_output(df, output_folder):
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)
    file_path = output_path / "CPT_files.xlsx"
    df.to_excel(str(file_path), index=False, engine="openpyxl")
    return str(file_path)


# --- UI: Section 1 - Folder Input ---
with st.container(border=True):
    st.markdown("<h3 style='margin-top:0;'>📥 Select Source Folder</h3>", unsafe_allow_html=True)

    col_browse, col_scan = st.columns([1, 1])
    with col_browse:
        if not IS_SERVER:
            if st.button("Browse source folder", key="ext_browse_src"):
                selected = browse_folder()
                if selected:
                    st.session_state["ext_folder_input"] = selected
                    st.session_state["ext_trigger_scan"] = True
                    st.rerun()
    with col_scan:
        scan_clicked = st.button("Scan folder", key="ext_scan_btn")

    folder_path = st.text_input(
        "Folder path containing Excel files",
        placeholder=r"e.g., C:\Users\Documents\ExcelFiles or \\server\share\folder",
        key="ext_folder_input",
    )

    # --- Server fallback: file uploader ---
    if IS_SERVER:
        st.info("💡 Enter a network path accessible from the server, **or** upload files directly below.")
        uploaded_src_files = st.file_uploader(
            "Upload Excel/CSV files",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="ext_upload_src",
        )
        if uploaded_src_files:
            upload_dir = save_uploaded_files(uploaded_src_files)
            st.session_state["ext_folder_input"] = upload_dir
            st.session_state["ext_trigger_scan"] = True
            folder_path = upload_dir

should_scan = scan_clicked or st.session_state.pop("ext_trigger_scan", False)

if should_scan and folder_path:
    try:
        files = scan_folder(folder_path)
        if files:
            st.session_state.ext_excel_files = files
            st.session_state.ext_folder_path = folder_path
            st.session_state.ext_sheet_cache = {}
            progress_bar = st.progress(0, text="Reading sheet names in parallel...")
            
            # Read sheet names in parallel
            file_paths = [str(Path(folder_path) / f) for f in files]
            sheet_results = {}
            with ThreadPoolExecutor(max_workers=8) as executor:
                future_to_file = {executor.submit(get_sheet_names_fast, fp): f for fp, f in zip(file_paths, files)}
                for i, future in enumerate(as_completed(future_to_file), 1):
                    f = future_to_file[future]
                    try:
                        sheet_results[f] = future.result()
                    except Exception as e:
                        sheet_results[f] = [f"ERROR: {str(e)[:50]}"]
                    progress_bar.progress(i / len(files), text=f"Read sheet names: {f}")
            
            st.session_state.ext_sheet_cache = sheet_results
            progress_bar.empty()
            st.success(f"✅ Found {len(files)} Excel file(s)")
        else:
            st.session_state.ext_excel_files = []
            st.warning("No Excel files (.xlsx, .xls) found in this folder.")
    except (FileNotFoundError, NotADirectoryError) as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Error scanning folder: {e}")
elif should_scan and not folder_path:
    st.warning("Please enter a folder path first.")

# --- UI: Section 2 - File Selection ---
if st.session_state.ext_excel_files:
    st.write("")
    with st.container(border=True):
        st.markdown("<h3 style='margin-top:0;'>📁 Select Files & Sheets</h3>", unsafe_allow_html=True)

        col_sa, col_da, _ = st.columns([1.2, 1.2, 4])
        with col_sa:
            if st.button("Select all", key="ext_sel_all"):
                for f in st.session_state.ext_excel_files:
                    st.session_state[f"ext_chk_{f}"] = True
                st.rerun()
        with col_da:
            if st.button("Deselect all", key="ext_desel_all"):
                for f in st.session_state.ext_excel_files:
                    st.session_state[f"ext_chk_{f}"] = False
                st.rerun()

        st.markdown("<div style='height:0.8rem;'></div>", unsafe_allow_html=True)

        selected_files_and_sheets = {}
        for filename in st.session_state.ext_excel_files:
            col_chk, col_sheet = st.columns([3, 2])
            with col_chk:
                is_selected = st.checkbox(
                    filename, key=f"ext_chk_{filename}",
                    value=st.session_state.get(f"ext_chk_{filename}", True),
                )
            with col_sheet:
                if is_selected:
                    sheets = st.session_state.ext_sheet_cache.get(filename, ["Sheet1"])
                    if sheets and not sheets[0].startswith("ERROR"):
                        selected_sheet = st.selectbox(
                            "Sheet", options=sheets,
                            key=f"ext_sheet_{filename}", label_visibility="collapsed",
                        )
                        selected_files_and_sheets[filename] = selected_sheet
                    else:
                        st.warning(f"⚠️ {sheets[0] if sheets else 'Cannot read file'}")

    # --- UI: Section 3 - Generate ---
    if selected_files_and_sheets:
        st.write("")
        if st.button("Generate metadata", type="primary", key="ext_gen_btn", use_container_width=True):
            with st.spinner("Extracting metadata..."):
                folder = st.session_state.get("ext_folder_path", folder_path)
                df = generate_metadata(folder, selected_files_and_sheets)
                st.session_state.ext_metadata_df = df
            st.success("✅ Metadata generated! Edit the table below.")

    # --- UI: Section 4 - Editable Table ---
    if st.session_state.ext_metadata_df is not None:
        st.write("")
        st.markdown("### 📝 Edit Metadata Table")
        edited_df = st.data_editor(
            st.session_state.ext_metadata_df,
            num_rows="dynamic", key="ext_editor",
            use_container_width=True,
        )
        st.session_state.ext_metadata_df = edited_df

        # --- UI: Section 5 - Save ---
        st.write("")
        with st.container(border=True):
            st.markdown("<h3 style='margin-top:0;'>💾 Save Output</h3>", unsafe_allow_html=True)

            if not IS_SERVER:
                if st.button("Browse output folder", key="ext_browse_out_btn"):
                    selected_out = browse_folder()
                    if selected_out:
                        st.session_state["ext_output_path"] = selected_out
                        st.rerun()

            output_path = st.text_input(
                "Output folder path",
                placeholder=r"e.g., C:\Users\Documents\Output",
                key="ext_output_path",
            )

            if st.button("Save CPT_files.xlsx", type="primary", key="ext_save_btn", use_container_width=True):
                if not output_path:
                    st.warning("Please enter an output folder path.")
                elif edited_df is None or edited_df.empty:
                    st.warning("No data to save.")
                else:
                    try:
                        saved_path = save_output(edited_df, output_path)
                        st.success(f"✅ Saved: `{saved_path}`")
                    except PermissionError:
                        st.error("Permission denied.")
                    except Exception as e:
                        st.error(f"Error: {e}")

            # --- Server fallback: download button ---
            if IS_SERVER and edited_df is not None and not edited_df.empty:
                import io
                buf = io.BytesIO()
                edited_df.to_excel(buf, index=False, engine="openpyxl")
                buf.seek(0)
                st.download_button(
                    "⬇️ Download CPT_files.xlsx",
                    data=buf,
                    file_name="CPT_files.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
else:
    st.write("")
    st.info("👆 Enter a folder path and click 'Scan Folder' to get started.")
