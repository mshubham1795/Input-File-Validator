"""Input File Validator — App Entry Point."""
import streamlit as st
import os
from utils import IS_SERVER

st.set_page_config(
    page_title="Input File Validator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS for Premium Design System
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }

    /* ===== SIDEBAR ===== */
    [data-testid="stSidebar"] {
        background: #0B0F17;
        padding-top: 1.5rem;
        border-right: 1px solid #1E293B;
        min-width: 240px !important;
        max-width: 240px !important;
        transform: none !important;
    }
    /* Hide the collapse/close button */
    [data-testid="stSidebar"] button[kind="header"] {
        display: none !important;
    }
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    [data-testid="stSidebar"] * {
        color: #94A3B8 !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-selected="true"] {
        background: linear-gradient(90deg, rgba(13, 148, 136, 0.15) 0%, rgba(13, 148, 136, 0.02) 100%) !important;
        border-left: 4px solid #0D9488 !important;
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-selected="true"] * {
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNavLink"] {
        border-radius: 8px;
        margin: 0.2rem 0.8rem;
        padding: 0.6rem 1rem;
        font-size: 0.9rem;
        transition: all 0.2s ease;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNavLink"]:hover {
        background: rgba(255, 255, 255, 0.03) !important;
        color: #F1F5F9 !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarNavSeparator"] {
        display: none;
    }

    /* ===== MAIN AREA ===== */
    .main {
        background: #F8FAFC;
    }
    .main .block-container {
        padding: 2.5rem 3rem;
        max-width: 1100px;
    }

    /* ===== PAGE TITLE ===== */
    h1 {
        font-size: 2rem !important;
        font-weight: 800 !important;
        color: #0F172A !important;
        margin-bottom: 0.2rem !important;
        letter-spacing: -0.5px;
    }

    /* Caption under title */
    [data-testid="stCaptionContainer"] p {
        color: #64748B !important;
        font-size: 0.95rem !important;
        margin-bottom: 1.5rem;
    }

    /* ===== SECTION HEADERS (h2) ===== */
    h2 {
        font-size: 1.25rem !important;
        font-weight: 700 !important;
        color: #1E293B !important;
        border: none !important;
        padding: 0 !important;
        margin-top: 1.5rem !important;
        margin-bottom: 1rem !important;
    }
    
    h3 {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        color: #334155 !important;
    }

    /* ===== DIVIDERS ===== */
    hr {
        border: none !important;
        border-top: 1px solid #E2E8F0 !important;
        margin: 2rem 0 !important;
    }

    /* ===== BUTTONS ===== */
    /* Primary button - teal gradient */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0D9488 0%, #0F766E 100%) !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.65rem 1.8rem !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        color: white !important;
        box-shadow: 0 4px 12px 0 rgba(13, 148, 136, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #0F766E 0%, #115E59 100%) !important;
        box-shadow: 0 6px 16px 0 rgba(13, 148, 136, 0.35) !important;
        transform: translateY(-1px) !important;
    }
    .stButton > button[kind="primary"]:active {
        transform: translateY(1px) !important;
    }

    /* Secondary buttons - outlined with hover background */
    .stButton > button:not([kind="primary"]) {
        border-radius: 8px !important;
        border: 1px solid #CBD5E1 !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
        color: #334155 !important;
        background: #FFFFFF !important;
        padding: 0.55rem 1.4rem !important;
        box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:not([kind="primary"]):hover {
        border-color: #0D9488 !important;
        color: #0D9488 !important;
        background: #F0FDF4 !important;
        transform: translateY(-1px) !important;
    }

    /* ===== TEXT INPUTS & SELECTBOXES ===== */
    .stTextInput label, .stSelectbox label {
        font-size: 0.75rem !important;
        font-weight: 700 !important;
        color: #475569 !important;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 0.4rem !important;
    }
    .stTextInput > div > div > input {
        border-radius: 8px !important;
        border: 1.5px solid #E2E8F0 !important;
        padding: 0.6rem 1rem !important;
        font-size: 0.92rem !important;
        color: #0F172A !important;
        background: #FFFFFF !important;
        transition: all 0.2s ease !important;
    }
    .stSelectbox > div > div {
        border-radius: 8px !important;
        border: 1.5px solid #E2E8F0 !important;
        font-size: 0.92rem !important;
        color: #0F172A !important;
        background: #FFFFFF !important;
        transition: all 0.2s ease !important;
    }
    .stTextInput > div > div > input:focus, .stSelectbox > div > div:focus-within {
        border-color: #0D9488 !important;
        box-shadow: 0 0 0 3px rgba(13, 148, 136, 0.15) !important;
    }

    /* ===== CHECKBOXES (as individual cards) ===== */
    div[data-testid="stCheckbox"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.2rem !important;
        margin-bottom: 0.6rem !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.02) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div[data-testid="stCheckbox"]:hover {
        border-color: #0D9488 !important;
        box-shadow: 0 6px 15px 0 rgba(13, 148, 136, 0.08) !important;
        transform: translateY(-1px) !important;
    }
    div[data-testid="stCheckbox"] label span[data-testid="stCheckboxLabel"] {
        font-size: 0.92rem !important;
        color: #334155 !important;
        font-weight: 500 !important;
    }

    /* ===== CONTAINER CARD STYLING ===== */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 14px !important;
        padding: 1.8rem !important;
        box-shadow: 0 4px 18px 0 rgba(0, 0, 0, 0.03) !important;
        margin-bottom: 1.5rem !important;
    }

    /* ===== DATAFRAME & DATA EDITOR ===== */
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 12px 0 rgba(0, 0, 0, 0.02) !important;
    }

    /* ===== ALERTS ===== */
    .stAlert {
        border-radius: 10px !important;
        border: 1px solid rgba(0, 0, 0, 0.05) !important;
        font-size: 0.9rem !important;
        font-weight: 500 !important;
    }

    /* ===== PROGRESS BAR ===== */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #0D9488, #2DD4BF) !important;
        border-radius: 8px !important;
    }

    /* ===== EXPANDER ===== */
    .streamlit-expanderHeader {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        background: #F8FAFC !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 0.6rem 1rem !important;
    }

    /* ===== HIDE STREAMLIT BRANDING ===== */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Detect if running on a server (Posit Connect / headless) vs local desktop
_is_server = IS_SERVER

# Sidebar with app title at top and shortcut button below
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding: 0 0.5rem 1.5rem; margin-top: -12px; border-bottom: 1px solid #1E293B; margin-bottom: 1.5rem;">
        <div style="font-size:1.35rem; font-weight:800; background: linear-gradient(135deg, #0D9488 0%, #2DD4BF 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: 0.3px;">⚡ Input File Validator</div>
    </div>
    """, unsafe_allow_html=True)
    # Desktop shortcut button — only shown when running locally on Windows
    if not _is_server:
        if st.button("🖥️ Create Desktop Shortcut", key="create_shortcut_btn", use_container_width=True):
            import subprocess, tempfile
            script_dir = os.path.dirname(os.path.abspath(__file__))
            target = os.path.join(script_dir, "Run_App.bat")
            ps_content = (
                '$desktop = [Environment]::GetFolderPath("Desktop")\n'
                '$shortcutPath = Join-Path $desktop "Input File Validator.lnk"\n'
                '$ws = New-Object -ComObject WScript.Shell\n'
                '$sc = $ws.CreateShortcut($shortcutPath)\n'
                f'$sc.TargetPath = "{target}"\n'
                f'$sc.WorkingDirectory = "{script_dir}"\n'
                '$sc.Description = "Input File Validator"\n'
                '$sc.WindowStyle = 7\n'
                '$sc.Save()\n'
                'Write-Output $shortcutPath\n'
            )
            ps_file = os.path.join(tempfile.gettempdir(), "_create_shortcut.ps1")
            with open(ps_file, "w", encoding="utf-8") as _f:
                _f.write(ps_content)
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_file],
                capture_output=True, text=True
            )
            try:
                os.remove(ps_file)
            except Exception:
                pass
            shortcut_path = result.stdout.strip()
            if shortcut_path and os.path.exists(shortcut_path):
                st.success("✅ Shortcut created on Desktop!")
            else:
                st.error("Could not create shortcut. Please run `create_shortcut.py` manually.")

pg = st.navigation([st.Page("pages/compare.py", title="Input File Validator", icon="⚡")], position="hidden")
pg.run()
