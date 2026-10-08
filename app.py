# app.py
import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io
import gspread
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------
# CONFIGURATION & PAGE SETUP
# ---------------------------------------------------------
APP_PASSWORD = "11277"

st.set_page_config(
    page_title="Force Fitters - Decorating Production Schedule", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s.lower() in ["nan", "none", "<na>"]:
        return ""
    return s

def detect_decorating_type(logo_text, notes_text="", sku_text=""):
    combined = f"{logo_text} {notes_text} {sku_text}".upper()
    types = []
    if re.search(r'\b(EMB|EMBROIDERY|EMBROIDERED)\b', combined):
        types.append('Embroidery')
    if re.search(r'\b(HT|HEAT TRANSFER|HEAT-TRANSFER|HEATSEAL|HEAT SEAL|TRANSFER)\b', combined):
        types.append('Heat Transfer')
    if re.search(r'\b(SP|SCREEN PRINT|SCREENPRINT|SCREEN)\b', combined):
        types.append('Screen Print')
    if re.search(r'\b(PTCH|PATCH)\b', combined):
        types.append('Patch')
    return ", ".join(types) if types else "Embroidery"

# ---------------------------------------------------------
# BACKGROUND GOOGLE SHEETS CONNECTION
# ---------------------------------------------------------
@st.cache_resource
def get_gspread_client():
    if "gcp_service_account" in st.secrets and "gsheets" in st.secrets:
        try:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=scopes
            )
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_url(st.secrets["gsheets"]["spreadsheet_url"])
            return spreadsheet
        except Exception as e:
            st.error(f"Google Sheets Connection Error: {e}")
            return None
    else:
        st.error("Streamlit Secrets missing [gcp_service_account] or [gsheets].")
        return None

def sync_to_google_sheet(summary_df, detailed_df):
    spr = get_gspread_client()
    if not spr:
        return False, "Could not connect to Google Spreadsheet."
    
    try:
        # Tab 1: Detailed Production Schedule
        try:
            sheet_details = spr.worksheet("Production Schedule")
        except gspread.exceptions.WorksheetNotFound:
            sheet_details = spr.add_worksheet(title="Production Schedule", rows=len(detailed_df) + 100, cols=12)
        
        clean_detailed = detailed_df.fillna("").astype(str)
        data_detailed = [clean_detailed.columns.tolist()] + clean_detailed.values.tolist()
        sheet_details.clear()
        sheet_details.update("A1", data_detailed)

        # Tab 2: Logo Batch Summary
        try:
            sheet_summary = spr.worksheet("Logo Batch Summary")
        except gspread.exceptions.WorksheetNotFound:
            sheet_summary = spr.add_worksheet(title="Logo Batch Summary", rows=len(summary_df) + 100, cols=10)
            
        clean_summary = summary_df.fillna("").astype(str)
        data_summary = [clean_summary.columns.tolist()] + clean_summary.values.tolist()
        sheet_summary.clear()
        sheet_summary.update("A1", data_summary)

        return True, "Successfully synced both sheets to Google Sheets!"
    except Exception as e:
        return False, f"Sync error: {e}"

# ---------------------------------------------------------
# CUSTOM INJECTED CSS
# ---------------------------------------------------------
st.markdown("""
<style>
    :root {
        --background-color: #F3F4F6 !important;
        --secondary-background-color: #FFFFFF !important;
        --text-color: #111827 !important;
    }
    header[data-testid="stHeader"], [data-testid="stHeader"] {
        display: none !important;
    }
    .block-container {
        padding-top: 0rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 100% !important;
    }
    .stApp {
        background-color: #F4F5F7 !important;
        color: #111827 !important;
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif !important;
    }
    p, span, label, h1, h2, h3, h4, h5, h6, div, .stMarkdown, .stCaption, small, button {
        color: #111827 !important;
    }
    .ff-navbar {
        background-color: #111111 !important;
        padding: 14px 28px !important;
        display: flex !important;
        align-items: center !important;
        border-bottom: 1px solid #222222 !important;
        margin-left: -2rem !important;
        margin-right: -2rem !important;
        margin-top: 0rem !important;
        margin-bottom: 1.2rem !important;
    }
    .ff-brand {
        color: #FFFFFF !important;
        font-size: 1.15rem !important;
        font-weight: 800 !important;
        letter-spacing: 0.5px !important;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E5E7EB !important;
        border-radius: 6px !important;
        padding: 10px 14px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div[data-testid="stMetricLabel"] p {
        color: #4B5563 !important;
        font-weight: 700 !important;
        font-size: 0.75rem !important;
        text-transform: uppercase;
    }
    div[data-testid="stMetricValue"] div {
        color: #111827 !important;
        font-weight: 800 !important;
    }
    div[data-testid="stDataFrame"], div[data-testid="stTable"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E5E7EB !important;
        border-radius: 6px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
    }
    .stButton > button, .stDownloadButton > button {
        background-color: #111827 !important;
        color: #FFFFFF !important;
        border: 1px solid #111827 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        padding: 6px 16px !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #374151 !important;
        color: #FFFFFF !important;
    }
</style>

<div class="ff-navbar">
    <div class="ff-brand">
        FORCE FITTERS <span style="color: #6B7280; margin: 0 8px;">|</span> Decorating Production Schedule
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# AUTHENTICATION
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 Security Check")
    with st.form("login_form"):
        user_input = st.text_input("Enter Access Key:", type="password")
        login_submitted = st.form_submit_button("Login")
        if login_submitted:
            if user_input == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect access key. Please try again.")
    st.stop()
