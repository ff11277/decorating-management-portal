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

def find_column(df_columns, candidate_names):
    for cand in candidate_names:
        for col in df_columns:
            if col.strip().lower() == cand.strip().lower():
                return col
    return None

# ---------------------------------------------------------
# GOOGLE SHEETS CONNECTION & SYNC HELPERS
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
        st.error("Streamlit Secrets missing [gcp_service_account] or [gsheets]. Please check Settings > Secrets.")
        return None

def write_worksheet_safely(spr, title, df_data):
    try:
        try:
            sheet = spr.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spr.add_worksheet(
                title=title, 
                rows=max(len(df_data) + 50, 100), 
                cols=max(len(df_data.columns) + 5, 15)
            )
        
        clean_df = df_data.fillna("").astype(str)
        data = [clean_df.columns.tolist()] + clean_df.values.tolist()
        
        required_rows = max(len(data) + 20, 100)
        required_cols = max(len(data[0]) + 5, 15)
        if sheet.row_count < required_rows or sheet.col_count < required_cols:
            sheet.resize(
                rows=max(sheet.row_count, required_rows), 
                cols=max(sheet.col_count, required_cols)
            )
            
        sheet.clear()
        try:
            sheet.update(range_name="A1", values=data)
        except TypeError:
            sheet.update("A1", data)
            
        return True, None
    except Exception as e:
        return False, str(e)

def sync_to_google_sheet(summary_df, detailed_df):
    spr = get_gspread_client()
    if not spr:
        return False, "Could not connect to Google Spreadsheet. Check secrets configuration."
    
    # 1. Sync Nested Production Schedule
    ok1, err1 = write_worksheet_safely(spr, "Production Schedule", detailed_df)
    if not ok1:
        return False, f"Failed updating 'Production Schedule': {err1}"

    # 2. Sync Logo Batch Summary
    ok2, err2 = write_worksheet_safely(spr, "Logo Batch Summary", summary_df)
    if not ok2:
        return False, f"Failed updating 'Logo Batch Summary': {err2}"

    return True, "Successfully synced both sheets to Google Sheets!"

# ---------------------------------------------------------
# AUTHENTICATION
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.subheader("🔒 Security Check")
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

# ---------------------------------------------------------
# MAIN APP HEADER & FILE UPLOADER
# ---------------------------------------------------------
st.title("FORCE FITTERS | Decorating Production Schedule")
st.caption("Organized for shop-floor production by Logo and Customer Order Date.")

uploaded_file = st.file_uploader(
    "Upload Force Fitters Orders Export (CSV format):", 
    type=["csv"]
)

df = None
if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file, encoding="utf-8")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding="latin1")
    st.session_state["dec_df"] = df
elif "dec_df" in st.session_state:
    df = st.session_state["dec_df"]

if df is not None:
    df.columns = [c.strip() for c in df.columns]
    
    # Optional status filter if present
    status_col = find_column(df.columns, ["Status", "Work Order Status", "WO Status"])
    if status_col:
        all_statuses = sorted([str(s) for s in df[status_col].dropna().unique()])
        default_statuses = [
            s for s in all_statuses 
            if any(target in s.lower() for target in ["checked in", "decorating"])
        ]
        if not default_statuses:
            default_statuses = all_statuses
            
        selected_statuses = st.multiselect(
            "Filter by Work Order Status:",
            options=all_statuses,
            default=default_statuses
        )
        if selected_statuses:
            df = df[df[status_col].astype(str).isin(selected_statuses)]

    wo_col = find_column(df.columns, ["Work Order Number", "Work Order #", "Work Order", "WO Number", "WO #"])
    order_date_col = find_column(df.columns, ["Date Ordered", "Order Date", "Date Order", "Customer Date Ordered"])
    logo_col = find_column(df.columns, ["Logo", "Logo Name", "Decoration Logo"])
    group_col = find_column(df.columns, ["Group", "Customer Group", "Company", "Customer"])
    qty_col = find_column(df.columns, ["Qty", "Quantity", "Units"])
    sku_col = find_column(df.columns, ["SKU", "Item SKU", "Product SKU", "Style", "Garment"])
    notes_col = find_column(df.columns, ["Notes", "Special Instructions", "Note", "Order Notes", "WO Notes"])

    current_date = pd.Timestamp.now().normalize()

    df['WO_Clean'] = df[wo_col].apply(clean_str).replace("", "No WO Number") if wo_col else "No WO Number"
    df['Logo_Clean'] = df[logo_col].apply(clean_str).replace("", "Unassigned Logo") if logo_col else "Unassigned Logo"
    df['Group_Clean'] = df[group_col].apply(clean_str).replace("", "Unassigned Group") if group_col else "Unassigned Group"
    df['Order_Date_Clean'] = pd.to_datetime(df[order_date_col], errors='coerce') if order_date_col else pd.NaT
    df['Qty_Clean'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0).astype(int) if qty_col else 1
    
    notes_series = df[notes_col].apply(clean_str) if notes_col else pd.Series([""] * len(df))
