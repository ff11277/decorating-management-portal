import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io
import gspread
from google.oauth2.service_account import Credentials
import openpyxl

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

def sync_to_google_sheet(summary_df, detailed_df, row_group_ranges):
    spr = get_gspread_client()
    if not spr:
        return False, "Could not connect to Google Spreadsheet. Check secrets configuration."
    
    try:
        # 1. Sync Detailed Production Schedule
        try:
            sheet_details = spr.worksheet("Production Schedule")
        except gspread.exceptions.WorksheetNotFound:
            sheet_details = spr.add_worksheet(
                title="Production Schedule", 
                rows=max(len(detailed_df) + 100, 100), 
                cols=max(len(detailed_df.columns) + 5, 15)
            )

        clean_detailed = detailed_df.fillna("").astype(str)
        data_detailed = [clean_detailed.columns.tolist()] + clean_detailed.values.tolist()
        
        req_rows = max(len(data_detailed) + 50, 100)
        req_cols = max(len(data_detailed[0]) + 5, 15)
        if sheet_details.row_count < req_rows or sheet_details.col_count < req_cols:
            sheet_details.resize(rows=max(sheet_details.row_count, req_rows), cols=max(sheet_details.col_count, req_cols))

        sheet_details.clear()
        try:
            sheet_details.update(range_name="A1", values=data_detailed)
        except TypeError:
            sheet_details.update("A1", data_detailed)

        # 2. Sync Logo Batch Summary
        try:
            sheet_summary = spr.worksheet("Logo Batch Summary")
        except gspread.exceptions.WorksheetNotFound:
            sheet_summary = spr.add_worksheet(
                title="Logo Batch Summary", 
                rows=max(len(summary_df) + 100, 100), 
                cols=max(len(summary_df.columns) + 5, 15)
            )

        clean_summary = summary_df.fillna("").astype(str)
        data_summary = [clean_summary.columns.tolist()] + clean_summary.values.tolist()

        req_rows_sum = max(len(data_summary) + 50, 100)
        req_cols_sum = max(len(data_summary[0]) + 5, 15)
        if sheet_summary.row_count < req_rows_sum or sheet_summary.col_count < req_cols_sum:
            sheet_summary.resize(rows=max(sheet_summary.row_count, req_rows_sum), cols=max(sheet_summary.col_count, req_cols_sum))

        sheet_summary.clear()
        try:
            sheet_summary.update(range_name="A1", values=data_summary)
        except TypeError:
            sheet_summary.update("A1", data_summary)

        # 3. Apply Google Sheets Collapsible Row Grouping
        try:
            meta = spr.fetch_sheet_metadata()
            existing_row_groups = []
            for s in meta.get("sheets", []):
                if s.get("properties", {}).get("sheetId") == sheet_details.id:
                    existing_row_groups = s.get("rowGroups", [])
                    break
            
            if existing_row_groups:
                delete_reqs = [
                    {"deleteDimensionGroup": {"range": g["range"]}}
                    for g in reversed(existing_row_groups)
                ]
                spr.batch_update({"requests": delete_reqs})

            # Configure [+] toggle to sit on top (on the Logo Header row)
            spr.batch_update({
                "requests": [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_details.id,
                            "gridProperties": {
                                "rowGroupControlAfter": False
                            }
                        },
                        "fields": "gridProperties.rowGroupControlAfter"
                    }
                }]
            })

            # Add row groups for each Logo's nested Work Orders
            if row_group_ranges:
                add_group_requests = [
                    {
                        "addDimensionGroup": {
                            "range": {
                                "sheetId": sheet_details.id,
                                "dimension": "ROWS",
                                "startIndex": start_idx,
                                "endIndex": end_idx
                            }
                        }
                    }
                    for start_idx, end_idx in row_group_ranges
                ]
                spr.batch_update({"requests": add_group_requests})
        except Exception:
            pass

        return True, "Successfully synced both sheets to Google Sheets with collapsible row groups!"
    except Exception as e:
        return False, f"Sync error: {e}"

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
    
    # Auto-filter for Checked In or Decorating if status column exists
    status_col = find_column(df.columns, ["Status", "Work Order Status", "WO Status"])
    if status_col:
        mask = df[status_col].astype(str).str.lower().str.contains("checked in|decorating", regex=True, na=False)
        if mask.any():
            df = df[mask]

    wo_col = find_column(df.columns, ["Work Order Number", "Work Order #", "Work Order", "WO Number", "WO #"])
    wo_date_col = find_column(df.columns, ["Date WO Created", "WO Date", "Date Created", "Created Date", "Work Order Date", "WO Created Date"])
    order_date_col = find_column(df.columns, ["Date Ordered", "Order Date", "Date Order", "Customer Date Ordered"])
    logo_col = find_column(df.columns, ["Logo", "Logo Name", "Decoration Logo"])
    group_col = find_column(df.columns, ["Group", "Customer Group", "Company", "Customer"])
    qty_col = find_column(df.
