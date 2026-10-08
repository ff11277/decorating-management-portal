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
        # 1. Sync Production Schedule (Hierarchical with Row Grouping)
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

        # Clear any existing row dimension groups on Production Schedule
        for _ in range(5):
            try:
                spr.batch_update({
                    "requests": [{
                        "deleteDimensionGroup": {
                            "range": {
                                "sheetId": sheet_details.id,
                                "dimension": "ROWS",
                                "startIndex": 0,
                                "endIndex": sheet_details.row_count
                            }
                        }
                    }]
                })
            except Exception:
                break

        # Apply row grouping (+ / - toggles) above the nested rows
        group_requests = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_details.id,
                        "rowGroupControlAfter": False
                    },
                    "fields": "rowGroupControlAfter"
                }
            }
        ]
        for start_idx, end_idx in row_group_ranges:
            group_requests.append({
                "addDimensionGroup": {
                    "range": {
                        "sheetId": sheet_details.id,
                        "dimension": "ROWS",
                        "startIndex": start_idx,
                        "endIndex": end_idx
                    }
                }
            })

        if len(group_requests) > 1:
            spr.batch_update({"requests": group_requests})

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
    sku_series = df[sku_col].apply(clean_str) if sku_col else pd.Series([""] * len(df))

    df['Dec_Type'] = [
        detect_decorating_type(l, n, s) 
        for l, n, s in zip(df['Logo_Clean'], notes_series, sku_series)
    ]

    # ---------------------------------------------------------
