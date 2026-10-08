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
    
    # Auto-filter for Checked In or Decorating if status column exists
    status_col = find_column(df.columns, ["Status", "Work Order Status", "WO Status"])
    if status_col:
        mask = df[status_col].astype(str).str.lower().str.contains("checked in|decorating", regex=True, na=False)
        if mask.any():
            df = df[mask]

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
    # LEVEL 1: LOGO BATCH SUMMARY (Sorted Oldest to Newest by Customer Order Date)
    # ---------------------------------------------------------
    logo_summary = df.groupby('Logo_Clean', dropna=False).agg(
        Oldest_Order_Date=('Order_Date_Clean', 'min'),
        Decorating_Type=('Dec_Type', 'first'),
        Total_WOs=('WO_Clean', 'nunique'),
        Total_Units=('Qty_Clean', 'sum'),
        Customer_Groups=('Group_Clean', lambda s: " | ".join(sorted(set(s))))
    ).reset_index()

    logo_summary['Days_Waiting'] = (current_date - logo_summary['Oldest_Order_Date']).dt.days
    logo_summary = logo_summary.sort_values(by='Oldest_Order_Date', ascending=True, na_position='last')
    sorted_logos = logo_summary['Logo_Clean'].tolist()

    # ---------------------------------------------------------
    # LEVEL 2: DETAILED WORK ORDERS BY LOGO
    # ---------------------------------------------------------
    def combine_unique(series):
        vals = [clean_str(x) for x in series.dropna().unique() if clean_str(x)]
        return " | ".join(vals) if vals else "-"

    agg_dict = {
        'Order_Date': ('Order_Date_Clean', 'min'),
        'Customer_Group': ('Group_Clean', lambda s: " | ".join(sorted(set(s)))),
        'Decorating_Type': ('Dec_Type', 'first'),
        'Units': ('Qty_Clean', 'sum')
    }
    if notes_col:
        agg_dict['Special_Notes'] = (notes_col, combine_unique)
    else:
        df['_notes_placeholder'] = "-"
        agg_dict['Special_Notes'] = ('_notes_placeholder', 'first')

    wo_detail = df.groupby(['Logo_Clean', 'WO_Clean'], dropna=False).agg(**agg_dict).reset_index()
    wo_detail['Days_Waiting'] = (current_date - wo_detail['Order_Date']).dt.days

    # ---------------------------------------------------------
    # BUILD HIERARCHICAL SCHEDULE TABLE (SEPARATE LOGO & WO# COLUMNS)
    # ---------------------------------------------------------
    nested_schedule_rows = []
    row_group_ranges = []
    excel_sub_row_indices = []

    current_sheet_row_idx = 1 # Index 0 is the table header row (Row 1 in Sheets)

    for logo in sorted_logos:
        sum_row = logo_summary[logo_summary['Logo_Clean'] == logo].iloc[0]
        oldest_dt = sum_row['Oldest_Order_Date']
        oldest_dt_str = oldest_dt.strftime('%m/%d/%Y') if pd.notna(oldest_dt) else "-"
        wait_days = int(sum_row['Days_Waiting']) if pd.notna(sum_row['Days_Waiting']) else "-"

        # 1. LOGO HEADER ROW (Logo in col 1, blank in col 2)
        nested_schedule_rows.append({
            'Logo': logo.upper(),
            'Work Order #': "",
            'Decorating Type': sum_row['Decorating_Type'],
            'Customer Group': sum_row['Customer_Groups'],
            'Customer Order Date': oldest_dt_str,
            'Days Waiting': wait_days,
            'Units': int(sum_row['Total_Units']),
            'Special Instructions / Notes': f"BATCH TOTAL: {sum_row['Total_WOs']} Work Order(s)"
        })
        current_sheet_row_idx += 1

        # 2. NESTED WORK ORDER ROWS (Blank in col 1, WO# in col 2)
        sub_wos = wo_detail[wo_detail['Logo_Clean'] == logo].sort_values(
            by='Order_Date', ascending=True, na_position='last'
        )
        
        group_start_idx = current_sheet_row_idx
        for _, wo_row in sub_wos.iterrows():
            wo_dt = wo_row['Order_Date']
            wo_dt_str = wo_dt.strftime('%m/%d/%Y') if pd.notna(wo_dt) else "-"
            wo_wait = int(wo_row['Days_Waiting']) if pd.notna(wo_row['Days_Waiting']) else "-"

            nested_schedule_rows.append({
                'Logo': "",
                'Work Order #': wo_row['WO_Clean'],
                'Decorating Type': "",
                'Customer Group': wo_row['Customer_Group'],
                'Customer Order Date': wo_dt_str,
                'Days Waiting': wo_wait,
                'Units': int(wo_row['Units']),
                'Special Instructions / Notes': wo_row['Special_Notes']
            })
            excel_sub_row_indices.append(current_sheet_row_idx + 1)
            current_sheet_row_idx += 1
            
        group_end_idx = current_sheet_row_idx
        if group_end_idx > group_start_idx:
            row_group_ranges.append((group_start_idx, group_end_idx))

    nested_schedule_df = pd.DataFrame(nested_schedule_rows)

    # Clean Logo Summary for display and sync
    logo_summary_display = logo_summary.copy()
    logo_summary_display['Oldest_Order_Date'] = logo_summary_display['Oldest_Order_Date'].apply(
        lambda d: d.strftime('%m/%d/%Y') if pd.notna(d) else "-"
    )
    logo_summary_display['Days_Waiting'] = logo_summary_display['Days_Waiting'].apply(
        lambda x: int(x) if pd.notna(x) else "-"
    )

    summary_cols = ['Logo_Clean', 'Decorating_Type', 'Oldest_Order_Date', 'Days_Waiting', 'Total_WOs', 'Total_Units', 'Customer_Groups']
    logo_summary_display = logo_summary_display[summary_cols].rename(columns={
        'Logo_Clean': 'Logo',
        'Decorating_Type': 'Decorating Type',
        'Oldest_Order_Date': 'Oldest Order Date',
        'Days_Waiting': 'Oldest Wait (Days)',
        'Total_WOs': 'Open WOs',
        'Total_Units': 'Total Units',
        'Customer_Groups': 'Customer Group(s)'
    })

    # Metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Garment Units", f"{df['Qty_Clean'].sum():,}")
    m2.metric("Total Work Orders", f"{df['WO_Clean'].nunique():,}")
    m3.metric("Distinct Logo Setups", f"{len(logo_summary):,}")
    valid_waits = logo_summary['Days_Waiting'].dropna()
    oldest_wait = f"{int(valid_waits.max())} Days" if not valid_waits.empty else "N/A"
    m4.metric("Oldest Customer Wait", oldest_wait)

    st.markdown("---")

    col_act1, col_act2, _ = st.columns([1.5, 1.5, 3])

    with col_act1:
        if st.button("📤 Sync to Decorator Google Sheet"):
            with st.spinner("Writing to Google Sheet & setting up collapsible row groups..."):
                success, msg = sync_to_google_sheet(logo_summary_display, nested_schedule_df, row_group_ranges)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)

    with col_act2:
        output_buffer = io.BytesIO()
        wb = openpyxl.Workbook()
        
        # Sheet 1: Production Schedule (with collapsible outline grouping)
        ws1 = wb.active
        ws1.title = 'Production Schedule'
        ws1.sheet_properties.outlinePr.summaryBelow = False
        ws1.append(nested_schedule_df.columns.tolist())
        for r_val in nested_schedule_df.values.tolist():
            ws1.append(r_val)
        for sub_r in excel_sub_row_indices:
            ws1.row_dimensions[sub_r].outlineLevel = 1

        # Sheet 2: Logo Batch Summary
        ws2 = wb.create_sheet(title='Logo Batch Summary')
        ws2.append(logo_summary_display.columns.tolist())
        for r_val in logo_summary_display.values.tolist():
            ws2.append(r_val)

        wb.save(output_buffer)
        output_buffer.seek(0)

        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=output_buffer,
            file_name=f"Decorating_Production_Schedule_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    tab1, tab2 = st.tabs(["📋 Work Order Production Schedule (Nested)", "📊 Logo Batch Summary"])

    with tab1:
        st.caption("Separate Logo and Work Order # columns. On Google Sheets / Excel, click [+] on the left margin to expand sub-rows.")
        st.dataframe(nested_schedule_df, hide_index=True)

    with tab2:
        st.caption("Machine setup queue for the decorating company.")
        st.dataframe(logo_summary_display, hide_index=True)

else:
    st.info("Upload your Orders export CSV above to generate the decorating production schedule.")
