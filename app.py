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
    # LEVEL 2: DETAILED WORK ORDERS BY LOGO (WITHOUT INDIVIDUAL SKUs)
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
    # NESTED HIERARCHICAL PRODUCTION SCHEDULE TABLE
    # ---------------------------------------------------------
    nested_schedule_rows = []

    for logo in sorted_logos:
        sum_row = logo_summary[logo_summary['Logo_Clean'] == logo].iloc[0]
        oldest_dt = sum_row['Oldest_Order_Date']
        oldest_dt_str = oldest_dt.strftime('%m/%d/%Y') if pd.notna(oldest_dt) else "-"
        wait_days = int(sum_row['Days_Waiting']) if pd.notna(sum_row['Days_Waiting']) else "-"

        # 1. HEADER ROW FOR LOGO BATCH
        nested_schedule_rows.append({
            'Logo / Work Order #': f"🔷 {logo.upper()}",
            'Decorating Type': sum_row['Decorating_Type'],
            'Customer Group': sum_row['Customer_Groups'],
            'Customer Order Date': oldest_dt_str,
            'Days Waiting': wait_days,
            'Units': int(sum_row['Total_Units']),
            'Special Instructions / Notes': f"BATCH TOTAL: {sum_row['Total_WOs']} Work Order(s)"
        })

        # 2. NESTED WORK ORDER ROWS UNDER THIS LOGO
        sub_wos = wo_detail[wo_detail['Logo_Clean'] == logo].sort_values(
            by='Order_Date', ascending=True, na_position='last'
        )
        for _, wo_row in sub_wos.iterrows():
            wo_dt = wo_row['Order_Date']
            wo_dt_str = wo_dt.strftime('%m/%d/%Y') if pd.notna(wo_dt) else "-"
            wo_wait = int(wo_row['Days_Waiting']) if pd.notna(wo_row['Days_Waiting']) else "-"

            nested_schedule_rows.append({
                'Logo / Work Order #': f"    ↳ WO #{wo_row['WO_Clean']}",
                'Decorating Type': "",
                'Customer Group': wo_row['Customer_Group'],
                'Customer Order Date': wo_dt_str,
                'Days Waiting': wo_wait,
                'Units': int(wo_row['Units']),
                'Special Instructions / Notes': wo_row['Special_Notes']
            })

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
            with st.spinner("Writing to Google Sheet..."):
                success, msg = sync_to_google_sheet(logo_summary_display, nested_schedule_df)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)

    with col_act2:
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            nested_schedule_df.to_excel(writer, sheet_name='Production Schedule', index=False)
            logo_summary_display.to_excel(writer, sheet_name='Logo Batch Summary', index=False)
        output_buffer.seek(0)

        st.download_button(
            label="📥 Download Excel (.xlsx)",
            data=output_buffer,
            file_name=f"Decorating_Production_Schedule_{datetime.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    tab1, tab2 = st.tabs(["📋 Work Order Production Schedule (Nested)", "📊 Logo Batch Summary"])

    with tab1:
        st.caption("Header row shows Logo totals; nested rows show individual Work Orders sorted oldest to newest.")
        st.dataframe(nested_schedule_df, use_container_width=True, hide_index=True)

    with tab2:
        st.caption("Machine setup queue for the decorating company.")
        st.dataframe(logo_summary_display, use_container_width=True, hide_index=True)

else:
    st.info("Upload your Orders export CSV above to generate the decorating production schedule.")
