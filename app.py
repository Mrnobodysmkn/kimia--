import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- Page Configuration ---
st.set_page_config(page_title="داشبورد جامع مالی کیمیا", page_icon="💎", layout="wide")

# --- Custom CSS ---
def set_custom_style():
    st.markdown(
        """
        <style>
            .stApp {
                background-color: #f0f2f6;
                font-family: 'Tahoma', sans-serif;
            }
            h1, h2, h3 {
                color: #1E3A8A;
                text-align: right;
            }
            .stButton > button {
                background-color: #2563EB;
                color: white;
                border-radius: 8px;
                width: 100%;
                font-weight: bold;
            }
            .stButton > button:hover {
                background-color: #1D4ED8;
            }
            .report-box {
                background-color: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                margin-bottom: 20px;
                direction: rtl;
            }
            .divider {
                border-top: 3px solid #1E3A8A;
                margin: 40px 0;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

set_custom_style()

# --- Helper Functions (Shared) ---

def to_float(val):
    """Clean and convert monetary strings to float."""
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).replace(",", "").replace("ریال", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def find_header_row(df, keywords):
    """Dynamically find the header row based on keywords."""
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        if any(keyword in val for val in row_values for keyword in keywords):
            return i
    return 0

# ==========================================
# بخش اول: گزارش مالی فروش (کد قبلی)
# ==========================================

def process_section_1():
    st.markdown("<div class='report-box'>", unsafe_allow_html=True)
    st.header("بخش اول: گزارش مالی فروش")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("فایل اکسل فروش را انتخاب کنید", type=["xlsx"], key="uploader_1")

    if uploaded_file is not None:
        if st.button("پردازش فایل فروش", key="btn_process_1"):
            try:
                # 1. Load Data
                try:
                    df_temp = pd.read_excel(uploaded_file, header=None)
                    header_row = find_header_row(df_temp, ["Date", "تاریخ"])
                    df = pd.read_excel(uploaded_file, header=header_row)
                except Exception as e:
                    st.error(f"خطا در خواندن فایل: {e}")
                    return

                # 2. Standardize Columns
                col_map = {
                    'Deposit': ['واریز', 'بستانکار', 'Deposit', 'مبلغ واریز'],
                    'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
                    'Balance': ['مانده', 'Balance', 'مانده حساب'],
                    'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
                    'Date': ['تاریخ', 'Date', 'تاریخ تراکنش']
                }
                
                for col in df.columns:
                    for std, aliases in col_map.items():
                        if any(alias in str(col) for alias in aliases):
                            df.rename(columns={col: std}, inplace=True)
                
                # Check critical columns
                if 'Date' not in df.columns:
                    st.error("ستون تاریخ (Date) یافت نشد.")
                    return

                # 3. Clean Data
                df = df.dropna(subset=['Date'])
                for col in ['Deposit', 'Withdrawal', 'Balance']:
                    if col in df.columns:
                        df[col] = df[col].apply(to_float)
                
                df['Description'] = df['Description'].fillna("").astype(str)

                # 4. Filters (Old Logic)
                card_mask = df['Description'].str.contains("انتقال از", na=False)
                fee_mask = df['Description'].str.contains("کارمزد", na=False)
                withdraw_mask = df['Description'].str.contains("انتقال وجه", na=False)
                snap_mask = df['Description'].str.contains("مدرن سامانه غذارسان اطلس", na=False)

                # 5. Grouping
                dates = df['Date'].unique()
                rep = pd.DataFrame(index=dates)
                
                rep['card'] = df[card_mask].groupby('Date')['Deposit'].sum()
                rep['fee'] = df[fee_mask].groupby('Date')['Withdrawal'].sum()
                rep['withdraw'] = df[withdraw_mask].groupby('Date')['Withdrawal'].sum()
                rep['snap'] = df[snap_mask].groupby('Date')['Deposit'].sum()
                
                if 'Balance' in df.columns:
                    # Assuming data is sorted, getting the last balance logic from previous code
                    # Note: Original code logic for balance was:
                    # rep['balance'] = df.groupby('Date')['Balance'].last()
                    # We keep it as is.
                    rep['balance'] = df.groupby('Date')['Balance'].last()
                else:
                    rep['balance'] = 0

                rep = rep.fillna(0)

                # 6. Calculations
                rep['sales'] = rep['card'] / 1.1
                rep['tax'] = rep['card'] - rep['sales']
                
                # Format for output
                final_df = rep.reset_index().rename(columns={'index': 'تاریخ'})
                output_columns = {
                    'Date': 'تاریخ',
                    'card': 'کارت به کارت',
                    'sales': 'فروش',
                    'tax': 'مالیات',
                    'fee': 'کارمزد',
                    'withdraw': 'برداشت روز',
                    'balance': 'مانده آخر روز',
                    'snap': 'واریزی اسنپ'
                }
                final_df = final_df.rename(columns=output_columns)
                
                # Keep only relevant columns and order them
                cols_to_keep = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
                final_df = final_df[cols_to_keep]

                st.success("پردازش بخش اول با موفقیت انجام شد.")
                st.dataframe(final_df.head(), use_container_width=True)

                # Download Button
                excel_data = generate_excel(final_df)
                st.download_button(
                    label="📥 دانلود گزارش فروش (بخش اول)",
                    data=excel_data,
                    file_name=f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_btn_1"
                )

            except Exception as e:
                st.error(f"خطا در پردازش بخش اول: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# ==========================================
# بخش دوم: گزارش برداشت و کارمزد (منطق جدید)
# ==========================================

def process_section_2():
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
    st.markdown("<div class='report-box'>", unsafe_allow_html=True)
    st.header("بخش دوم: گزارش برداشت و کارمزد")
    st.markdown("---")

    uploaded_file = st.file_uploader("فایل اکسل صورتحساب را انتخاب کنید", type=["xlsx"], key="uploader_2")

    if uploaded_file is not None:
        if st.button("پردازش صورتحساب", key="btn_process_2"):
            try:
                # 1. Load Data with dynamic header detection
                try:
                    df_temp = pd.read_excel(uploaded_file, header=None)
                    # Find header row containing "تاریخ" or "Date"
                    header_row = find_header_row(df_temp, ["Date", "تاریخ"])
                    df = pd.read_excel(uploaded_file, header=header_row)
                except Exception as e:
                    st.error(f"خطا در خواندن فایل: {e}")
                    return

                # 2. Standardize Columns
                col_map = {
                    'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
                    'Balance': ['مانده', 'Balance', 'مانده حساب'],
                    'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
                    'Date': ['تاریخ', 'Date', 'تاریخ تراکنش'],
                    'Time': ['زمان', 'Time', 'زمان تراکنش']
                }
                
                for col in df.columns:
                    for std, aliases in col_map.items():
                        if any(alias in str(col) for alias in aliases):
                            df.rename(columns={col: std}, inplace=True)
                
                # Validation
                required_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    st.error(f"ستون‌های الزامی زیر یافت نشدند: {missing}")
                    return

                # 3. Clean Data
                df = df.dropna(subset=['Date'])
                for col in ['Withdrawal', 'Balance']:
                    df[col] = df[col].apply(to_float)
                
                df['Description'] = df['Description'].fillna("").astype(str)

                # Ensure Sorting for correct Balance calculation (Last row of the day)
                # If 'Time' column exists, sort by Date and Time
                if 'Time' in df.columns:
                    df = df.sort_values(by=['Date', 'Time'])
                else:
                    # If no time, we assume the file order is chronological
                    pass

                # 4. New Logic Implementation
                
                # Logic A: Daily Withdrawal (برداشت روز)
                # Keywords: "انتقال از", "برداشت از"
                w_keywords = ["انتقال از", "برداشت از"]
                
                # Logic B: Daily Fee (کارمزد روز)
                # Keywords: "برداشت برای کارمزد", "دریافت کارمزد"
                f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

                def calculate_daily_stats(group):
                    # Filter for Withdrawal
                    w_sum = group[
                        group['Description'].apply(lambda x: any(k in x for k in w_keywords))
                    ]['Withdrawal'].sum()

                    # Filter for Fees
                    f_sum = group[
                        group['Description'].apply(lambda x: any(k in x for k in f_keywords))
                    ]['Withdrawal'].sum()

                    # Balance: Last row of the day
                    last_balance = group['Balance'].iloc[-1]

                    return pd.Series({
                        'برداشت روز': w_sum,
                        'کارمزد روز': f_sum,
                        'مانده روز': last_balance
                    })

                # Group by Date and apply logic
                result_df = df.groupby('Date').apply(calculate_daily_stats).reset_index()

                # 5. Display and Download
                st.success("پردازش بخش دوم با موفقیت انجام شد.")
                
                # Rename Date column for consistency
                result_df = result_df.rename(columns={'Date': 'تاریخ'})
                
                st.dataframe(result_df, use_container_width=True)

                # Generate Excel
                excel_data = generate_excel_section2(result_df)
                st.download_button(
                    label="📥 دانلود گزارش صورتحساب (بخش دوم)",
                    data=excel_data,
                    file_name=f"Withdrawal_Fee_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_btn_2"
                )

            except Exception as e:
                st.error(f"خطا در پردازش بخش دوم: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# --- Excel Generation Functions ---

def generate_excel(df):
    """Generates Excel for Section 1 (Old Logic)"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    
    # Headers from dataframe
    headers = list(df.columns)
    ws.append(headers)
    
    # Style Header
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Tahoma')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color="DDEBF7", fill_type="solid")
    
    # Data
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    # Format Numbers
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter != 'A': # Assuming Date is first column
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            
    # Auto-fit columns
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
        
    wb.save(output)
    return output.getvalue()

def generate_excel_section2(df):
    """Generates Excel for Section 2 (New Logic)"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    
    # Columns: تاریخ, برداشت روز, کارمزد روز, مانده روز
    headers = list(df.columns)
    ws.append(headers)
    
    # Style Header
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Tahoma')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color="E2EFDA", fill_type="solid") # Greenish for distinction
    
    # Data
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    # Format Numbers
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter != 'A': 
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            
    # Auto-fit columns
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
        
    wb.save(output)
    return output.getvalue()

def dataframe_to_rows(df, index=False, header=False):
    """Helper to convert DF to rows for openpyxl"""
    from openpyxl.utils.dataframe import dataframe_to_rows as dtr
    return dtr(df, index=index, header=header)


# --- Main Application ---
def main():
    st.title("سیستم پردازش فایل‌های مالی")
    
    # Section 1
    process_section_1()
    
    # Section 2
    process_section_2()

if __name__ == "__main__":
    main()
