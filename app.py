import os
import re
import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(page_title="Advanced OD Interest & Statement Analyzer", page_icon="🏦", layout="wide")

def parse_bank_statement_pdf_perfect(pdf_file):
    all_rows = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            
            lines = {}
            for w in words:
                top_key = round(w['top'], 1)
                matched_line = None
                for k in lines.keys():
                    if abs(k - top_key) < 3.5:
                        matched_line = k
                        break
                if matched_line is not None:
                    lines[matched_line].append(w)
                else:
                    lines[top_key] = [w]
            
            sorted_line_keys = sorted(lines.keys())
            
            for k in sorted_line_keys:
                line_words = sorted(lines[k], key=lambda x: x['x0'])
                line_text = " ".join([w['text'] for w in line_words])
                
                if "Date" in line_text and "Closing Balance" in line_text:
                    continue
                if "Withdrawal" in line_text or "Deposit" in line_text or "Narration" in line_text:
                    continue
                    
                first_word = line_words[0]['text'] if line_words else ""
                if not re.match(r'^\d{2}/\d{2}/\d{2}$', first_word):
                    continue
                
                date_str = first_word
                narration_words = []
                ref_words = []
                value_dt_str = ""
                withdrawal_str = ""
                deposit_str = ""
                balance_str = ""
                
                for w in line_words:
                    x0 = w['x0']
                    txt = w['text']
                    
                    if x0 < 60:
                        continue
                    elif 60 <= x0 < 220:
                        narration_words.append(txt)
                    elif 220 <= x0 < 340:
                        ref_words.append(txt)
                    elif 340 <= x0 < 400:
                        if re.match(r'^\d{2}/\d{2}/\d{2}$', txt):
                            value_dt_str = txt
                    elif 400 <= x0 < 470:
                        withdrawal_str = txt
                    elif 470 <= x0 < 540:
                        deposit_str = txt
                    elif x0 >= 540:
                        balance_str = txt
                
                all_rows.append({
                    "Date": date_str,
                    "Narration": " ".join(narration_words),
                    "Chq/Ref.No.": " ".join(ref_words),
                    "Value Dt": value_dt_str if value_dt_str else date_str,
                    "Withdrawal Amt": clean_number(withdrawal_str),
                    "Deposit Amt": clean_number(deposit_str),
                    "Closing Balance": clean_number(balance_str)
                })

    df = pd.DataFrame(all_rows)
    return df

def clean_number(val):
    if not val or str(val).strip() == "" or str(val).lower() == "none":
        return 0.0
    val_str = str(val).replace(",", "").strip()
    match = re.search(r'[-+]?\d*\.\d+|\d+', val_str)
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return 0.0
    return 0.0

def calculate_od_interest(df, od_limit, annual_interest_rate, start_date=None, end_date=None):
    df['Parsed_Date'] = pd.to_datetime(df['Value Dt'], format='%d/%m/%y', errors='coerce')
    df['Parsed_Date'] = df['Parsed_Date'].fillna(pd.to_datetime(df['Date'], format='%d/%m/%y', errors='coerce'))
    
    df = df.dropna(subset=['Parsed_Date']).sort_values(by='Parsed_Date').reset_index(drop=True)
    
    # Date Range Filtering
    if start_date and end_date:
        df = df[(df['Parsed_Date'].dt.date >= start_date) & (df['Parsed_Date'].dt.date <= end_date)]
    
    if df.empty:
        return df, pd.DataFrame()

    # Day End Balance Calculation
    daily_df = df.groupby('Parsed_Date').last().reset_index()
    
    daily_df['Utilized_OD_Amount'] = od_limit - daily_df['Closing Balance']
    daily_df['Utilized_OD_Amount'] = daily_df['Utilized_OD_Amount'].apply(lambda x: max(0.0, x))
    
    daily_rate = (annual_interest_rate / 100.0) / 365.0
    daily_df['Daily_Interest'] = daily_df['Utilized_OD_Amount'] * daily_rate
    
    return df, daily_df

def format_excel_sheet(workbook, sheet_name):
    ws = workbook[sheet_name]
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

        header_name = str(col[0].value)
        if any(amt_key in header_name for amt_key in ["Amt", "Balance", "Amount", "Interest", "Total"]):
            for cell in list(col)[1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")

# ==================== STREAMLIT UI ====================

st.title("🏦 Bank Statement OD & Financial Dashboard")
st.write("Live Summary, Date-wise Interest Calculation, aur Custom Excel Export")

st.divider()

# Sidebar Setup
st.sidebar.header("⚙️ Settings & Inputs")
uploaded_file = st.sidebar.file_uploader("Bank Statement PDF Upload", type=["pdf"])
od_limit = st.sidebar.number_input("OD Limit (INR)", value=1000000.0, step=50000.0, format="%.2f")
interest_rate = st.sidebar.number_input("Interest Rate (% p.a.)", value=9.5, step=0.1, format="%.2f")

if uploaded_file is not None:
    # Initial Data Parse
    raw_df = parse_bank_statement_pdf_perfect(uploaded_file)
    
    if raw_df.empty:
        st.error("PDF se data read nahi ho saka. Format verify karein.")
    else:
        # Date Conversion for Filter
        raw_df['Parsed_Date'] = pd.to_datetime(raw_df['Value Dt'], format='%d/%m/%y', errors='coerce')
        raw_df['Parsed_Date'] = raw_df['Parsed_Date'].fillna(pd.to_datetime(raw_df['Date'], format='%d/%m/%y', errors='coerce'))
        raw_df = raw_df.dropna(subset=['Parsed_Date']).sort_values(by='Parsed_Date').reset_index(drop=True)
        
        min_date = raw_df['Parsed_Date'].min().date()
        max_date = raw_df['Parsed_Date'].max().date()
        
        st.sidebar.subheader("📅 Filter Date Range")
        date_range = st.sidebar.date_input("Select Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        
        start_date, end_date = None, None
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range[0], date_range[1]
        
        # Calculate OD & Filtered Data
        filtered_df, daily_summary = calculate_od_interest(raw_df, od_limit, interest_rate, start_date, end_date)
        
        if filtered_df.empty:
            st.warning("Selected Date Range me koi transactions nahi hain.")
        else:
            # Financial Metrics Summary
            total_withdrawal = filtered_df['Withdrawal Amt'].sum()
            total_deposit = filtered_df['Deposit Amt'].sum()
            
            # Cash Deposits Calculation (Keywords match: CASH, CDM, CSH, etc.)
            cash_mask = filtered_df['Narration'].str.contains(r'CASH|CDM|CSH|DEPOSIT BY CASH', case=False, na=False)
            total_cash_deposit = filtered_df[cash_mask & (filtered_df['Deposit Amt'] > 0)]['Deposit Amt'].sum()
            
            # Charges Calculation (Keywords match: CHARGE, CHG, FEE, INT.COLL, SMS, TAX, GST)
            charge_mask = filtered_df['Narration'].str.contains(r'CHARGE|CHG|FEE|INT\.COLL|TAX|GST|COMMISSION|PENALTY', case=False, na=False)
            total_charges = filtered_df[charge_mask & (filtered_df['Withdrawal Amt'] > 0)]['Withdrawal Amt'].sum()
            
            total_interest = daily_summary['Daily_Interest'].sum() if not daily_summary.empty else 0.0

            # ---------------- DISPLAY LIVE METRICS ----------------
            st.subheader("📊 Live Statement Summary")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            col1.metric("Total Deposits (Credit)", f"₹{total_deposit:,.2f}")
            col2.metric("Total Withdrawals (Debit)", f"₹{total_withdrawal:,.2f}")
            col3.metric("Total Cash Deposited", f"₹{total_cash_deposit:,.2f}")
            col4.metric("Bank Charges / Fees", f"₹{total_charges:,.2f}")
            col5.metric("Calculated OD Interest", f"₹{total_interest:,.2f}")

            st.divider()

            # ---------------- DOWNLOAD EXCEL BUTTON ----------------
            output_excel = "OD_Analysis_Report.xlsx"
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                # Sheet 1: Transactions
                export_df = filtered_df.drop(columns=['Parsed_Date'])
                export_df.to_excel(writer, sheet_name='Transactions', index=False)
                
                # Sheet 2: Daily OD Interest
                if not daily_summary.empty:
                    summary_export = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                    summary_export['Parsed_Date'] = summary_export['Parsed_Date'].dt.strftime('%d/%m/%Y')
                    summary_export.columns = ['Date', 'Closing Balance', 'Utilized OD Amount', 'Daily Interest (INR)']
                    summary_export.to_excel(writer, sheet_name='Daily_OD_Interest', index=False)
                
                # Sheet 3: Financial Summary Card
                overview_df = pd.DataFrame([
                    {"Metric": "Selected Start Date", "Value": str(start_date)},
                    {"Metric": "Selected End Date", "Value": str(end_date)},
                    {"Metric": "OD Limit", "Value": od_limit},
                    {"Metric": "Interest Rate (%)", "Value": interest_rate},
                    {"Metric": "Total Credits (Deposits)", "Value": total_deposit},
                    {"Metric": "Total Debits (Withdrawals)", "Value": total_withdrawal},
                    {"Metric": "Total Cash Deposited", "Value": total_cash_deposit},
                    {"Metric": "Total Bank Charges", "Value": total_charges},
                    {"Metric": "Total OD Interest Payable", "Value": total_interest}
                ])
                overview_df.to_excel(writer, sheet_name='Financial_Summary', index=False)
                
                wb = writer.book
                format_excel_sheet(wb, 'Transactions')
                if not daily_summary.empty:
                    format_excel_sheet(wb, 'Daily_OD_Interest')
                format_excel_sheet(wb, 'Financial_Summary')

            with open(output_excel, "rb") as fp:
                st.download_button(
                    label="📥 Download Detailed Excel Report",
                    data=fp,
                    file_name="OD_Analysis_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            # ---------------- TABS FOR DETAILED TABLES ----------------
            tab1, tab2, tab3 = st.tabs(["📝 All Transactions", "📈 Daily Interest Calculation", "🧾 Charges & Cash Breakdown"])
            
            with tab1:
                st.write("### Filtered Transactions")
                show_df = filtered_df.drop(columns=['Parsed_Date'])
                st.dataframe(show_df, use_container_width=True)
                
            with tab2:
                st.write("### Day-wise OD Interest Breakdown")
                if not daily_summary.empty:
                    disp_daily = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                    disp_daily['Parsed_Date'] = disp_daily['Parsed_Date'].dt.strftime('%d/%m/%Y')
                    disp_daily.columns = ['Date', 'Closing Balance', 'Utilized OD Amount', 'Daily Interest (INR)']
                    st.dataframe(disp_daily, use_container_width=True)
                else:
                    st.info("No OD calculations available for this range.")
                    
            with tab3:
                c1, c2 = st.columns(2)
                with c1:
                    st.write("### Cash Deposit Transactions")
                    cash_df = filtered_df[cash_mask & (filtered_df['Deposit Amt'] > 0)].drop(columns=['Parsed_Date'])
                    st.dataframe(cash_df, use_container_width=True)
                with c2:
                    st.write("### Detected Bank Charges / Taxes")
                    charges_df = filtered_df[charge_mask & (filtered_df['Withdrawal Amt'] > 0)].drop(columns=['Parsed_Date'])
                    st.dataframe(charges_df, use_container_width=True)

else:
    st.info("👈 Please upload a Bank Statement PDF file from the sidebar to view live dashboard.")
