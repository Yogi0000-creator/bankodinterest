import os
import re
import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(
    page_title="FinPulse | Bank Statement & OD Analytics", 
    page_icon="💳", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM PREMIUM STYLING (CSS) ====================
st.markdown("""
    <style>
    /* Main Background & Font Adjustments */
    .stApp {
        background-color: #F8FAFC;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Header Container */
    .header-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 24px 32px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
    }
    .header-title {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
        margin: 0 !important;
        letter-spacing: -0.5px;
    }
    .header-subtitle {
        font-size: 14px;
        color: #94A3B8;
        margin-top: 4px;
    }

    /* Metric Cards Styling */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.03), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
    }
    .metric-label {
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #64748B;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 22px;
        font-weight: 700;
        color: #0F172A;
    }
    
    /* Custom Color Accents for Metrics */
    .val-credit { color: #16A34A; }
    .val-debit { color: #DC2626; }
    .val-cash { color: #2563EB; }
    .val-charge { color: #D97706; }
    .val-interest { color: #7C3AED; }

    /* Button Customization */
    .stDownloadButton > button {
        width: 100%;
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: white !important;
        border: none !important;
        padding: 10px 18px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15) !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.3) !important;
        transform: translateY(-1px);
    }
    
    /* Hide Default Streamlit Style Elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ==================== PARSER & CALCULATION LOGIC ====================

def parse_bank_statement_pdf_perfect(pdf_file, password=None):
    all_rows = []
    pdf_kwargs = {}
    if password:
        pdf_kwargs['password'] = password

    with pdfplumber.open(pdf_file, **pdf_kwargs) as pdf:
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

    return pd.DataFrame(all_rows)

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
    
    if start_date and end_date:
        df = df[(df['Parsed_Date'].dt.date >= start_date) & (df['Parsed_Date'].dt.date <= end_date)]
    
    if df.empty:
        return df, pd.DataFrame()

    daily_df = df.groupby('Parsed_Date').last().reset_index()
    
    daily_df['Utilized_OD_Amount'] = od_limit - daily_df['Closing Balance']
    daily_df['Utilized_OD_Amount'] = daily_df['Utilized_OD_Amount'].apply(lambda x: max(0.0, x))
    
    daily_rate = (annual_interest_rate / 100.0) / 365.0
    daily_df['Daily_Interest'] = daily_df['Utilized_OD_Amount'] * daily_rate
    
    return df, daily_df

def format_excel_sheet(workbook, sheet_name):
    ws = workbook[sheet_name]
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

        header_name = str(col[0].value)
        if any(amt_key in header_name for amt_key in ["Amt", "Balance", "Amount", "Interest", "Total", "Value"]):
            for cell in list(col)[1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '₹#,##0.00'
                    cell.alignment = Alignment(horizontal="right")

def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# ==================== STREAMLIT UI LAYOUT ====================

st.markdown("""
    <div class="header-container">
        <div class="header-title">💼 Financial Statement & OD Analytics</div>
        <div class="header-subtitle">Automated Bank Statement Audit, Interest Computation & Category-wise Entry Exports</div>
    </div>
""", unsafe_allow_html=True)

# Sidebar Design
st.sidebar.markdown("### ⚙️ Configuration")
uploaded_file = st.sidebar.file_uploader("Upload Bank Statement (PDF)", type=["pdf"])
pdf_password = st.sidebar.text_input("PDF Password (If Encrypted)", type="password")

st.sidebar.markdown("---")
st.sidebar.markdown("### 💰 Credit Limits & Rates")
od_limit = st.sidebar.number_input("Sanctioned OD Limit (₹)", value=1000000.0, step=50000.0, format="%.2f")
interest_rate = st.sidebar.number_input("Interest Rate (% p.a.)", value=9.5, step=0.1, format="%.2f")

if uploaded_file is not None:
    try:
        raw_df = parse_bank_statement_pdf_perfect(uploaded_file, password=pdf_password if pdf_password else None)
        
        if raw_df.empty:
            st.error("⚠️ Unable to extract statement data. Please check PDF layout.")
        else:
            raw_df['Parsed_Date'] = pd.to_datetime(raw_df['Value Dt'], format='%d/%m/%y', errors='coerce')
            raw_df['Parsed_Date'] = raw_df['Parsed_Date'].fillna(pd.to_datetime(raw_df['Date'], format='%d/%m/%y', errors='coerce'))
            raw_df = raw_df.dropna(subset=['Parsed_Date']).sort_values(by='Parsed_Date').reset_index(drop=True)
            
            min_date = raw_df['Parsed_Date'].min().date()
            max_date = raw_df['Parsed_Date'].max().date()
            
            st.sidebar.markdown("---")
            st.sidebar.markdown("### 📅 Analysis Period")
            date_range = st.sidebar.date_input("Filter Dates", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            
            start_date, end_date = None, None
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date, end_date = date_range[0], date_range[1]
            
            filtered_df, daily_summary = calculate_od_interest(raw_df, od_limit, interest_rate, start_date, end_date)
            
            if filtered_df.empty:
                st.warning("No transactions found in selected date range.")
            else:
                # Prepare Category Dataframes
                deposits_df = filtered_df[filtered_df['Deposit Amt'] > 0].drop(columns=['Parsed_Date'])
                withdrawals_df = filtered_df[filtered_df['Withdrawal Amt'] > 0].drop(columns=['Parsed_Date'])
                
                cash_mask = filtered_df['Narration'].str.contains(r'CASH|CDM|CSH|DEPOSIT BY CASH', case=False, na=False)
                cash_df = filtered_df[cash_mask & (filtered_df['Deposit Amt'] > 0)].drop(columns=['Parsed_Date'])
                
                charge_mask = filtered_df['Narration'].str.contains(r'CHARGE|CHG|FEE|INT\.COLL|TAX|GST|COMMISSION|PENALTY', case=False, na=False)
                charges_df = filtered_df[charge_mask & (filtered_df['Withdrawal Amt'] > 0)].drop(columns=['Parsed_Date'])

                # Metrics Calculation
                total_withdrawal = filtered_df['Withdrawal Amt'].sum()
                total_deposit = filtered_df['Deposit Amt'].sum()
                total_cash_deposit = cash_df['Deposit Amt'].sum() if not cash_df.empty else 0.0
                total_charges = charges_df['Withdrawal Amt'].sum() if not charges_df.empty else 0.0
                total_interest = daily_summary['Daily_Interest'].sum() if not daily_summary.empty else 0.0

                # Metric Cards UI
                c1, c2, c3, c4, c5 = st.columns(5)
                
                c1.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total Deposits</div>
                        <div class="metric-value val-credit">₹{total_deposit:,.2f}</div>
                    </div>
                """, unsafe_allow_html=True)
                c1.download_button("📥 Export Deposits", convert_df_to_csv(deposits_df), "Deposits_Entries.csv", "text/csv")

                c2.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Total Debits</div>
                        <div class="metric-value val-debit">₹{total_withdrawal:,.2f}</div>
                    </div>
                """, unsafe_allow_html=True)
                c2.download_button("📥 Export Debits", convert_df_to_csv(withdrawals_df), "Debits_Entries.csv", "text/csv")

                c3.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Cash Deposited</div>
                        <div class="metric-value val-cash">₹{total_cash_deposit:,.2f}</div>
                    </div>
                """, unsafe_allow_html=True)
                c3.download_button("📥 Export Cash", convert_df_to_csv(cash_df), "Cash_Deposits_Entries.csv", "text/csv")

                c4.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">Bank Charges</div>
                        <div class="metric-value val-charge">₹{total_charges:,.2f}</div>
                    </div>
                """, unsafe_allow_html=True)
                c4.download_button("📥 Export Charges", convert_df_to_csv(charges_df), "Bank_Charges_Entries.csv", "text/csv")

                c5.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-label">OD Interest</div>
                        <div class="metric-value val-interest">₹{total_interest:,.2f}</div>
                    </div>
                """, unsafe_allow_html=True)
                if not daily_summary.empty:
                    daily_exp = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                    daily_exp['Parsed_Date'] = daily_exp['Parsed_Date'].dt.strftime('%d/%m/%Y')
                    c5.download_button("📥 Export OD Ledger", convert_df_to_csv(daily_exp), "OD_Interest_Ledger.csv", "text/csv")

                st.markdown("<br>", unsafe_allow_html=True)

                # Export Complete Excel File
                output_excel = "OD_Analysis_Report.xlsx"
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    filtered_df.drop(columns=['Parsed_Date']).to_excel(writer, sheet_name='All_Transactions', index=False)
                    deposits_df.to_excel(writer, sheet_name='Deposits_Only', index=False)
                    withdrawals_df.to_excel(writer, sheet_name='Debits_Only', index=False)
                    cash_df.to_excel(writer, sheet_name='Cash_Deposits', index=False)
                    charges_df.to_excel(writer, sheet_name='Bank_Charges', index=False)
                    
                    if not daily_summary.empty:
                        summary_export = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                        summary_export['Parsed_Date'] = summary_export['Parsed_Date'].dt.strftime('%d/%m/%Y')
                        summary_export.columns = ['Date', 'Closing Balance', 'Utilized OD Amount', 'Daily Interest (INR)']
                        summary_export.to_excel(writer, sheet_name='Daily_OD_Interest', index=False)
                    
                    overview_df = pd.DataFrame([
                        {"Metric": "Statement Start Date", "Value": str(start_date)},
                        {"Metric": "Statement End Date", "Value": str(end_date)},
                        {"Metric": "OD Limit (INR)", "Value": od_limit},
                        {"Metric": "Interest Rate (% p.a.)", "Value": interest_rate},
                        {"Metric": "Total Deposits (Credits)", "Value": total_deposit},
                        {"Metric": "Total Withdrawals (Debits)", "Value": total_withdrawal},
                        {"Metric": "Total Cash Deposited", "Value": total_cash_deposit},
                        {"Metric": "Total Bank Charges", "Value": total_charges},
                        {"Metric": "Total Calculated Interest", "Value": total_interest}
                    ])
                    overview_df.to_excel(writer, sheet_name='Executive_Summary', index=False)
                    
                    wb = writer.book
                    for sheet in wb.sheetnames:
                        format_excel_sheet(wb, sheet)

                st.sidebar.markdown("---")
                with open(output_excel, "rb") as fp:
                    st.sidebar.download_button(
                        label="📊 Download Full Multi-Sheet Excel",
                        data=fp,
                        file_name="OD_Executive_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # Modern Tabbed Interface for Detailed Viewing
                tab1, tab2, tab3, tab4 = st.tabs(["📋 Master Transactions", "📈 OD Interest Ledger", "💵 Cash Entries", "🏛️ Bank Charges"])
                
                with tab1:
                    show_df = filtered_df.drop(columns=['Parsed_Date'])
                    st.dataframe(show_df, use_container_width=True, height=420)
                    
                with tab2:
                    if not daily_summary.empty:
                        disp_daily = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                        disp_daily['Parsed_Date'] = disp_daily['Parsed_Date'].dt.strftime('%d/%m/%Y')
                        disp_daily.columns = ['Date', 'Closing Balance', 'Utilized OD Amount', 'Daily Interest (INR)']
                        st.dataframe(disp_daily, use_container_width=True, height=420)
                    else:
                        st.info("No OD calculations available for this period.")
                        
                with tab3:
                    st.markdown("##### 💵 All Cash Deposit Entries")
                    st.dataframe(cash_df, use_container_width=True, height=350)
                    
                with tab4:
                    st.markdown("##### 🏛️ All Detected Bank Charges & Taxes")
                    st.dataframe(charges_df, use_container_width=True, height=350)

    except Exception as e:
        err_msg = str(e).lower()
        if "password" in err_msg or "encrypted" in err_msg or "authenticate" in err_msg:
            st.error("🔒 PDF is Password Protected. Please provide correct password in sidebar.")
        else:
            st.error(f"Error processing document: {str(e)}")

else:
    st.info("👋 Welcome! Upload a Bank Statement PDF from the left sidebar to get started.")
