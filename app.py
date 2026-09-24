import os
import re
import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(page_title="Bank Statement OD Calculator", page_icon="📊", layout="centered")

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

def calculate_od_interest(df, od_limit, annual_interest_rate):
    df['Parsed_Date'] = pd.to_datetime(df['Value Dt'], format='%d/%m/%y', errors='coerce')
    df['Parsed_Date'] = df['Parsed_Date'].fillna(pd.to_datetime(df['Date'], format='%d/%m/%y', errors='coerce'))
    
    df = df.dropna(subset=['Parsed_Date']).sort_values(by='Parsed_Date').reset_index(drop=True)
    
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
        if any(amt_key in header_name for amt_key in ["Amt", "Balance", "Amount", "Interest"]):
            for cell in list(col)[1:]:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")

# ==================== STREAMLIT UI ====================

st.title("🏦 Bank Statement OD Interest Calculator")
st.write("PDF statement upload karein, OD limit aur interest rate set karke instant calculation dekhein.")

st.divider()

# Input Fields
uploaded_file = st.file_uploader("Bank Statement PDF Upload Karein", type=["pdf"])

col1, col2 = st.columns(2)
with col1:
    od_limit = st.number_input("OD Limit (INR)", value=1000000.0, step=50000.0, format="%.2f")
with col2:
    interest_rate = st.number_input("Interest Rate (% p.a.)", value=9.5, step=0.1, format="%.2f")

if uploaded_file is not None:
    if st.button("Process & Calculate", type="primary", use_container_width=True):
        with st.spinner("PDF process ho raha hai..."):
            full_df = parse_bank_statement_pdf_perfect(uploaded_file)
            
            if full_df.empty:
                st.error("PDF se transactions extract nahi ho sake. File format check karein.")
            else:
                full_df, daily_summary = calculate_od_interest(full_df, od_limit, interest_rate)
                total_interest = daily_summary['Daily_Interest'].sum()
                
                st.success(f"Calculated Total Interest: INR {total_interest:,.2f}")
                
                # Excel Generation in Memory
                output_excel = "OD_Interest_Report.xlsx"
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    export_df = full_df.drop(columns=['Parsed_Date'])
                    export_df.to_excel(writer, sheet_name='Transactions', index=False)
                    
                    summary_export = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                    summary_export['Parsed_Date'] = summary_export['Parsed_Date'].dt.strftime('%d/%m/%Y')
                    summary_export.columns = ['Date', 'Closing Balance', 'Utilized OD Amount', 'Daily Interest (INR)']
                    summary_export.to_excel(writer, sheet_name='Daily_OD_Interest', index=False)
                    
                    wb = writer.book
                    format_excel_sheet(wb, 'Transactions')
                    format_excel_sheet(wb, 'Daily_OD_Interest')
                
                with open(output_excel, "rb") as fp:
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=fp,
                        file_name="OD_Interest_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.subheader("Daily Interest Summary Preview")
                preview_df = daily_summary[['Parsed_Date', 'Closing Balance', 'Utilized_OD_Amount', 'Daily_Interest']].copy()
                preview_df['Parsed_Date'] = preview_df['Parsed_Date'].dt.strftime('%d/%m/%Y')
                st.dataframe(preview_df, use_container_width=True)