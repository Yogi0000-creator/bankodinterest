def calculate_od_interest(df, od_limit, annual_interest_rate, start_date=None, end_date=None):
    df['Parsed_Date'] = pd.to_datetime(df['Value Dt'], format='%d/%m/%y', errors='coerce')
    df['Parsed_Date'] = df['Parsed_Date'].fillna(pd.to_datetime(df['Date'], format='%d/%m/%y', errors='coerce'))
    df = df.dropna(subset=['Parsed_Date']).sort_values(by='Parsed_Date').reset_index(drop=True)
    
    if start_date and end_date:
        df = df[(df['Parsed_Date'].dt.date >= start_date) & (df['Parsed_Date'].dt.date <= end_date)]
    
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()

    # 1. Daily Last Balance
    daily_last = df.groupby('Parsed_Date')['Closing Balance'].last().reset_index()
    
    # 2. Continuous Date Range
    full_date_range = pd.date_range(start=daily_last['Parsed_Date'].min(), end=daily_last['Parsed_Date'].max(), freq='D')
    daily_df = pd.DataFrame({'Parsed_Date': full_date_range})
    daily_df = pd.merge(daily_df, daily_last, on='Parsed_Date', how='left')
    daily_df['Closing Balance'] = daily_df['Closing Balance'].ffill()

    # 3. Utilized Amount & Daily Interest Calculation
    daily_df['Utilized_OD_Amount'] = daily_df['Closing Balance'].apply(lambda bal: abs(bal) if bal < 0 else 0.0)
    daily_rate = (annual_interest_rate / 100.0) / 365.0
    daily_df['Daily_Interest'] = daily_df['Utilized_OD_Amount'] * daily_rate

    # 4. Monthly System Calculated Interest
    daily_df['Year_Month'] = daily_df['Parsed_Date'].dt.to_period('M')
    monthly_calc = daily_df.groupby('Year_Month')['Daily_Interest'].sum().reset_index()
    monthly_calc.rename(columns={'Daily_Interest': 'Calculated Interest (System)'}, inplace=True)

    # 5. Extract Actual Bank Charged Interest (Flexible Regex Matching)
    int_pattern = r'INT|INTEREST|INT\.COLL|INTEREST DEBITED|INT DEBIT|OD INT'
    int_mask = df['Narration'].str.contains(int_pattern, case=False, na=False)
    
    bank_int_df = df[int_mask & (df['Withdrawal Amt'] > 0)].copy()
    
    # Custom Function: Check if narration specifies a target month (e.g. "TILL 31-AUG-2026")
    def get_target_month(row):
        narration = str(row['Narration']).upper()
        # Look for month patterns like 31-AUG-2026 or AUG-2026
        match = re.search(r'(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[a-zA-Z]*[\s\-]*(20\d{2}|\d{2})', narration)
        if match:
            month_str = match.group(1)
            year_str = match.group(2)
            if len(year_str) == 2:
                year_str = "20" + year_str
            dt_parsed = pd.to_datetime(f"01-{month_str}-{year_str}", format='%d-%b-%Y', errors='coerce')
            if not pd.isna(dt_parsed):
                return dt_parsed.to_period('M')
        # Default to transaction date's month
        return row['Parsed_Date'].to_period('M')

    if not bank_int_df.empty:
        bank_int_df['Year_Month'] = bank_int_df.apply(get_target_month, axis=1)
        monthly_bank = bank_int_df.groupby('Year_Month')['Withdrawal Amt'].sum().reset_index()
        monthly_bank.rename(columns={'Withdrawal Amt': 'Bank Charged Interest (Actual)'}, inplace=True)
    else:
        monthly_bank = pd.DataFrame(columns=['Year_Month', 'Bank Charged Interest (Actual)'])

    # 6. Merge & Reconciliation
    monthly_report = pd.merge(monthly_calc, monthly_bank, on='Year_Month', how='left').fillna(0.0)
    monthly_report['Difference (Excess/Short)'] = monthly_report['Bank Charged Interest (Actual)'] - monthly_report['Calculated Interest (System)']
    
    monthly_report['Month'] = monthly_report['Year_Month'].astype(str)
    monthly_report = monthly_report[['Month', 'Calculated Interest (System)', 'Bank Charged Interest (Actual)', 'Difference (Excess/Short)']]

    return df, daily_df, monthly_report
