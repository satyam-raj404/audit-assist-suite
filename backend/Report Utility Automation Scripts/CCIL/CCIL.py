import pandas as pd
import os
from datetime import timedelta, datetime, time
import os
import numpy as np
import warnings
warnings.filterwarnings('ignore')

def CCIL_main(excel_file, output_folder_path):

    output_folder_path_upd = rf"{output_folder_path}\CCIL"
    # Define file paths
    output_file_forex = "CCIL_Forex_Result.xlsx"
    output_file_der = "CCIL_Der_Result.xlsx"
    full_path_forex = rf"{output_folder_path_upd}\{output_file_forex}"
    os.makedirs(os.path.dirname(full_path_forex), exist_ok=True)
    full_path_der = rf"{output_folder_path_upd}\{output_file_der}"
    os.makedirs(os.path.dirname(full_path_der), exist_ok=True)

    # Forex

    # Forex Client
    df_deal_status_report = pd.read_excel(excel_file, sheet_name = 'Client Trade status report ',header=1)
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal dump Merchant')

    df_merged = pd.merge(df_deal_status_report, df_deal_dump, left_on = 'Transaction Reference No', right_on = 'System ID', how = 'inner')
    df_merged = df_merged[['Transaction Reference No', 'Transaction Type',
                        'Client Id','Client Name', 'Trade Date_y', 'Input Time',
                        'Reporting Date and Time'
                        ]]
    df_merged = df_merged.rename(columns={
        'Trade Date_y': 'Trade Date',
    })
    df_merged['Input Time'] = df_merged['Input Time'].astype(str).str.zfill(6)

    # Insert colons to match HH:MM:SS format
    df_merged['Formatted Time'] = df_merged['Input Time'].str.slice(0,2) + ':' + \
                                df_merged['Input Time'].str.slice(2,4) + ':' + \
                                df_merged['Input Time'].str.slice(4,6)

    df_merged['Reporting Date as CCIL'] = pd.to_datetime(df_merged['Reporting Date and Time']).dt.date
    df_merged['Reported Time as CCIL'] = pd.to_datetime(df_merged['Reporting Date and Time']).dt.time
    df_merged['Trade Date'] = pd.to_datetime(df_merged['Trade Date'])
    df_merged['To be reported by'] = df_merged['Trade Date'] + timedelta(days=1)
    df_merged['To Be Reported within'] = datetime.strptime('12:00:00', '%H:%M:%S').time()
    df_final = pd.DataFrame({
        'System No.': df_merged['Transaction Reference No'],
        'Transaction Type': df_merged['Transaction Type'],
        'Trade Date': df_merged['Trade Date'], 
        'Input time': df_merged['Formatted Time'], 
        'To be reported by': pd.to_datetime(df_merged['To be reported by']).dt.date,     
        'To Be Reported within': df_merged['To Be Reported within'],
        'Reporting Date as CCIL Raw': df_merged['Reporting Date and Time'],
        'Reporting Date as CCIL' : pd.to_datetime(df_merged['Reporting Date as CCIL']).dt.date,
        'Reported Time as CCIL' : df_merged['Reported Time as CCIL']})

    df_final['Remarks']=df_final.apply(
    lambda row: 'PASS' if (row['To be reported by'] == row['Reporting Date as CCIL']) 
        #or (row['Reported Time as CCIL'] < row['To Be Reported within']) 
        else 'FAIL',
        axis=1)
    df_final.apply(
        lambda row: 'PASS' if (row['To be reported by'] == row['Reporting Date as CCIL']) 
            or (row['Reported Time as CCIL'] < row['To Be Reported within']) 
            else 'FAIL',
            axis=1) 
    
    df_CCIL_forex_client = df_final

    # Forex FCY FCY
    df_deal_status_report = pd.read_excel(excel_file, sheet_name = 'FCYFWD deal status report ')
    df_outstanding_trade_report = pd.read_excel(excel_file, sheet_name = 'FCYFWD Outstanding report')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal DUMP IB ')

        # joining deal status report and outstanding report to get deal received time
    tran_id_deal_time = pd.merge(df_deal_status_report, df_outstanding_trade_report, left_on = 'Transcation Reference No.', right_on = 'Member Ref Num', how = 'inner')[['Transcation Reference No.', 'Deal Received Time']]

    # memeber id after "/" for comparison
    tran_id_deal_time['New_tran_no'] = tran_id_deal_time['Transcation Reference No.'].str.split('/').str[-1]
    df_merged = pd.merge(tran_id_deal_time, df_deal_dump, left_on = 'New_tran_no', right_on = 'System ID', how = 'inner')

    # splitting deal received time into date and time
    
    # Extract the first 8 characters as date (YYYYMMDD)
    df_merged['Reporting Date as CCIL'] = pd.to_datetime(df_merged['Deal Received Time']).dt.date
    df_merged['Reported Time as CCIL'] = pd.to_datetime(df_merged['Deal Received Time']).dt.time

    # converting to correct time format
    df_merged['Input Time'] = df_merged['Input Time'].astype(str).str.zfill(6)
    df_merged['Formatted Time'] = pd.to_datetime(df_merged['Input Time'], format='%H%M%S').dt.time

    df_merged['Trade Date'] = pd.to_datetime(df_merged['Trade Date'])
    df_merged['Trade Timestamp'] = df_merged.apply(
        lambda row: datetime.combine(row['Trade Date'], row['Formatted Time']),
        axis=1
    )

    def get_reporting_deadline(trade_time):
        cutoff = trade_time.replace(hour=17, minute=0, second=0, microsecond=0)
        if trade_time <= cutoff:
            report_date = trade_time.date()
            report_time = datetime.combine(report_date, datetime.strptime('17:30:00', '%H:%M:%S').time())
        else:
            report_date = trade_time.date() + timedelta(days=1)
            report_time = datetime.combine(report_date, datetime.strptime('10:00:00', '%H:%M:%S').time())
        return pd.Series([report_date, report_time.time()])

    df_merged[['To be reported by', 'To Be Reported within']] = df_merged['Trade Timestamp'].apply(get_reporting_deadline)
    df_final = pd.DataFrame({
    'System No.': df_merged['New_tran_no'],
    'Product': df_merged['Prod'],
    'Trade Date': df_merged['Trade Date'], 
    'Input time': df_merged['Formatted Time'], 
    'To be reported by': df_merged['To be reported by'],     
    'To Be Reported within': df_merged['To Be Reported within'],
    'Reporting Date as CCIL Raw': df_merged['Deal Received Time'],
    'Reporting Date as CCIL' : df_merged['Reporting Date as CCIL'],
    'Reported Time as CCIL' : df_merged['Reported Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if (row['To be reported by'] == row['Reporting Date as CCIL']) and (row['Reported Time as CCIL'] < row['To Be Reported within']) else 'FAIL',
        axis=1)
    
    df_CCIL_forex_FCY_FCY = df_final

    # Forex FCY INR
    df_deal_status_report = pd.read_excel(excel_file, sheet_name = 'INR FWD DEAL STATUS REPORT')
    df_outstanding_trade_report = pd.read_excel(excel_file, sheet_name = 'INRFWD Outstanding trade report')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal DUMP IB ')

    tran_id_deal_time = pd.merge(df_deal_status_report, df_outstanding_trade_report, left_on = 'Transcation Reference No.', right_on = 'Member Ref Num', how = 'inner')[['Transcation Reference No.', 'Deal Received Time']]
    tran_id_deal_time['New_tran_no'] = tran_id_deal_time['Transcation Reference No.'].str.split('/').str[-1]
    df_merged = pd.merge(tran_id_deal_time, df_deal_dump, left_on = 'New_tran_no', right_on = 'System ID', how = 'inner')
    df_merged['Input Time'] = df_merged['Input Time'].astype(str).str.zfill(6)

    # Insert colons to match HH:MM:SS format
    df_merged['Formatted Time'] = df_merged['Input Time'].str.slice(0,2) + ':' + \
                                df_merged['Input Time'].str.slice(2,4) + ':' + \
                                df_merged['Input Time'].str.slice(4,6)
    
    # Create new columns without modifying the original
    df_merged['Reporting Date as CCIL'] = pd.to_datetime(df_merged['Deal Received Time']).dt.date
    df_merged['Reported Time as CCIL'] = pd.to_datetime(df_merged['Deal Received Time']).dt.time

    times = pd.to_datetime(df_merged['Formatted Time'], format='%H:%M:%S')
    def calculate_reporting_time(t):
        upper_bound = t.replace(minute=0, second=0) + timedelta(hours=1)
        reporting_time = upper_bound + timedelta(minutes=30)
        
        # Ensure same date (we assume all times are on the same day)
        if reporting_time.date() != t.date():
            reporting_time = t.replace(hour=23, minute=59, second=0)
        
        return reporting_time.time()

    # Apply function
    reporting_times = times.map(calculate_reporting_time)
    df_merged['To Be Reported within'] = pd.to_datetime(reporting_times.astype(str), format='%H:%M:%S').dt.time

    df_final = pd.DataFrame({
    'System No.': df_merged['New_tran_no'],
    'Product': df_merged['Prod'],
    'Trade Date': df_merged['Trade Date'], 
    'Input time': df_merged['Formatted Time'], 
    'To be reported by': df_merged['Trade Date'],     
    'To Be Reported within': df_merged['To Be Reported within'],
   'Reporting Date as CCIL Raw': df_merged['Deal Received Time'],
    'Reporting Date as CCIL' : df_merged['Reporting Date as CCIL'],
    'Reported Time as CCIL' : df_merged['Reported Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if row['To Be Reported within'] > row['Reported Time as CCIL'] else 'FAIL',
        axis=1)

    df_CCIL_forex_FCY_INR = df_final

        # Save all KPI results into one Excel file with separate sheets
    with pd.ExcelWriter(full_path_forex, engine="openpyxl") as writer:
        df_CCIL_forex_client.to_excel(writer, sheet_name="Forex Client", index=False)
        df_CCIL_forex_FCY_FCY.to_excel(writer, sheet_name="Forex FCY FCY", index=False)
        df_CCIL_forex_FCY_INR.to_excel(writer, sheet_name="Forex FCY INR", index=False)

    # Derivatives

    # INR IRS
    df_inr_irs = pd.read_excel(excel_file, sheet_name = 'INR IRS')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal Dump')  
    df_deal_dump = df_deal_dump[['System ID', 'Prod', 'Cpty', 'Counterparty', 'Instrument', 'Amount1', 'Input Date', 'Input Time', 'Trade Date']].rename(columns={'Amount1':'Amount'})
    inr_irs_deal_dump = df_deal_dump[(df_deal_dump['Prod'].astype(str).str.strip() == 'IRS') & (df_deal_dump['Instrument'].astype(str).str.strip() == 'INR')]
    inr_irs_deal_dump['Input Time'] = inr_irs_deal_dump['Input Time'].astype(str).str.zfill(6)
    inr_irs_deal_dump['Formatted Time'] = inr_irs_deal_dump['Input Time'].str.slice(0,2) + ':' + \
                                        inr_irs_deal_dump['Input Time'].str.slice(2,4) + ':' + \
                                        inr_irs_deal_dump['Input Time'].str.slice(4,6)
    inr_irs_deal_dump['Input Date'] = pd.to_datetime(inr_irs_deal_dump['Input Date']).dt.date
    inr_irs_deal_dump['Trade Date'] = pd.to_datetime(inr_irs_deal_dump['Trade Date']).dt.date
    df_inr_irs = df_inr_irs[["Member's Ref.",'Date of Reporting']]
    inr_irs_merged = pd.merge(df_inr_irs, inr_irs_deal_dump, how='inner', left_on = "Member's Ref.", right_on ='System ID', indicator=True)
    inr_irs_merged.rename(columns={'_merge': 'Source'}, inplace=True)

    inr_irs_merged['Source'] = inr_irs_merged['Source'].map({
        'left_only': 'Only in INR IRS',
        'right_only': 'Only in Deal Dump',
        'both': 'In Both Reports'
    })
    inr_irs_merged['Final_ID'] = inr_irs_merged["Member's Ref."].fillna(inr_irs_merged['System ID'])
    times = pd.to_datetime(inr_irs_merged['Formatted Time'], format='%H:%M:%S')
    def calculate_reporting_time(t):
        upper_bound = t.replace(minute=0, second=0) + timedelta(hours=1)
        reporting_time = upper_bound + timedelta(minutes=30)
        
        # Ensure same date (we assume all times are on the same day)
        if reporting_time.date() != t.date():
            reporting_time = t.replace(hour=23, minute=59, second=0)
        
        return reporting_time.time()

    # Apply function
    reporting_times = times.map(calculate_reporting_time)
    inr_irs_merged['To Be Reported within'] = pd.to_datetime(reporting_times.astype(str), format='%H:%M:%S').dt.time
    inr_irs_merged['Reporting Date as CCIL'] = pd.to_datetime(inr_irs_merged['Date of Reporting']).dt.date
    inr_irs_merged['Reported Time as CCIL'] = pd.to_datetime(inr_irs_merged['Date of Reporting']).dt.time
    df_final = pd.DataFrame({
    'System No.': inr_irs_merged['Final_ID'],
    'Product': inr_irs_merged['Prod'],
    'Trade Date': inr_irs_merged['Trade Date'], 
    'Input time': inr_irs_merged['Formatted Time'], 
    'To be reported by': inr_irs_merged['Trade Date'],     
    'To Be Reported within': inr_irs_merged['To Be Reported within'],
   'Date of Reporting': inr_irs_merged['Date of Reporting'],
    'Reporting Date as CCIL' : inr_irs_merged['Reporting Date as CCIL'],
    'Reported Time as CCIL' : inr_irs_merged['Reported Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if row['To Be Reported within'] > row['Reported Time as CCIL'] else 'FAIL',
        axis=1)
    
    def calculate_failure_duration(row):
        if row['Remarks'] == 'FAIL':
            # Combine Trade Date with time to create full datetime objects
            reported_dt = row['Date of Reporting']
            to_be_reported_dt = datetime.combine(row['To be reported by'], row['To Be Reported within'])
            
            delta = abs(reported_dt - to_be_reported_dt)
            days = delta.days
            hours, remainder = divmod(delta.seconds, 3600)
            minutes = remainder // 60
            return f"{days} days, {hours} hours, {minutes} minutes"
        else:
            return None

    df_final['Failure Duration'] = df_final.apply(calculate_failure_duration, axis=1)
 
    df_der_INR_IRS = df_final

    # IB FCY IRS
    # Reading both required excel sheets
    df_ib_fcy_irs = pd.read_excel(excel_file, sheet_name = 'IB FCY IRS')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal Dump')

    # Selecting the required columns from deal dump
    df_deal_dump = df_deal_dump[['System ID', 'Prod', 'Cpty', 'Counterparty', 'Instrument', 'Amount1', 'Input Date', 'Input Time', 'Trade Date']].rename(columns={'Amount1':'Amount'})

    ib_fcy_irs_deal_dump = df_deal_dump[
        (df_deal_dump['Prod'].astype(str).str.strip() == 'IRS') &
        (df_deal_dump['Instrument'].astype(str).str.strip().ne('INR')) &
        (~df_deal_dump['Instrument'].astype(str).str.contains('-')) &
        (df_deal_dump['Counterparty'].astype(str).str.contains('BANK', case=False, na=False))
    ]

    ib_fcy_irs_deal_dump['Input Time'] = ib_fcy_irs_deal_dump['Input Time'].astype(str).str.zfill(6)
    ib_fcy_irs_deal_dump['Formatted Time'] = ib_fcy_irs_deal_dump['Input Time'].str.slice(0,2) + ':' + \
                                        ib_fcy_irs_deal_dump['Input Time'].str.slice(2,4) + ':' + \
                                        ib_fcy_irs_deal_dump['Input Time'].str.slice(4,6)
    ib_fcy_irs_deal_dump['Input Date'] = pd.to_datetime(ib_fcy_irs_deal_dump['Input Date']).dt.date
    ib_fcy_irs_deal_dump['Trade Date'] = pd.to_datetime(ib_fcy_irs_deal_dump['Trade Date']).dt.date

    ib_fcy_irs_deal_dump['Formatted Date-Time'] = pd.to_datetime(ib_fcy_irs_deal_dump['Input Date'].astype(str) + ' ' + ib_fcy_irs_deal_dump['Formatted Time'].astype(str))

    df_ib_fcy_irs = df_ib_fcy_irs[["Member reference No",'Transaction Type', 'Matching Date and Time']]
    ib_fcy_irs_merged = pd.merge(df_ib_fcy_irs, ib_fcy_irs_deal_dump, how='inner', left_on = "Member reference No", right_on ='System ID', indicator=True)
    ib_fcy_irs_merged.rename(columns={'_merge': 'Source'}, inplace=True)

    ib_fcy_irs_merged['Source'] = ib_fcy_irs_merged['Source'].map({
        'left_only': 'Only in IB FCY IRS',
        'right_only': 'Only in Deal Dump',
        'both': 'In Both Reports'
    })
    ib_fcy_irs_merged['Final_ID'] = ib_fcy_irs_merged["Member reference No"].fillna(ib_fcy_irs_merged['System ID'])
    times = pd.to_datetime(ib_fcy_irs_merged['Formatted Date-Time'])

    def calculate_reporting_time(t):
        cutoff_time = time(17, 0)  # 5:00 PM

        if t.time() < cutoff_time:
            # Report by 5:30 PM same day
            reporting_time = t.replace(hour=17, minute=30, second=0)
        else:
            # Report by 10:00 AM next day
            reporting_time = (t + timedelta(days=1)).replace(hour=10, minute=0, second=0)

        return reporting_time

    # Apply function
    reporting_times = times.map(calculate_reporting_time)

    ib_fcy_irs_merged['To Be Reported within'] = pd.to_datetime(reporting_times.astype(str))
    ib_fcy_irs_merged['Reporting Date-Time as CCIL'] = pd.to_datetime(ib_fcy_irs_merged['Matching Date and Time'])
    df_final = pd.DataFrame({
    'System No.': ib_fcy_irs_merged['Final_ID'],
    'Product': ib_fcy_irs_merged['Prod'],
    'Instrument':ib_fcy_irs_merged['Instrument'],
    'Counterparty':ib_fcy_irs_merged['Counterparty'],
    'Trade Date': ib_fcy_irs_merged['Trade Date'], 
    'Input time': ib_fcy_irs_merged['Formatted Time'],     
    'To Be Reported within': ib_fcy_irs_merged['To Be Reported within'],
    'Date of Reporting': ib_fcy_irs_merged['Matching Date and Time'],
    'Reporting Date-Time as CCIL' : ib_fcy_irs_merged['Reporting Date-Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if row['To Be Reported within'] > row['Reporting Date-Time as CCIL'] else 'FAIL',
        axis=1)
    
    def calculate_failure_duration(row):
        if row['Remarks'] == 'FAIL':
            # Combine Trade Date with time to create full datetime objects
            reported_dt = row['Reporting Date-Time as CCIL']
            to_be_reported_dt = row['To Be Reported within']
            
            delta = abs(reported_dt - to_be_reported_dt)
            days = delta.days
            hours, remainder = divmod(delta.seconds, 3600)
            minutes = remainder // 60
            return f"{days} days, {hours} hours, {minutes} minutes"
        else:
            return None

    df_final['Failure Duration'] = df_final.apply(calculate_failure_duration, axis=1)
    df_der_IB_FCY_IRS = df_final
    
    # Client FCY IRS
    # Reading both required excel sheets
    df_Client_fcy_irs = pd.read_excel(excel_file, sheet_name = 'Client FCY IRS')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal Dump')

    # Selecting the required columns from deal dump
    df_deal_dump = df_deal_dump[['System ID', 'Prod', 'Cpty', 'Counterparty', 'Instrument', 'Amount1', 'Input Date', 'Input Time', 'Trade Date']].rename(columns={'Amount1':'Amount'})

    Client_fcy_irs_deal_dump = df_deal_dump[
        (df_deal_dump['Prod'].astype(str).str.strip() == 'IRS') &
        (df_deal_dump['Instrument'].astype(str).str.strip().ne('INR')) &
        (~df_deal_dump['Instrument'].astype(str).str.contains('-')) &
        (~df_deal_dump['Counterparty'].astype(str).str.contains('BANK', case=False, na=False))
    ]
    Client_fcy_irs_deal_dump['Input Time'] = Client_fcy_irs_deal_dump['Input Time'].astype(str).str.zfill(6)
    Client_fcy_irs_deal_dump['Formatted Time'] = Client_fcy_irs_deal_dump['Input Time'].str.slice(0,2) + ':' + \
                                        Client_fcy_irs_deal_dump['Input Time'].str.slice(2,4) + ':' + \
                                        Client_fcy_irs_deal_dump['Input Time'].str.slice(4,6)
    Client_fcy_irs_deal_dump['Input Date'] = pd.to_datetime(Client_fcy_irs_deal_dump['Input Date']).dt.date
    Client_fcy_irs_deal_dump['Trade Date'] = pd.to_datetime(Client_fcy_irs_deal_dump['Trade Date']).dt.date

    Client_fcy_irs_deal_dump['Formatted Date-Time'] = pd.to_datetime(Client_fcy_irs_deal_dump['Input Date'].astype(str) + ' ' + Client_fcy_irs_deal_dump['Formatted Time'].astype(str))
    df_Client_fcy_irs = df_Client_fcy_irs[["Member reference No", 'Matching Date and Time']]
    Client_fcy_irs_merged = pd.merge(df_Client_fcy_irs, Client_fcy_irs_deal_dump, how='inner', left_on = "Member reference No", right_on ='System ID', indicator=True)
    Client_fcy_irs_merged.rename(columns={'_merge': 'Source'}, inplace=True)

    Client_fcy_irs_merged['Source'] = Client_fcy_irs_merged['Source'].map({
        'left_only': 'Only in IB FCY IRS',
        'right_only': 'Only in Deal Dump',
        'both': 'In Both Reports'
    })
    Client_fcy_irs_merged['Final_ID'] = Client_fcy_irs_merged["Member reference No"].fillna(Client_fcy_irs_merged['System ID'])
    times = pd.to_datetime(Client_fcy_irs_merged['Formatted Date-Time'])

    def calculate_reporting_time(t):
        reporting_time = (t + timedelta(days=1)).replace(hour=12, minute=0, second=0)
        return reporting_time

    # Apply function
    reporting_times = times.map(calculate_reporting_time)
    Client_fcy_irs_merged['To Be Reported within'] = pd.to_datetime(reporting_times.astype(str))
    Client_fcy_irs_merged['Reporting Date-Time as CCIL'] = pd.to_datetime(Client_fcy_irs_merged['Matching Date and Time'])
    df_final = pd.DataFrame({
    'System No.': Client_fcy_irs_merged['Final_ID'],
    'Product': Client_fcy_irs_merged['Prod'],
    'Instrument':Client_fcy_irs_merged['Instrument'],
    'Counterparty':Client_fcy_irs_merged['Counterparty'],
    'Trade Date': Client_fcy_irs_merged['Trade Date'], 
    'Input time': Client_fcy_irs_merged['Formatted Time'],     
    'To Be Reported within': Client_fcy_irs_merged['To Be Reported within'],
   'Date of Reporting': Client_fcy_irs_merged['Matching Date and Time'],
    'Reporting Date-Time as CCIL' : Client_fcy_irs_merged['Reporting Date-Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if row['To Be Reported within'] > row['Reporting Date-Time as CCIL'] else 'FAIL',
        axis=1)
    
    def calculate_failure_duration(row):
        if row['Remarks'] == 'FAIL':
            # Combine Trade Date with time to create full datetime objects
            reported_dt = row['Reporting Date-Time as CCIL']
            to_be_reported_dt = row['To Be Reported within']
            
            delta = abs(reported_dt - to_be_reported_dt)
            days = delta.days
            hours, remainder = divmod(delta.seconds, 3600)
            minutes = remainder // 60
            return f"{days} days, {hours} hours, {minutes} minutes"
        else:
            return None

    df_final['Failure Duration'] = df_final.apply(calculate_failure_duration, axis=1)
    df_der_Client_FCY_IRS = df_final

    # Option_FCY_INR_IB

    # Reading both required excel sheets
    df_ib_fcy_irs = pd.read_excel(excel_file, sheet_name = 'Option FCY INR IB')
    df_deal_dump = pd.read_excel(excel_file, sheet_name = 'Deal Dump')

    # Selecting the required columns from deal dump
    df_deal_dump = df_deal_dump[['System ID', 'Prod', 'Cpty', 'Counterparty', 'Instrument', 'Amount1', 'Input Date', 'Input Time', 'Trade Date']].rename(columns={'Amount1':'Amount'})

    ib_fcy_irs_deal_dump = df_deal_dump[
        (df_deal_dump['Prod'].astype(str).str.strip() == 'OPT') &
        (df_deal_dump['Instrument'].astype(str).str.strip().ne('INR')) &
        (df_deal_dump['Instrument'].astype(str).str.contains('-')) &
        (df_deal_dump['Counterparty'].astype(str).str.contains('BANK', case=False, na=False))
    ]

    ib_fcy_irs_deal_dump['Input Time'] = ib_fcy_irs_deal_dump['Input Time'].astype(str).str.zfill(6)
    ib_fcy_irs_deal_dump['Formatted Time'] = ib_fcy_irs_deal_dump['Input Time'].str.slice(0,2) + ':' + \
                                        ib_fcy_irs_deal_dump['Input Time'].str.slice(2,4) + ':' + \
                                        ib_fcy_irs_deal_dump['Input Time'].str.slice(4,6)
    ib_fcy_irs_deal_dump['Input Date'] = pd.to_datetime(ib_fcy_irs_deal_dump['Input Date']).dt.date
    ib_fcy_irs_deal_dump['Trade Date'] = pd.to_datetime(ib_fcy_irs_deal_dump['Trade Date']).dt.date

    ib_fcy_irs_deal_dump['Formatted Date-Time'] = pd.to_datetime(ib_fcy_irs_deal_dump['Input Date'].astype(str) + ' ' + ib_fcy_irs_deal_dump['Formatted Time'].astype(str))
    df_ib_fcy_irs = df_ib_fcy_irs[["Mbr Rfrnce",'Txn TYPE', 'Reporting DATE TIME']]
    ib_fcy_irs_merged = pd.merge(df_ib_fcy_irs, ib_fcy_irs_deal_dump, how='inner', left_on = "Mbr Rfrnce", right_on ='System ID', indicator=True)
    ib_fcy_irs_merged.rename(columns={'_merge': 'Source'}, inplace=True)

    ib_fcy_irs_merged['Source'] = ib_fcy_irs_merged['Source'].map({
        'left_only': 'Only in IB FCY IRS',
        'right_only': 'Only in Deal Dump',
        'both': 'In Both Reports'
    })
    ib_fcy_irs_merged['Final_ID'] = ib_fcy_irs_merged["Mbr Rfrnce"].fillna(ib_fcy_irs_merged['System ID'])
    times = pd.to_datetime(ib_fcy_irs_merged['Formatted Date-Time'])

    def calculate_reporting_time(t):
        cutoff_time = time(17, 0)  # 5:00 PM

        if t.time() < cutoff_time:
            # Report by 9:00 PM same day
            reporting_time = t.replace(hour=21, minute=0, second=0)
        else:
            # Report by 10:00 AM next day
            reporting_time = (t + timedelta(days=1)).replace(hour=10, minute=0, second=0)

        return reporting_time

    # Apply function
    reporting_times = times.map(calculate_reporting_time)
    ib_fcy_irs_merged['To Be Reported within'] = pd.to_datetime(reporting_times.astype(str))
    ib_fcy_irs_merged['Reporting Date-Time as CCIL'] = pd.to_datetime(ib_fcy_irs_merged['Reporting DATE TIME'])
    df_final = pd.DataFrame({
    'System No.': ib_fcy_irs_merged['Final_ID'],
    'Product': ib_fcy_irs_merged['Prod'],
    'Instrument':ib_fcy_irs_merged['Instrument'],
    'Counterparty':ib_fcy_irs_merged['Counterparty'],
    'Trade Date': ib_fcy_irs_merged['Trade Date'], 
    'Input time': ib_fcy_irs_merged['Formatted Time'],     
    'To Be Reported within': ib_fcy_irs_merged['To Be Reported within'],
   'Date of Reporting': ib_fcy_irs_merged['Reporting DATE TIME'],
    'Reporting Date-Time as CCIL' : ib_fcy_irs_merged['Reporting Date-Time as CCIL']})

    df_final['Remarks']=df_final.apply(
        lambda row: 'PASS' if row['To Be Reported within'] > row['Reporting Date-Time as CCIL'] else 'FAIL',
        axis=1)
    
    def calculate_failure_duration(row):
        if row['Remarks'] == 'FAIL':
            # Combine Trade Date with time to create full datetime objects
            reported_dt = row['Reporting Date-Time as CCIL']
            to_be_reported_dt = row['To Be Reported within']
            
            delta = abs(reported_dt - to_be_reported_dt)
            days = delta.days
            hours, remainder = divmod(delta.seconds, 3600)
            minutes = remainder // 60
            return f"{days} days, {hours} hours, {minutes} minutes"
        else:
            return None

    df_final['Failure Duration'] = df_final.apply(calculate_failure_duration, axis=1)
    df_der_Option_FCY_INR_IB = df_final

    # Save all KPI results into one Excel file with separate sheets
    with pd.ExcelWriter(full_path_der, engine="openpyxl") as writer:
        df_der_Client_FCY_IRS.to_excel(writer, sheet_name="Derivative Client FCY IRS", index=False)
        df_der_IB_FCY_IRS.to_excel(writer, sheet_name="Derivative IB FCY IRS", index=False)
        df_der_INR_IRS.to_excel(writer, sheet_name="Derivative INR IRS", index=False)
        df_der_Option_FCY_INR_IB.to_excel(writer, sheet_name="Derivative Option FCY INR IB", index=False)

    return output_folder_path_upd