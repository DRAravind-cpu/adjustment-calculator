import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
import zipfile

st.set_page_config(page_title="Energy Adjustment Calculator", layout="wide")

# Title with author name
col1, col2 = st.columns([3, 1])
with col1:
    st.title('Energy Adjustment Calculator (Streamlit)')
with col2:
    st.markdown('<p style="text-align: right; font-weight: bold; margin-top: 20px;">Author: Er.Aravind MRT VREDC</p>', unsafe_allow_html=True)

st.markdown('Upload your generated and consumed energy Excel files, enter parameters, and download the PDF report.')

# Allow multiple file uploads
generated_files = st.file_uploader('Generated Energy Excel Files (MW)', type=['xlsx', 'xls'], accept_multiple_files=True)
consumed_files = st.file_uploader('Consumed Energy Excel Files (kWh)', type=['xlsx', 'xls'], accept_multiple_files=True)

consumer_number = st.text_input('Consumer Number')
consumer_name = st.text_input('Consumer Name')
t_and_d_loss = st.number_input('T&D Loss (%)', min_value=0.0, max_value=100.0, value=0.0, step=0.01)
multiplication_factor = st.number_input('Multiplication Factor (for Consumed Energy)', min_value=0.0, value=1.0, step=0.01)
auto_detect_month = st.checkbox('Auto-detect month and year from files', value=True)
month = st.selectbox('Month (optional)', options=[''] + [str(i) for i in range(1, 13)], format_func=lambda x: x if x == '' else pd.to_datetime(x, format='%m').strftime('%B'))
year = st.text_input('Year (e.g. 2024)')
date_filter = st.text_input('Date (optional, dd/mm/yyyy)')

# PDF output options
st.subheader('PDF Output Options')
show_excess_only = st.checkbox('Show only slots with excess (loss)', value=True)
show_all_slots = st.checkbox('Show all slots (15-min slot-wise)')
show_daywise = st.checkbox('Show day-wise summary (all days in month)')

def slot_time_to_minutes(slot_time):
    try:
        start = slot_time.split('-')[0].strip()
        h, m = map(int, start.split(':'))
        return h * 60 + m
    except Exception:
        return 0

def process_and_generate(generated_files, consumed_files, consumer_number, consumer_name, t_and_d_loss, multiplication_factor, month, year, date_filter, show_excess_only, show_all_slots, show_daywise, auto_detect_month=False):
    # Function to get month name from month number
    def get_month_name(month_num):
        month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December']
        try:
            return month_names[int(month_num) - 1]
        except (ValueError, IndexError):
            return str(month_num)
    
    # Process multiple generated energy Excel files
    gen_dfs = []
    for gen_file in generated_files:
        try:
            temp_df = pd.read_excel(gen_file, header=0)
            if temp_df.shape[1] < 3:
                return None, f"Generated energy Excel file '{gen_file.name}' must have at least 3 columns: Date, Time, and Energy in MW.", None
            
            # Add filename to help with debugging
            temp_df['Source_File'] = gen_file.name
            gen_dfs.append(temp_df)
        except Exception as e:
            return None, f"Error reading generated energy Excel file '{gen_file.name}': {str(e)}", None
    
    # Combine all generated energy dataframes
    if not gen_dfs:
        return None, "No valid generated energy Excel files were found.", None
    gen_df = pd.concat(gen_dfs, ignore_index=True)
    gen_df = gen_df.iloc[:, :3]
    gen_df.columns = ['Date', 'Time', 'Energy_MW']
    gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
    gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
    gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
    
    # Auto-detect month and year if enabled
    auto_detect_info = ""
    if auto_detect_month and not (month and year):
        # Extract unique months and years from the data
        unique_months = gen_df['Date'].dt.month.unique()
        unique_years = gen_df['Date'].dt.year.unique()
        
        if len(unique_months) == 1 and not month:
            month = str(unique_months[0])
            st.info(f"Auto-detected month: {month} ({get_month_name(month)})")
        elif len(unique_months) > 1 and not month:
            # If multiple months, use the most frequent one
            month = str(gen_df['Date'].dt.month.value_counts().idxmax())
            st.info(f"Multiple months detected, using most frequent: {month} ({get_month_name(month)})")
        
        if len(unique_years) == 1 and not year:
            year = str(unique_years[0])
            st.info(f"Auto-detected year: {year}")
        elif len(unique_years) > 1 and not year:
            # If multiple years, use the most frequent one
            year = str(gen_df['Date'].dt.year.value_counts().idxmax())
            st.info(f"Multiple years detected, using most frequent: {year}")
            
        # Add information to be displayed in PDF
        auto_detect_info = f"Auto-detected from {len(generated_files)} generated and {len(consumed_files)} consumed files"
    
    filtered_gen = gen_df.copy()
    if year and month:
        year_int = int(year)
        month_int = int(month)
        # Start: 1st of month at 00:00, End: last day of month at 23:45 (inclusive)
        start_date = pd.Timestamp(year_int, month_int, 1, 0, 0)
        if month_int == 12:
            end_date = pd.Timestamp(year_int, 12, 31, 23, 45)
        else:
            # Get the last day of the month
            if month_int == 2:
                # Handle February and leap years
                if (year_int % 4 == 0 and year_int % 100 != 0) or (year_int % 400 == 0):
                    last_day = 29
                else:
                    last_day = 28
            elif month_int in [4, 6, 9, 11]:
                last_day = 30
            else:
                last_day = 31
            end_date = pd.Timestamp(year_int, month_int, last_day, 23, 45)
        # Extract slot start time from range if needed
        def extract_start_time(t):
            t = str(t).strip()
            if '-' in t:
                return t.split('-')[0].strip()
            return t
        gen_df['Slot_Start'] = gen_df['Time'].apply(extract_start_time)
        # Combine Date and Slot_Start for filtering
        gen_df['DateTime'] = pd.to_datetime(gen_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + gen_df['Slot_Start'], errors='coerce')
        # Filter to include only slots within the same day (00:00 to 23:45)
        filtered_gen = gen_df[(gen_df['DateTime'] >= start_date) & (gen_df['DateTime'] <= end_date)]
    else:
        if year:
            year_int = int(year)
            filtered_gen = filtered_gen[filtered_gen['Date'].dt.year == year_int]
        if month:
            month_int = int(month)
            filtered_gen = filtered_gen[filtered_gen['Date'].dt.month == month_int]
    if date_filter:
        try:
            date_obj = pd.to_datetime(date_filter, dayfirst=True)
            filtered_gen = filtered_gen[filtered_gen['Date'] == date_obj]
        except Exception:
            return None, f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.", None
    if (year or month or date_filter) and filtered_gen.empty:
        # Debug output for root cause
        debug_msg = []
        debug_msg.append(f"[DEBUG] Filtered gen_df is empty after applying filters.")
        debug_msg.append(f"[DEBUG] year: {year}, month: {month}, date_filter: {date_filter}")
        available_months = ', '.join(sorted(gen_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
        debug_msg.append(f"[DEBUG] Available dates: {available_months}")
        return None, f"No data for the selected filter in the GENERATED file. Available dates: {available_months}\n\n" + '\n'.join(debug_msg), None
    
    gen_df = filtered_gen
    gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
    gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
    gen_df['After_Loss'] = gen_df['Energy_kWh'] * (1 - t_and_d_loss / 100)
    
    # Standardize slot time to 'HH:MM - HH:MM' format
    def slot_time_range(row):
        t = str(row['Time']).strip()
        if '-' in t:
            # Already in 'HH:MM - HH:MM' format
            return t
        # Accept both '0:15' and '00:15' as valid
        try:
            # Try both formats
            try:
                start = pd.to_datetime(t, format='%H:%M').time()
            except Exception:
                start = pd.to_datetime(t, format='%H:%M').time() if len(t) == 4 else pd.to_datetime(t).time()
        except Exception:
            try:
                start = pd.to_datetime(t).time()
            except Exception:
                return t
        end_dt = (pd.Timestamp.combine(pd.Timestamp.today(), start) + pd.Timedelta(minutes=15)).time()
        return f"{start.strftime('%H:%M')} - {end_dt.strftime('%H:%M')}"
    
    gen_df['Slot_Time'] = gen_df.apply(slot_time_range, axis=1)
    # Fix slot time: change '23:45 - 24:00' to '23:45 - 00:00'
    gen_df['Slot_Time'] = gen_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
    gen_df['Slot_Date'] = gen_df['Date'].dt.strftime('%d/%m/%Y')

    # Process multiple consumed energy Excel files
    cons_dfs = []
    for cons_file in consumed_files:
        try:
            temp_df = pd.read_excel(cons_file, header=0)
            if temp_df.shape[1] < 3:
                return None, f"Consumed energy Excel file '{cons_file.name}' must have at least 3 columns: Date, Time, and Energy in kWh.", None
            
            # Add filename to help with debugging
            temp_df['Source_File'] = cons_file.name
            cons_dfs.append(temp_df)
        except Exception as e:
            return None, f"Error reading consumed energy Excel file '{cons_file.name}': {str(e)}", None
    
    # Combine all consumed energy dataframes
    if not cons_dfs:
        return None, "No valid consumed energy Excel files were found.", None
    cons_df = pd.concat(cons_dfs, ignore_index=True)
    cons_df = cons_df.iloc[:, :3]
    cons_df.columns = ['Date', 'Time', 'Energy_kWh']
    cons_df['Date'] = cons_df['Date'].astype(str).str.strip()
    cons_df['Time'] = cons_df['Time'].astype(str).str.strip()
    cons_df['Date'] = pd.to_datetime(cons_df['Date'], errors='coerce', dayfirst=True)
    
    filtered_cons = cons_df.copy()
    if year and month:
        year_int = int(year)
        month_int = int(month)
        # Start: 1st of month at 00:00, End: last day of month at 23:45 (inclusive)
        start_date = pd.Timestamp(year_int, month_int, 1, 0, 0)
        if month_int == 12:
            end_date = pd.Timestamp(year_int, 12, 31, 23, 45)
        else:
            # Get the last day of the month
            if month_int == 2:
                # Handle February and leap years
                if (year_int % 4 == 0 and year_int % 100 != 0) or (year_int % 400 == 0):
                    last_day = 29
                else:
                    last_day = 28
            elif month_int in [4, 6, 9, 11]:
                last_day = 30
            else:
                last_day = 31
            end_date = pd.Timestamp(year_int, month_int, last_day, 23, 45)
        def extract_start_time(t):
            t = str(t).strip()
            if '-' in t:
                return t.split('-')[0].strip()
            return t
        cons_df['Slot_Start'] = cons_df['Time'].apply(extract_start_time)
        cons_df['DateTime'] = pd.to_datetime(cons_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + cons_df['Slot_Start'], errors='coerce')
        # Filter to include only slots within the same day (00:00 to 23:45)
        filtered_cons = cons_df[(cons_df['DateTime'] >= start_date) & (cons_df['DateTime'] <= end_date)]
    else:
        if year:
            year_int = int(year)
            filtered_cons = filtered_cons[filtered_cons['Date'].dt.year == year_int]
        if month:
            month_int = int(month)
            filtered_cons = filtered_cons[filtered_cons['Date'].dt.month == month_int]
    if date_filter:
        try:
            date_obj = pd.to_datetime(date_filter, dayfirst=True)
            filtered_cons = filtered_cons[filtered_cons['Date'] == date_obj]
        except Exception:
            return None, f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.", None
    if (year or month or date_filter) and filtered_cons.empty:
        # Debug output for root cause
        debug_msg = []
        debug_msg.append(f"[DEBUG] Filtered cons_df is empty after applying filters.")
        debug_msg.append(f"[DEBUG] year: {year}, month: {month}, date_filter: {date_filter}")
        available_months = ', '.join(sorted(cons_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
        debug_msg.append(f"[DEBUG] Available dates: {available_months}")
        return None, f"No data for the selected filter in the CONSUMED file. Available dates: {available_months}\n\n" + '\n'.join(debug_msg), None
    
    cons_df = filtered_cons
    cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce') * multiplication_factor
    cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1)
    cons_df['Slot_Time'] = cons_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
    cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')

    # Add TOD (Time of Day) classification
    def classify_tod(slot_time):
        # Extract start time from 'HH:MM - HH:MM'
        try:
            start_time = slot_time.split('-')[0].strip()
            hour, minute = map(int, start_time.split(':'))
            
            # Morning peak: 6 AM - 10 AM (C1)
            if 6 <= hour < 10:
                return 'C1'
            # Evening peak: 6 PM - 10 PM (C2)
            elif 18 <= hour < 22:
                return 'C2'
            # Normal hours: 5 AM - 6 AM + 10 AM to 6 PM (C4)
            elif (5 <= hour < 6) or (10 <= hour < 18):
                return 'C4'
            # Night hours: 10 PM to 5 AM (C5)
            else:  # 22 <= hour < 24 or 0 <= hour < 5
                return 'C5'
        except Exception:
            return 'Unknown'
    
    gen_slots_set = set((d, t) for d, t in zip(gen_df['Slot_Date'], gen_df['Slot_Time']))
    cons_slots_set = set((d, t) for d, t in zip(cons_df['Slot_Date'], cons_df['Slot_Time']))
    all_slots = pd.DataFrame(list(gen_slots_set | cons_slots_set), columns=['Slot_Date', 'Slot_Time'])
    merged = pd.merge(all_slots, gen_df[['Slot_Date', 'Slot_Time', 'After_Loss', 'Energy_kWh']], on=['Slot_Date', 'Slot_Time'], how='left')
    merged = pd.merge(merged, cons_df[['Slot_Date', 'Slot_Time', 'Energy_kWh']], on=['Slot_Date', 'Slot_Time'], how='left', suffixes=('_gen', '_cons'))
    merged['After_Loss'] = merged['After_Loss'].fillna(0)
    if 'Energy_kWh_gen' in merged.columns:
        merged['Energy_kWh_gen'] = merged['Energy_kWh_gen'].fillna(0)
    else:
        merged['Energy_kWh_gen'] = 0
    if 'Energy_kWh_cons' in merged.columns:
        merged['Energy_kWh_cons'] = merged['Energy_kWh_cons'].fillna(0)
    else:
        merged['Energy_kWh_cons'] = 0
    merged['Excess'] = merged['After_Loss'] - merged['Energy_kWh_cons']
    merged['Excess'] = merged['Excess'].apply(lambda x: x if x > 0 else 0)
    merged['Missing_Info'] = ''
    merged['is_missing_gen'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in gen_slots_set, axis=1)
    merged['is_missing_cons'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in cons_slots_set, axis=1)
    merged.loc[merged['is_missing_gen'], 'Missing_Info'] += '[Missing in GENERATED] '
    merged.loc[merged['is_missing_cons'], 'Missing_Info'] += '[Missing in CONSUMED] '
    merged.drop(['is_missing_gen', 'is_missing_cons'], axis=1, inplace=True)
    
    # Add TOD classification
    merged['TOD_Category'] = merged['Slot_Time'].apply(classify_tod)
    
    # Group excess energy by TOD category
    tod_excess = merged.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Additional charges for specific TOD categories
    c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
    c1_c2_additional = c1_c2_excess * 0.20  # 20% additional for C1 and C2
    
    c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
    c5_additional = c5_excess * 0.10  # 10% additional for C5
    
    merged['Slot_Date_dt'] = pd.to_datetime(merged['Slot_Date'], format='%d/%m/%Y', errors='coerce')
    merged['Slot_Time_min'] = merged['Slot_Time'].apply(slot_time_to_minutes)
    merged = merged.sort_values(['Slot_Date_dt', 'Slot_Time_min']).reset_index(drop=True)
    merged.drop(['Slot_Date_dt', 'Slot_Time_min'], axis=1, inplace=True)
    merged_excess = merged[merged['Excess'] > 0].copy()
    merged_all = merged.copy()

    def generate_pdf(pdf_data, sum_injection, total_generated_after_loss, comparison, total_consumed, total_excess, excess_status, filename):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Energy Adjustment Report', ln=True, align='C')
        
        # Add author name
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Author: Er.Aravind MRT VREDC', ln=True, align='R')
        
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month:
            pdf.cell(0, 10, f'Month: {month} ({get_month_name(month)})', ln=True)
        if year:
            pdf.cell(0, 10, f'Year: {year}', ln=True)
        if auto_detect_month:
            pdf.cell(0, 10, f'Note: {auto_detect_info}', ln=True)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Step-by-step Calculation:', ln=True)
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, "1. For each 15-minute slot, generated energy (MW) is converted to kWh (MW * 250).\n2. T&D loss is deducted: After_Loss = Generated_kWh * (1 - T&D loss / 100).\n3. Consumed energy is multiplied by the entered multiplication factor.\n4. For each slot, Excess = Generated_After_Loss - Consumed_Energy.\n5. The table below shows the slot-wise calculation and excess.")
        pdf.ln(2)
        
        # Define TOD descriptions for reference
        tod_descriptions = {
            'C1': 'Morning Peak',
            'C2': 'Evening Peak',
            'C4': 'Normal Hours',
            'C5': 'Night Hours',
            'Unknown': 'Unknown Time Slot'
        }
        
        # 15-min slot-wise table with improved layout
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(20, 8, 'Date', 1)
        pdf.cell(25, 8, 'Time', 1)
        pdf.cell(40, 8, 'TOD Category', 1)
        pdf.cell(25, 8, 'Gen. After Loss', 1)
        pdf.cell(25, 8, 'Consumed', 1)
        pdf.cell(25, 8, 'Excess', 1)
        pdf.cell(30, 8, 'Missing Info', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 10)
        for idx, row in pdf_data.iterrows():
            pdf.cell(20, 8, str(row['Slot_Date']), 1)
            pdf.cell(25, 8, str(row['Slot_Time']), 1)
            
            # Combine TOD category and description
            tod_cat = row.get('TOD_Category', '')
            tod_desc = tod_descriptions.get(tod_cat, '')
            pdf.cell(40, 8, f"{tod_cat} - {tod_desc}", 1)
            
            pdf.cell(25, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(25, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(25, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(30, 8, row.get('Missing_Info', ''), 1)
            pdf.ln()
        pdf.ln(2)
        
        # Add TOD-wise excess energy breakdown
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'TOD Category', 1)
        pdf.cell(60, 8, 'Description', 1)
        pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
        pdf.cell(50, 8, 'Additional Charges', 1)
        pdf.ln()
        
        # Get TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
        
        # Calculate additional charges for each TOD category
        tod_excess['Additional'] = 0.0
        tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Additional'] = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'] * 0.20
        tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Additional'] = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'] * 0.10
        
        pdf.set_font('Arial', '', 10)
        for idx, row in tod_excess.iterrows():
            category = row['TOD_Category']
            description = tod_descriptions.get(category, 'Unknown')
            
            # Add rate information to description
            if category in ['C1', 'C2']:
                description += ' (20% additional)'
            elif category == 'C5':
                description += ' (10% additional)'
            
            pdf.cell(40, 8, category, 1)
            pdf.cell(60, 8, description, 1)
            pdf.cell(40, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Additional']:.4f}", 1)
            pdf.ln()
        
        # Add total row
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'Total', 1)
        pdf.cell(60, 8, '', 1)
        pdf.cell(40, 8, f"{tod_excess['Excess'].sum():.4f}", 1)
        pdf.cell(50, 8, f"{tod_excess['Additional'].sum():.4f}", 1)
        pdf.ln(15)
        
        # Additional charges for specific TOD categories
        c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
        c1_c2_additional = c1_c2_excess * 0.20  # 20% additional for C1 and C2
        
        c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
        c5_additional = c5_excess * 0.10  # 10% additional for C5
        
        # Summary section
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Summary:', ln=True)
        pdf.set_font('Arial', '', 11)
        pdf.cell(0, 8, f'Sum of Injection (Generated before Loss, kWh): {pdf_data["Energy_kWh_gen"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Generated Energy after Loss (kWh): {pdf_data["After_Loss"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Consumed Energy (kWh, after multiplication): {pdf_data["Energy_kWh_cons"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Excess Energy (kWh): {pdf_data["Excess"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Morning & Evening Peak Additional (C1+C2, 20%): {c1_c2_additional:.4f}', ln=True)
        pdf.cell(0, 8, f'Night Hours Additional (C5, 10%): {c5_additional:.4f}', ln=True)
        pdf.cell(0, 8, f'Total Additional Charges: {(c1_c2_additional + c5_additional):.4f}', ln=True)
        pdf.cell(0, 8, f'Grand Total (Excess + Additional): {(pdf_data["Excess"].sum() + c1_c2_additional + c5_additional):.4f}', ln=True)
        
        pdf_output = io.BytesIO()
        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output.write(pdf_bytes)
        pdf_output.seek(0)
        return pdf_output, filename

    def generate_daywise_pdf(pdf_data, month, year, filename):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Energy Adjustment Day-wise Summary', ln=True, align='C')
        
        # Add author name
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Author: Er.Aravind MRT VREDC', ln=True, align='R')
        
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month:
            pdf.cell(0, 10, f'Month: {month} ({get_month_name(month)})', ln=True)
        if year:
            pdf.cell(0, 10, f'Year: {year}', ln=True)
        if auto_detect_month:
            pdf.cell(0, 10, f'Note: {auto_detect_info}', ln=True)
        pdf.ln(5)
        
        # Day-wise summary table
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Day-wise Summary (All Days in Month):', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(30, 8, 'Date', 1)
        pdf.cell(40, 8, 'Total Gen. After Loss', 1)
        pdf.cell(40, 8, 'Total Consumed', 1)
        pdf.cell(40, 8, 'Total Excess', 1)
        pdf.cell(40, 8, 'Additional Charges', 1)
        pdf.ln()
        
        pdf.set_font('Arial', '', 10)
        from datetime import datetime, timedelta
        if month and year:
            month_int = int(month)
            year_int = int(year)
            start_date = datetime(year_int, month_int, 1)
            if month_int == 12:
                end_date = datetime(year_int + 1, 1, 1)
            else:
                end_date = datetime(year_int, month_int + 1, 1)
        else:
            all_dates = pd.to_datetime(pdf_data['Slot_Date'], dayfirst=True, errors='coerce')
            start_date = all_dates.min()
            end_date = all_dates.max() + timedelta(days=1)
        
        all_days = pd.date_range(start=start_date, end=end_date - timedelta(days=1), freq='D')
        
        # Group by date and TOD category to calculate additional charges
        daywise_tod = pdf_data.groupby(['Slot_Date', 'TOD_Category']).agg({
            'Excess': 'sum'
        }).reset_index()
        
        # Calculate additional charges for each TOD category by day
        daywise_tod['Additional'] = 0.0
        daywise_tod.loc[daywise_tod['TOD_Category'].isin(['C1', 'C2']), 'Additional'] = daywise_tod.loc[daywise_tod['TOD_Category'].isin(['C1', 'C2']), 'Excess'] * 0.20
        daywise_tod.loc[daywise_tod['TOD_Category'] == 'C5', 'Additional'] = daywise_tod.loc[daywise_tod['TOD_Category'] == 'C5', 'Excess'] * 0.10
        
        # Sum additional charges by day
        additional_by_day = daywise_tod.groupby('Slot_Date')['Additional'].sum().reset_index()
        
        # Regular day-wise aggregation
        daywise = pdf_data.groupby('Slot_Date').agg({
            'After_Loss': 'sum',
            'Energy_kWh_cons': 'sum',
            'Excess': 'sum'
        })
        
        daywise = daywise.reindex(all_days.strftime('%d/%m/%Y'), fill_value=0).reset_index()
        daywise = daywise.rename(columns={'index': 'Slot_Date'})
        
        # Merge with additional charges
        daywise = pd.merge(daywise, additional_by_day, on='Slot_Date', how='left')
        daywise['Additional'] = daywise['Additional'].fillna(0)
        
        for idx, row in daywise.iterrows():
            pdf.cell(30, 8, row['Slot_Date'], 1)
            pdf.cell(40, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(40, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(40, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(40, 8, f"{row['Additional']:.4f}", 1)
            pdf.ln()
        
        # Add total row
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(30, 8, 'Total', 1)
        pdf.cell(40, 8, f"{daywise['After_Loss'].sum():.4f}", 1)
        pdf.cell(40, 8, f"{daywise['Energy_kWh_cons'].sum():.4f}", 1)
        pdf.cell(40, 8, f"{daywise['Excess'].sum():.4f}", 1)
        pdf.cell(40, 8, f"{daywise['Additional'].sum():.4f}", 1)
        pdf.ln(15)
        
        # TOD-wise summary
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'TOD Category', 1)
        pdf.cell(60, 8, 'Description', 1)
        pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
        pdf.cell(50, 8, 'Additional Charges', 1)
        pdf.ln()
        
        # Define TOD descriptions
        tod_descriptions = {
            'C1': 'Morning Peak',
            'C2': 'Evening Peak',
            'C4': 'Normal Hours',
            'C5': 'Night Hours',
            'Unknown': 'Unknown Time Slot'
        }
        
        # Get TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
        
        # Calculate additional charges for each TOD category
        tod_excess['Additional'] = 0.0
        tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Additional'] = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'] * 0.20
        tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Additional'] = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'] * 0.10
        
        pdf.set_font('Arial', '', 10)
        for idx, row in tod_excess.iterrows():
            category = row['TOD_Category']
            description = tod_descriptions.get(category, 'Unknown')
            
            # Add rate information to description
            if category in ['C1', 'C2']:
                description += ' (20% additional)'
            elif category == 'C5':
                description += ' (10% additional)'
            
            pdf.cell(40, 8, category, 1)
            pdf.cell(60, 8, description, 1)
            pdf.cell(40, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Additional']:.4f}", 1)
            pdf.ln()
        
        # Add total row
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'Total', 1)
        pdf.cell(60, 8, '', 1)
        pdf.cell(40, 8, f"{tod_excess['Excess'].sum():.4f}", 1)
        pdf.cell(50, 8, f"{tod_excess['Additional'].sum():.4f}", 1)
        pdf.ln(15)
        
        # Additional charges for specific TOD categories
        c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
        c1_c2_additional = c1_c2_excess * 0.20  # 20% additional for C1 and C2
        
        c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
        c5_additional = c5_excess * 0.10  # 10% additional for C5
        
        # Summary section
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Summary:', ln=True)
        pdf.set_font('Arial', '', 11)
        pdf.cell(0, 8, f'Sum of Injection (Generated before Loss, kWh): {pdf_data["Energy_kWh_gen"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Generated Energy after Loss (kWh): {pdf_data["After_Loss"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Consumed Energy (kWh, after multiplication): {pdf_data["Energy_kWh_cons"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Total Excess Energy (kWh): {pdf_data["Excess"].sum():.4f}', ln=True)
        pdf.cell(0, 8, f'Morning & Evening Peak Additional (C1+C2, 20%): {c1_c2_additional:.4f}', ln=True)
        pdf.cell(0, 8, f'Night Hours Additional (C5, 10%): {c5_additional:.4f}', ln=True)
        pdf.cell(0, 8, f'Total Additional Charges: {(c1_c2_additional + c5_additional):.4f}', ln=True)
        pdf.cell(0, 8, f'Grand Total (Excess + Additional): {(pdf_data["Excess"].sum() + c1_c2_additional + c5_additional):.4f}', ln=True)
        
        pdf_output = io.BytesIO()
        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output.write(pdf_bytes)
        pdf_output.seek(0)
        return pdf_output, filename

    pdfs = []
    if show_excess_only:
        pdf_obj, fname = generate_pdf(merged_excess, merged_excess['Energy_kWh_gen'].sum(), merged_excess['After_Loss'].sum(), merged_excess['Energy_kWh_gen'].sum() - merged_excess['After_Loss'].sum(), merged_excess['Energy_kWh_cons'].sum(), merged_excess['Excess'].sum(), 'Excess' if merged_excess['Excess'].sum() > 0 else 'No Excess', 'energy_adjustment_excess_only.pdf')
        pdfs.append((fname, pdf_obj))
    if show_all_slots:
        pdf_obj, fname = generate_pdf(merged_all, merged_all['Energy_kWh_gen'].sum(), merged_all['After_Loss'].sum(), merged_all['Energy_kWh_gen'].sum() - merged_all['After_Loss'].sum(), merged_all['Energy_kWh_cons'].sum(), merged_all['Excess'].sum(), 'Excess' if merged_all['Excess'].sum() > 0 else 'No Excess', 'energy_adjustment_all_slots.pdf')
        pdfs.append((fname, pdf_obj))
    if show_daywise:
        pdf_obj, fname = generate_daywise_pdf(merged_all, month, year, 'energy_adjustment_daywise.pdf')
        pdfs.append((fname, pdf_obj))
    if len(pdfs) == 1:
        return pdfs[0][1].getvalue(), pdfs[0][0], None
    elif len(pdfs) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            for fname, pdf_io in pdfs:
                zf.writestr(fname, pdf_io.getvalue())
        zip_buffer.seek(0)
        return zip_buffer.getvalue(), 'energy_adjustment_reports.zip', 'application/zip'
    else:
        return None, None, None

if st.button('Generate PDF Report'):
    if not generated_files or not consumed_files:
        st.error('Please upload both generated and consumed Excel files.')
    elif not consumer_number or not consumer_name:
        st.error('Please fill in all required fields (Consumer Number and Consumer Name).')
    elif not year and not auto_detect_month:
        st.error('Please either enter a year or enable auto-detect month and year.')
    else:
        with st.spinner('Processing...'):
            file_bytes, file_name, mime = process_and_generate(
                generated_files, consumed_files, consumer_number, consumer_name, 
                t_and_d_loss, multiplication_factor, month, year, date_filter, 
                show_excess_only, show_all_slots, show_daywise, auto_detect_month
            )
            if file_bytes and file_name:
                st.success('Report generated!')
                st.download_button('Download', file_bytes, file_name=file_name, mime=mime or 'application/pdf')
            else:
                if isinstance(file_name, str):
                    st.error(file_name)  # Display the error message
                else:
                    st.error('Failed to generate report. Please check your input data.')
