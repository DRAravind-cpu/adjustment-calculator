import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
import zipfile
from datetime import datetime, timedelta
import traceback

st.set_page_config(page_title="Energy Adjustment Calculator", layout="wide")

# Title with author name
col1, col2 = st.columns([3, 1])
with col1:
    st.title('Energy Adjustment Calculator (Streamlit)')
with col2:
    st.markdown('<p style="text-align: right; font-weight: bold; margin-top: 20px;">Author: Er.Aravind MRT VREDC</p>', unsafe_allow_html=True)

st.markdown('Upload your generated and consumed energy Excel files, enter parameters, and download the PDF report.')

# --- UI Elements ---
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

def process_and_generate(generated_files, consumed_files, consumer_number, consumer_name, t_and_d_loss, multiplication_factor, month, year, date_filter, show_excess_only, show_all_slots, show_daywise, auto_detect_month=False):
    
    # --- Helper Functions (from app.py) ---
    def get_month_name(month_num):
        month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        try:
            return month_names[int(month_num) - 1]
        except (ValueError, IndexError):
            return str(month_num)

    def slot_time_to_minutes(slot_time):
        try:
            start = slot_time.split('-')[0].strip()
            h, m = map(int, start.split(':'))
            return h * 60 + m
        except Exception:
            return 0
            
    def safe_date_str(d):
        if pd.isnull(d):
            return ''
        if isinstance(d, str):
            return d
        return d.strftime('%d/%m/%Y')

    # --- Data Processing (from app.py) ---
    # Process multiple generated energy Excel files
    gen_dfs = []
    for gen_file in generated_files:
        try:
            temp_df = pd.read_excel(gen_file, header=0)
            if temp_df.shape[1] < 3:
                return None, f"Generated energy Excel file '{gen_file.name}' must have at least 3 columns: Date, Time, and Energy in MW.", None
            temp_df['Source_File'] = gen_file.name
            gen_dfs.append(temp_df)
        except Exception as e:
            return None, f"Error reading generated energy Excel file '{gen_file.name}': {str(e)}", None
    
    if not gen_dfs:
        return None, "No valid generated energy Excel files were found.", None
    gen_df = pd.concat(gen_dfs, ignore_index=True)
    gen_df = gen_df.iloc[:, :3]
    gen_df.columns = ['Date', 'Time', 'Energy_MW']
    gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
    gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
    gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
    
    auto_detect_info = ""
    if auto_detect_month and not (month and year):
        unique_months = gen_df['Date'].dt.month.unique()
        unique_years = gen_df['Date'].dt.year.unique()
        if len(unique_months) == 1 and not month:
            month = str(unique_months[0])
        elif len(unique_months) > 1 and not month:
            month = str(gen_df['Date'].dt.month.value_counts().idxmax())
        if len(unique_years) == 1 and not year:
            # Convert to integer first to remove any decimal part
            year = str(int(unique_years[0]))
        elif len(unique_years) > 1 and not year:
            # Convert to integer first to remove any decimal part
            year = str(int(gen_df['Date'].dt.year.value_counts().idxmax()))
        auto_detect_info = f"Auto-detected from {len(generated_files)} generated and {len(consumed_files)} consumed files"
        st.info(f"Auto-detected Month: {get_month_name(month)}, Year: {year}")

    filtered_gen = gen_df.copy()
    if year and month:
        try:
            # Handle potential float strings like '2025.0'
            year_int = int(float(year))
            month_int = int(float(month))
            start_date = pd.Timestamp(year_int, month_int, 1)
        except ValueError as e:
            st.error(f"Invalid year or month format: {e}")
            return None, f"Error: Invalid year '{year}' or month '{month}' format", None
        end_of_month = start_date + pd.offsets.MonthEnd(1)
        end_date = pd.Timestamp(end_of_month.year, end_of_month.month, end_of_month.day, 23, 45)

        def extract_start_time(t):
            t = str(t).strip()
            return t.split('-')[0].strip() if '-' in t else t
        
        gen_df['Slot_Start'] = gen_df['Time'].apply(extract_start_time)
        gen_df['DateTime'] = pd.to_datetime(gen_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + gen_df['Slot_Start'], errors='coerce')
        filtered_gen = gen_df[(gen_df['DateTime'] >= start_date) & (gen_df['DateTime'] <= end_date)]

    if date_filter:
        try:
            date_obj = pd.to_datetime(date_filter, dayfirst=True)
            filtered_gen = filtered_gen[filtered_gen['Date'] == date_obj]
        except Exception:
            return None, f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.", None

    if (year or month or date_filter) and filtered_gen.empty:
        available_dates = ', '.join(sorted(gen_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
        return None, f"No data for the selected filter in GENERATED files. Available dates: {available_dates}", None

    gen_df = filtered_gen
    gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
    gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
    gen_df['After_Loss'] = gen_df['Energy_kWh'] * (1 - t_and_d_loss / 100)

    def slot_time_range(row):
        t = str(row['Time']).strip()
        if '-' in t: return t
        try:
            start = pd.to_datetime(t, format='%H:%M').time()
        except Exception:
            try:
                start = pd.to_datetime(t).time()
            except Exception:
                return t
        end_dt = (pd.Timestamp.combine(pd.Timestamp.today(), start) + pd.Timedelta(minutes=15)).time()
        return f"{start.strftime('%H:%M')} - {end_dt.strftime('%H:%M')}"

    gen_df['Slot_Time'] = gen_df.apply(slot_time_range, axis=1).replace({'23:45 - 24:00': '23:45 - 00:00'})
    gen_df['Slot_Date'] = gen_df['Date'].dt.strftime('%d/%m/%Y')

    # Process consumed files
    cons_dfs = []
    for cons_file in consumed_files:
        try:
            temp_df = pd.read_excel(cons_file, header=0)
            if temp_df.shape[1] < 3:
                return None, f"Consumed energy Excel file '{cons_file.name}' must have at least 3 columns.", None
            temp_df['Source_File'] = cons_file.name
            cons_dfs.append(temp_df)
        except Exception as e:
            return None, f"Error reading consumed energy Excel file '{cons_file.name}': {str(e)}", None

    if not cons_dfs:
        return None, "No valid consumed energy Excel files were found.", None
    cons_df = pd.concat(cons_dfs, ignore_index=True)
    cons_df = cons_df.iloc[:, :3]
    cons_df.columns = ['Date', 'Time', 'Energy_kWh']
    cons_df['Date'] = pd.to_datetime(cons_df['Date'].astype(str).str.strip(), errors='coerce', dayfirst=True)
    cons_df['Time'] = cons_df['Time'].astype(str).str.strip()

    filtered_cons = cons_df.copy()
    if year and month:
        try:
            # Handle potential float strings like '2025.0'
            year_int = int(float(year))
            month_int = int(float(month))
            start_date = pd.Timestamp(year_int, month_int, 1)
        except ValueError as e:
            st.error(f"Invalid year or month format: {e}")
            return None, f"Error: Invalid year '{year}' or month '{month}' format", None
        end_of_month = start_date + pd.offsets.MonthEnd(1)
        end_date = pd.Timestamp(end_of_month.year, end_of_month.month, end_of_month.day, 23, 45)
        
        def extract_start_time(t):
            t = str(t).strip()
            return t.split('-')[0].strip() if '-' in t else t

        cons_df['Slot_Start'] = cons_df['Time'].apply(extract_start_time)
        cons_df['DateTime'] = pd.to_datetime(cons_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + cons_df['Slot_Start'], errors='coerce')
        filtered_cons = cons_df[(cons_df['DateTime'] >= start_date) & (cons_df['DateTime'] <= end_date)]

    if date_filter:
        try:
            date_obj = pd.to_datetime(date_filter, dayfirst=True)
            filtered_cons = filtered_cons[filtered_cons['Date'] == date_obj]
        except Exception:
            return None, f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.", None

    if (year or month or date_filter) and filtered_cons.empty:
        available_dates = ', '.join(sorted(cons_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
        return None, f"No data for the selected filter in CONSUMED files. Available dates: {available_dates}", None

    cons_df = filtered_cons
    cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce') * multiplication_factor
    cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1).replace({'23:45 - 24:00': '23:45 - 00:00'})
    cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')

    # Merge data
    gen_slots_set = set(zip(gen_df['Slot_Date'], gen_df['Slot_Time']))
    cons_slots_set = set(zip(cons_df['Slot_Date'], cons_df['Slot_Time']))
    all_slots = pd.DataFrame(list(gen_slots_set | cons_slots_set), columns=['Slot_Date', 'Slot_Time'])
    
    merged = pd.merge(all_slots, gen_df[['Slot_Date', 'Slot_Time', 'After_Loss', 'Energy_kWh']], on=['Slot_Date', 'Slot_Time'], how='left')
    merged = pd.merge(merged, cons_df[['Slot_Date', 'Slot_Time', 'Energy_kWh']], on=['Slot_Date', 'Slot_Time'], how='left', suffixes=('_gen', '_cons'))
    
    merged.fillna(0, inplace=True)
    merged['Excess'] = (merged['After_Loss'] - merged['Energy_kWh_cons']).clip(lower=0)
    
    merged['Missing_Info'] = ''
    merged['is_missing_gen'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in gen_slots_set, axis=1)
    merged['is_missing_cons'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in cons_slots_set, axis=1)
    merged.loc[merged['is_missing_gen'], 'Missing_Info'] += '[No Gen] '
    merged.loc[merged['is_missing_cons'], 'Missing_Info'] += '[No Cons] '
    merged.drop(['is_missing_gen', 'is_missing_cons'], axis=1, inplace=True)

    merged['Slot_Date_dt'] = pd.to_datetime(merged['Slot_Date'], format='%d/%m/%Y', errors='coerce')
    merged['Slot_Time_min'] = merged['Slot_Time'].apply(slot_time_to_minutes)
    merged = merged.sort_values(['Slot_Date_dt', 'Slot_Time_min']).reset_index(drop=True)
    
    def classify_tod(slot_time):
        try:
            start_time = slot_time.split('-')[0].strip()
            hour = int(start_time.split(':')[0])
            if 6 <= hour < 10: return 'C1'
            if 18 <= hour < 22: return 'C2'
            if (5 <= hour < 6) or (10 <= hour < 18): return 'C4'
            return 'C5'
        except:
            return 'Unknown'
    merged['TOD_Category'] = merged['Slot_Time'].apply(classify_tod)
    
    # Financial Calculations
    total_excess = merged['Excess'].sum()
    tod_excess = merged.groupby('TOD_Category')['Excess'].sum()
    
    base_rate = 7.25
    base_amount = total_excess * base_rate
    
    c1_c2_excess = tod_excess.get('C1', 0) + tod_excess.get('C2', 0)
    c1_c2_additional = c1_c2_excess * 1.81
    
    c5_excess = tod_excess.get('C5', 0)
    c5_additional = c5_excess * 0.3625
    
    total_amount = base_amount + c1_c2_additional + c5_additional
    etax = total_amount * 0.05
    total_with_etax = total_amount + etax
    
    etax_on_iex = total_excess * 0.1
    cross_subsidy_surcharge = total_excess * 1.92
    
    loss_percentage = t_and_d_loss
    if 100 - loss_percentage > 0:
        x_value = (total_excess * loss_percentage) / (100 - loss_percentage)
        y_value = (total_excess * loss_percentage / x_value) + total_excess if x_value != 0 else total_excess
        wheeling_charges = y_value * 1.04
    else:
        wheeling_charges = 0

    final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)

    merged.drop(['Slot_Date_dt', 'Slot_Time_min'], axis=1, inplace=True)
    
    # Totals for PDF
    sum_injection = merged['Energy_kWh_gen'].sum()
    total_generated_after_loss = merged['After_Loss'].sum()
    total_consumed = merged['Energy_kWh_cons'].sum()
    comparison = sum_injection - total_generated_after_loss
    excess_status = 'Excess' if total_excess > 0 else 'No Excess'
    unique_days_gen_full = gen_df['Slot_Date'].nunique()
    unique_days_cons_full = cons_df['Slot_Date'].nunique()

    # --- PDF Generation Functions ---
    def generate_pdf(pdf_data, filename, use_total_consumed=None):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Energy Adjustment Report', ln=True, align='C')
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month: pdf.cell(0, 10, f'Month: {get_month_name(month)}', ln=True)
        if year: pdf.cell(0, 10, f'Year: {year}', ln=True)
        if auto_detect_info:
            pdf.set_font('Arial', 'I', 10)
            pdf.cell(0, 10, auto_detect_info, ln=True)
        
        # Table
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(20, 8, 'Date', 1)
        pdf.cell(25, 8, 'Time', 1)
        pdf.cell(60, 8, 'TOD Category & Description', 1)
        pdf.cell(25, 8, 'Gen. After Loss', 1)
        pdf.cell(25, 8, 'Consumed', 1)
        pdf.cell(25, 8, 'Excess', 1)
        pdf.cell(10, 8, 'Miss', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 9)
        
        tod_descriptions = {'C1': 'Morning Peak', 'C2': 'Evening Peak', 'C4': 'Normal Hours', 'C5': 'Night Hours', 'Unknown': 'Unknown'}
        for _, row in pdf_data.iterrows():
            pdf.cell(20, 8, safe_date_str(row['Slot_Date']), 1)
            pdf.cell(25, 8, str(row['Slot_Time']), 1)
            tod_cat = row.get('TOD_Category', '')
            pdf.cell(60, 8, f"{tod_cat}: {tod_descriptions.get(tod_cat, '')}", 1)
            pdf.cell(25, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(25, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(25, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(10, 8, row.get('Missing_Info', ''), 1)
            pdf.ln()

        # Summary
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        # If use_total_consumed is provided, use it instead of calculating from pdf_data
        # This ensures consistency across all PDFs
        displayed_total_consumed = use_total_consumed if use_total_consumed is not None else pdf_data['Energy_kWh_cons'].sum()
        
        # Complete calculation details matching the Flask app
        pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {sum_injection:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {total_generated_after_loss:.4f}', ln=True)
        pdf.cell(0, 10, f'Comparison (Injection - After Loss): {comparison:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {displayed_total_consumed:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Excess Energy (kWh): {total_excess:.4f}', ln=True)
        pdf.cell(0, 10, f'Unique Days Used (Generated): {unique_days_gen_full}', ln=True)
        pdf.cell(0, 10, f'Unique Days Used (Consumed): {unique_days_cons_full}', ln=True)
        pdf.cell(0, 10, f'Status: {excess_status}', ln=True)
        
        # Add TOD-wise excess energy breakdown
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'TOD Category', 1)
        pdf.cell(40, 8, 'Description', 1)
        pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 10)
        
        # Get TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
        
        # TOD descriptions with time ranges removed for better readability
        tod_descriptions = {
            'C1': 'Morning Peak',
            'C2': 'Evening Peak',
            'C4': 'Normal Hours',
            'C5': 'Night Hours',
            'Unknown': 'Unknown Time Slot'
        }
        
        for _, row in tod_excess.iterrows():
            category = row['TOD_Category']
            description = tod_descriptions.get(category, 'Unknown')
            excess = row['Excess']
            pdf.cell(40, 8, category, 1)
            pdf.cell(40, 8, description, 1)
            pdf.cell(40, 8, f"{excess:.4f}", 1)
            pdf.ln()
        
        # Financials
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Financial Calculations:', ln=True)
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 8, f"1. Base Rate: {total_excess:.4f} kWh x Rs.{base_rate:.2f} = Rs.{base_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"2. C1+C2 Additional: {c1_c2_excess:.4f} kWh x Rs.1.81 = Rs.{c1_c2_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"3. C5 Additional: {c5_excess:.4f} kWh x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"4. Total Amount: Rs.{total_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"5. E-Tax (5%): Rs.{etax:.2f}", ln=True)
        pdf.cell(0, 8, f"6. Total with E-Tax: Rs.{total_with_etax:.2f}", ln=True)
        pdf.cell(0, 8, f"7. E-Tax on IEX: Rs.{etax_on_iex:.2f}", ln=True)
        pdf.cell(0, 8, f"8. Cross Subsidy Surcharge: Rs.{cross_subsidy_surcharge:.2f}", ln=True)
        pdf.cell(0, 8, f"9. Wheeling Charges: Rs.{wheeling_charges:.2f}", ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 10, f"10. Total Amount to be Collected: Rs.{final_amount:.2f}", ln=True)

        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output = io.BytesIO(pdf_bytes)
        return pdf_output

    def generate_daywise_pdf(pdf_data, filename):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Energy Adjustment Day-wise Summary', ln=True, align='C')
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month: pdf.cell(0, 10, f'Month: {get_month_name(month)}', ln=True)
        if year: pdf.cell(0, 10, f'Year: {year}', ln=True)

        # Day-wise table
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'Date', 1)
        pdf.cell(50, 8, 'Total Gen. After Loss', 1)
        pdf.cell(50, 8, 'Total Consumed', 1)
        pdf.cell(50, 8, 'Total Excess', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 10)

        daywise = pdf_data.groupby('Slot_Date').agg({
            'After_Loss': 'sum',
            'Energy_kWh_cons': 'sum',
            'Excess': 'sum'
        }).reset_index()
        
        for _, row in daywise.iterrows():
            pdf.cell(40, 8, row['Slot_Date'], 1)
            pdf.cell(50, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Excess']:.4f}", 1)
            pdf.ln()
        
        # Summary and Financials (same as other PDF)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        # Complete calculation details matching the Flask app
        pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {sum_injection:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {total_generated_after_loss:.4f}', ln=True)
        pdf.cell(0, 10, f'Comparison (Injection - After Loss): {comparison:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {total_consumed:.4f}', ln=True)
        pdf.cell(0, 10, f'Total Excess Energy (kWh): {total_excess:.4f}', ln=True)
        pdf.cell(0, 10, f'Unique Days Used (Generated): {unique_days_gen_full}', ln=True)
        pdf.cell(0, 10, f'Unique Days Used (Consumed): {unique_days_cons_full}', ln=True)
        pdf.cell(0, 10, f'Status: {excess_status}', ln=True)
        
        # Add TOD-wise excess energy breakdown
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'TOD Category', 1)
        pdf.cell(40, 8, 'Description', 1)
        pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 10)
        
        # Get TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
        
        # TOD descriptions with time ranges removed for better readability
        tod_descriptions = {
            'C1': 'Morning Peak',
            'C2': 'Evening Peak',
            'C4': 'Normal Hours',
            'C5': 'Night Hours',
            'Unknown': 'Unknown Time Slot'
        }
        
        for _, row in tod_excess.iterrows():
            category = row['TOD_Category']
            description = tod_descriptions.get(category, 'Unknown')
            excess = row['Excess']
            pdf.cell(40, 8, category, 1)
            pdf.cell(40, 8, description, 1)
            pdf.cell(40, 8, f"{excess:.4f}", 1)
            pdf.ln()
        
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Financial Calculations:', ln=True)
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 8, f"1. Base Rate: {total_excess:.4f} kWh x Rs.{base_rate:.2f} = Rs.{base_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"2. C1+C2 Additional: {c1_c2_excess:.4f} kWh x Rs.1.81 = Rs.{c1_c2_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"3. C5 Additional: {c5_excess:.4f} kWh x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"4. Total Amount: Rs.{total_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"5. E-Tax (5%): Rs.{etax:.2f}", ln=True)
        pdf.cell(0, 8, f"6. Total with E-Tax: Rs.{total_with_etax:.2f}", ln=True)
        pdf.cell(0, 8, f"7. E-Tax on IEX: Rs.{etax_on_iex:.2f}", ln=True)
        pdf.cell(0, 8, f"8. Cross Subsidy Surcharge: Rs.{cross_subsidy_surcharge:.2f}", ln=True)
        pdf.cell(0, 8, f"9. Wheeling Charges: Rs.{wheeling_charges:.2f}", ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 10, f"10. Total Amount to be Collected: Rs.{final_amount:.2f}", ln=True)

        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output = io.BytesIO(pdf_bytes)
        return pdf_output

    # --- PDF Generation Logic ---
    pdfs = []
    if not (show_excess_only or show_all_slots or show_daywise):
        show_all_slots = True # Default option
    
    if show_excess_only:
        try:
            # Create a filtered dataframe for display but use the total consumed energy from all slots
            merged_excess = merged[merged['Excess'] > 0].copy()
            # We'll pass the filtered data but the total_consumed will be from all slots
            pdf_obj = generate_pdf(merged_excess, 'energy_adjustment_excess_only.pdf', use_total_consumed=total_consumed)
            pdfs.append(('energy_adjustment_excess_only.pdf', pdf_obj))
        except Exception as e:
            st.error(f"Error generating 'Excess Only' PDF: {e}")
            st.error(traceback.format_exc())

    if show_all_slots:
        try:
            # For all slots, we use the same total_consumed value for consistency
            pdf_obj = generate_pdf(merged, 'energy_adjustment_all_slots.pdf', use_total_consumed=total_consumed)
            pdfs.append(('energy_adjustment_all_slots.pdf', pdf_obj))
        except Exception as e:
            st.error(f"Error generating 'All Slots' PDF: {e}")
            st.error(traceback.format_exc())

    if show_daywise:
        try:
            # For daywise PDF, we also use the same total_consumed value for consistency
            pdf_obj = generate_daywise_pdf(merged, 'energy_adjustment_daywise.pdf')
            pdfs.append(('energy_adjustment_daywise.pdf', pdf_obj))
        except Exception as e:
            st.error(f"Error generating 'Day-wise' PDF: {e}")
            st.error(traceback.format_exc())

    if not pdfs:
        return None, "No PDF was generated due to errors or no option selected.", None

    if len(pdfs) == 1:
        fname, pdf_io = pdfs[0]
        return pdf_io.getvalue(), fname, 'application/pdf'
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zf:
            for fname, pdf_io in pdfs:
                zf.writestr(fname, pdf_io.getvalue())
        zip_buffer.seek(0)
        return zip_buffer.getvalue(), 'energy_adjustment_reports.zip', 'application/zip'

# --- Button to trigger processing ---
if st.button('Generate PDF Report'):
    if not generated_files or not consumed_files:
        st.error('Please upload both generated and consumed Excel files.')
    elif not consumer_number or not consumer_name:
        st.error('Please fill in all required fields (Consumer Number and Consumer Name).')
    elif not (year and month) and not auto_detect_month:
        st.error('Please either enter a month and year or enable auto-detect.')
    else:
        with st.spinner('Processing...'):
            try:
                file_bytes, file_name, mime = process_and_generate(
                    generated_files, consumed_files, consumer_number, consumer_name, 
                    t_and_d_loss, multiplication_factor, month, year, date_filter, 
                    show_excess_only, show_all_slots, show_daywise, auto_detect_month
                )
                if file_bytes and file_name:
                    st.success('Report generated!')
                    st.download_button('Download Report', file_bytes, file_name=file_name, mime=mime)
                else:
                    st.error(file_name or 'An unknown error occurred during PDF generation.')
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")
                st.error(traceback.format_exc())
