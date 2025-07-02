import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
import zipfile

st.title('Energy Adjustment Calculator (Streamlit)')

st.markdown('Upload your generated and consumed energy Excel files, enter parameters, and download the PDF report.')

generated_file = st.file_uploader('Generated Energy Excel (MW)', type=['xlsx', 'xls'])
consumed_file = st.file_uploader('Consumed Energy Excel (kWh)', type=['xlsx', 'xls'])

consumer_number = st.text_input('Consumer Number')
consumer_name = st.text_input('Consumer Name')
t_and_d_loss = st.number_input('T&D Loss (%)', min_value=0.0, max_value=100.0, value=0.0, step=0.01)
multiplication_factor = st.number_input('Multiplication Factor (for Consumed Energy)', min_value=0.0, value=1.0, step=0.01)
month = st.selectbox('Month (optional)', options=[''] + [str(i) for i in range(1, 13)], format_func=lambda x: x if x == '' else pd.to_datetime(x, format='%m').strftime('%B'))
year = st.text_input('Year (e.g. 2024)')
date_filter = st.text_input('Date (optional, dd/mm/yyyy)')

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

def process_and_generate(generated_file, consumed_file, consumer_number, consumer_name, t_and_d_loss, multiplication_factor, month, year, date_filter, show_excess_only, show_all_slots, show_daywise):
    gen_df = pd.read_excel(generated_file, header=0)
    gen_df = gen_df.iloc[:, :3]
    gen_df.columns = ['Date', 'Time', 'Energy_MW']
    gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
    gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
    gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
    filtered_gen = gen_df.copy()
    if year and month:
        year_int = int(year)
        month_int = int(month)
        start_date = pd.Timestamp(year_int, month_int, 1, 0, 15)
        if month_int == 12:
            end_date = pd.Timestamp(year_int + 1, 1, 1, 0, 0)
        else:
            end_date = pd.Timestamp(year_int, month_int + 1, 1, 0, 0)
        def extract_start_time(t):
            t = str(t).strip()
            if '-' in t:
                return t.split('-')[0].strip()
            return t
        gen_df['Slot_Start'] = gen_df['Time'].apply(extract_start_time)
        gen_df['DateTime'] = pd.to_datetime(gen_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + gen_df['Slot_Start'], errors='coerce')
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
            return None, 'Invalid date format for filter: {}. Use dd/mm/yyyy.'.format(date_filter)
    gen_df = filtered_gen
    gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
    gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
    gen_df['After_Loss'] = gen_df['Energy_kWh'] * (1 - t_and_d_loss / 100)
    def slot_time_range(row):
        t = str(row['Time']).strip()
        if '-' in t:
            return t
        try:
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
    gen_df['Slot_Time'] = gen_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
    gen_df['Slot_Date'] = gen_df['Date'].dt.strftime('%d/%m/%Y')

    cons_df = pd.read_excel(consumed_file, header=0)
    cons_df = cons_df.iloc[:, :3]
    cons_df.columns = ['Date', 'Time', 'Energy_kWh']
    cons_df['Date'] = cons_df['Date'].astype(str).str.strip()
    cons_df['Time'] = cons_df['Time'].astype(str).str.strip()
    cons_df['Date'] = pd.to_datetime(cons_df['Date'], errors='coerce', dayfirst=True)
    filtered_cons = cons_df.copy()
    if year and month:
        year_int = int(year)
        month_int = int(month)
        start_date = pd.Timestamp(year_int, month_int, 1, 0, 15)
        if month_int == 12:
            end_date = pd.Timestamp(year_int + 1, 1, 1, 0, 0)
        else:
            end_date = pd.Timestamp(year_int, month_int + 1, 1, 0, 0)
        def extract_start_time(t):
            t = str(t).strip()
            if '-' in t:
                return t.split('-')[0].strip()
            return t
        cons_df['Slot_Start'] = cons_df['Time'].apply(extract_start_time)
        cons_df['DateTime'] = pd.to_datetime(cons_df['Date'].dt.strftime('%Y-%m-%d') + ' ' + cons_df['Slot_Start'], errors='coerce')
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
            return None, 'Invalid date format for filter: {}. Use dd/mm/yyyy.'.format(date_filter)
    cons_df = filtered_cons
    cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce') * multiplication_factor
    cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1)
    cons_df['Slot_Time'] = cons_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
    cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')

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
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month:
            pdf.cell(0, 10, f'Month: {month}', ln=True)
        if year:
            pdf.cell(0, 10, f'Year: {year}', ln=True)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Step-by-step Calculation:', ln=True)
        pdf.set_font('Arial', '', 11)
        pdf.multi_cell(0, 8, "1. For each 15-minute slot, generated energy (MW) is converted to kWh (MW * 250).\n2. T&D loss is deducted: After_Loss = Generated_kWh * (1 - T&D loss / 100).\n3. Consumed energy is multiplied by the entered multiplication factor.\n4. For each slot, Excess = Generated_After_Loss - Consumed_Energy.\n5. The table below shows the slot-wise calculation and excess.")
        pdf.ln(2)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(30, 8, 'Date', 1)
        pdf.cell(25, 8, 'Time', 1)
        pdf.cell(35, 8, 'Gen. After Loss', 1)
        pdf.cell(35, 8, 'Consumed', 1)
        pdf.cell(35, 8, 'Excess', 1)
        pdf.cell(30, 8, 'Missing Info', 1)
        pdf.ln()
        pdf.set_font('Arial', '', 10)
        for idx, row in pdf_data.iterrows():
            pdf.cell(30, 8, str(row['Slot_Date']), 1)
            pdf.cell(25, 8, str(row['Slot_Time']), 1)
            pdf.cell(35, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(35, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(35, 8, f"{row['Excess']:.4f}", 1)
            pdf.cell(30, 8, row.get('Missing_Info', ''), 1)
            pdf.ln()
        pdf.ln(2)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {pdf_data["Energy_kWh_gen"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {pdf_data["After_Loss"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {pdf_data["Energy_kWh_cons"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Excess Energy (kWh): {pdf_data["Excess"].sum():.4f}', ln=True)
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
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f'Consumer Number: {consumer_number}', ln=True)
        pdf.cell(0, 10, f'Consumer Name: {consumer_name}', ln=True)
        pdf.cell(0, 10, f'T&D Loss (%): {t_and_d_loss}', ln=True)
        pdf.cell(0, 10, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
        if month:
            pdf.cell(0, 10, f'Month: {month}', ln=True)
        if year:
            pdf.cell(0, 10, f'Year: {year}', ln=True)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Day-wise Summary (All Days in Month):', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 8, 'Date', 1)
        pdf.cell(50, 8, 'Total Gen. After Loss', 1)
        pdf.cell(50, 8, 'Total Consumed', 1)
        pdf.cell(50, 8, 'Total Excess', 1)
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
        daywise = pdf_data.groupby('Slot_Date').agg({
            'After_Loss': 'sum',
            'Energy_kWh_cons': 'sum',
            'Excess': 'sum'
        })
        daywise = daywise.reindex(all_days.strftime('%d/%m/%Y'), fill_value=0).reset_index()
        daywise = daywise.rename(columns={'index': 'Slot_Date'})
        for idx, row in daywise.iterrows():
            pdf.cell(40, 8, row['Slot_Date'], 1)
            pdf.cell(50, 8, f"{row['After_Loss']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
            pdf.cell(50, 8, f"{row['Excess']:.4f}", 1)
            pdf.ln()
        pdf.ln(2)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {pdf_data["Energy_kWh_gen"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {pdf_data["After_Loss"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {pdf_data["Energy_kWh_cons"].sum():.4f}', ln=True)
        pdf.cell(0, 10, f'Total Excess Energy (kWh): {pdf_data["Excess"].sum():.4f}', ln=True)
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
    if not generated_file or not consumed_file:
        st.error('Please upload both Excel files.')
    elif not consumer_number or not consumer_name or not year:
        st.error('Please fill in all required fields.')
    else:
        with st.spinner('Processing...'):
            file_bytes, file_name, mime = process_and_generate(
                generated_file, consumed_file, consumer_number, consumer_name, t_and_d_loss, multiplication_factor, month, year, date_filter, show_excess_only, show_all_slots, show_daywise
            )
            if file_bytes and file_name:
                st.success('Report generated!')
                st.download_button('Download', file_bytes, file_name=file_name, mime=mime or 'application/pdf')
            else:
                st.error('Failed to generate report. Please check your input data.')
