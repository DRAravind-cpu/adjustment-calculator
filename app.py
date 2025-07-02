from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        print('DEBUG request.form:', dict(request.form))
        # Get PDF options from form
        show_excess_only = request.form.get('show_excess_only') == '1'
        show_all_slots = request.form.get('show_all_slots') == '1'
        show_daywise = request.form.get('show_daywise') == '1'

        # Get form data
        t_and_d_loss = float(request.form['t_and_d_loss'])
        consumer_number = request.form['consumer_number']
        consumer_name = request.form['consumer_name']
        multiplication_factor = float(request.form.get('multiplication_factor', 1))
        month = request.form.get('month', '')
        year = request.form.get('year', '')
        date_filter = request.form.get('date', '').strip()  # Optional single date filter (format: dd/mm/yyyy)

        # Get uploaded files
        generated_file = request.files['generated_file']
        consumed_file = request.files['consumed_file']

        # Read and validate generated energy Excel
        gen_df = pd.read_excel(generated_file, header=0)
        if gen_df.shape[1] < 3:
            return render_template('index.html', error="Generated energy Excel must have at least 3 columns: Date, Time, and Energy in MW.")
        gen_df = gen_df.iloc[:, :3]
        gen_df.columns = ['Date', 'Time', 'Energy_MW']
        # Strip whitespace from Date and Time columns
        gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
        gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
        # Standardize date format to yyyy-mm-dd for robust filtering
        gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
        # Filter by year/month with custom slot logic (handle slot ranges in Time column)
        filtered_gen = gen_df.copy()
        if year and month:
            year_int = int(year)
            month_int = int(month)
            # Start: 1st of month at 00:15, End: 1st of next month at 00:00 (inclusive)
            start_date = pd.Timestamp(year_int, month_int, 1, 0, 15)
            if month_int == 12:
                end_date = pd.Timestamp(year_int + 1, 1, 1, 0, 0)
            else:
                end_date = pd.Timestamp(year_int, month_int + 1, 1, 0, 0)
            # Extract slot start time from range if needed
            def extract_start_time(t):
                t = str(t).strip()
                if '-' in t:
                    return t.split('-')[0].strip()
                return t
            gen_df['Slot_Start'] = gen_df['Time'].apply(extract_start_time)
            # Combine Date and Slot_Start for filtering
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
                return render_template('index.html', error=f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.")
        if (year or month or date_filter) and filtered_gen.empty:
            # Debug output for root cause
            debug_msg = []
            debug_msg.append(f"[DEBUG] Filtered gen_df is empty after applying filters.")
            debug_msg.append(f"[DEBUG] year: {year}, month: {month}, date_filter: {date_filter}")
            debug_msg.append(f"[DEBUG] gen_df['Date'] sample: {gen_df['Date'].head(10).tolist()}")
            debug_msg.append(f"[DEBUG] gen_df['Time'] sample: {gen_df['Time'].head(10).tolist()}")
            if 'DateTime' in gen_df.columns:
                debug_msg.append(f"[DEBUG] gen_df['DateTime'] sample: {gen_df['DateTime'].head(10).tolist()}")
            available_months = ', '.join(sorted(gen_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
            debug_msg.append(f"[DEBUG] Available dates: {available_months}")
            return render_template('index.html', error=f"No data for the selected filter in the GENERATED file. Available dates: {available_months}\n\n" + '\n'.join(debug_msg))
        gen_df = filtered_gen
        gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
        gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
        gen_df['After_Loss'] = gen_df['Energy_kWh'] * (1 - t_and_d_loss / 100)
        # Shift generated time by -15 minutes to represent slot start
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

        # Read and validate consumed energy Excel
        cons_df = pd.read_excel(consumed_file, header=0)
        if cons_df.shape[1] < 3:
            return render_template('index.html', error="Consumed energy Excel must have at least 3 columns: Date, Time, and Energy in kWh.")
        cons_df = cons_df.iloc[:, :3]
        cons_df.columns = ['Date', 'Time', 'Energy_kWh']
        # Strip whitespace from Date and Time columns
        cons_df['Date'] = cons_df['Date'].astype(str).str.strip()
        cons_df['Time'] = cons_df['Time'].astype(str).str.strip()
        # Standardize date format to yyyy-mm-dd for robust filtering
        cons_df['Date'] = pd.to_datetime(cons_df['Date'], errors='coerce', dayfirst=True)
        # Debug: Print first 10 unique slot dates and times for both files (after slot columns are created for both)
        if 'Slot_Date' in gen_df.columns and 'Slot_Time' in gen_df.columns:
            print('GEN Slot_Date:', gen_df['Slot_Date'].unique()[:10])
            print('GEN Slot_Time:', gen_df['Slot_Time'].unique()[:10])
        else:
            print('GEN Slot_Date/Slot_Time columns missing!')
        if 'Slot_Date' in cons_df.columns and 'Slot_Time' in cons_df.columns:
            print('CON Slot_Date:', cons_df['Slot_Date'].unique()[:10])
            print('CON Slot_Time:', cons_df['Slot_Time'].unique()[:10])
        else:
            print('CON Slot_Date/Slot_Time columns missing!')
        # Filter by year/month with custom slot logic (handle slot ranges in Time column)
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
                return render_template('index.html', error=f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy.")
        if (year or month or date_filter) and filtered_cons.empty:
            available_months = ', '.join(sorted(cons_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
            return render_template('index.html', error=f"No data for the selected filter in the CONSUMED file. Available dates: {available_months}")
        cons_df = filtered_cons
        cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce') * multiplication_factor
        # Standardize slot time to 'HH:MM - HH:MM' format for consumption too
        cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1)
        # Fix slot time: change '23:45 - 24:00' to '23:45 - 00:00' (if present)
        cons_df['Slot_Time'] = cons_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
        cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')

        # Debug: Print first 10 unique slot dates and times for both files (after slot columns are created)
        if 'Slot_Date' in gen_df.columns and 'Slot_Time' in gen_df.columns:
            print('GEN Slot_Date:', gen_df['Slot_Date'].unique()[:10])
            print('GEN Slot_Time:', gen_df['Slot_Time'].unique()[:10])
        else:
            print('GEN Slot_Date/Slot_Time columns missing!')
        if 'Slot_Date' in cons_df.columns and 'Slot_Time' in cons_df.columns:
            print('CON Slot_Date:', cons_df['Slot_Date'].unique()[:10])
            print('CON Slot_Time:', cons_df['Slot_Time'].unique()[:10])
        else:
            print('CON Slot_Date/Slot_Time columns missing!')

        # Merge on slot date and time (inner join: only slots present in both)
        # Build a full set of all Slot_Date/Slot_Time pairs from both files
        gen_slots_set = set((d, t) for d, t in zip(gen_df['Slot_Date'], gen_df['Slot_Time']))
        cons_slots_set = set((d, t) for d, t in zip(cons_df['Slot_Date'], cons_df['Slot_Time']))
        all_slots = pd.DataFrame(
            list(gen_slots_set | cons_slots_set),
            columns=['Slot_Date', 'Slot_Time'])
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
        # Track missing slots for reporting (only if slot not present in original data, not if value is zero)
        merged['Missing_Info'] = ''
        merged['is_missing_gen'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in gen_slots_set, axis=1)
        merged['is_missing_cons'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in cons_slots_set, axis=1)
        merged.loc[merged['is_missing_gen'], 'Missing_Info'] += '[Missing in GENERATED] '
        merged.loc[merged['is_missing_cons'], 'Missing_Info'] += '[Missing in CONSUMED] '
        merged.drop(['is_missing_gen', 'is_missing_cons'], axis=1, inplace=True)
        # Compose error/warning message for PDF
        error_message = ''
        local_missing_days_msg = globals().get('missing_days_msg', '')
        local_slot_mismatch_msg = globals().get('slot_mismatch_msg', '')
        if local_missing_days_msg or local_slot_mismatch_msg:
            error_message = local_missing_days_msg + local_slot_mismatch_msg + '\nProceeding with only the matching days and slots (missing slots filled with zero).'

        # Check for missing days in either file (priority: match Slot_Date between files)
        gen_days = set(gen_df['Slot_Date'].unique())
        cons_days = set(cons_df['Slot_Date'].unique())
        common_days = gen_days & cons_days
        missing_in_gen = sorted(list(cons_days - gen_days))
        missing_in_cons = sorted(list(gen_days - cons_days))
        missing_days_msg = ""
        if missing_in_gen:
            missing_days_msg += f"Warning: The following days are present in CONSUMED but missing in GENERATED: {', '.join(missing_in_gen)}\n"
        if missing_in_cons:
            missing_days_msg += f"Warning: The following days are present in GENERATED but missing in CONSUMED: {', '.join(missing_in_cons)}\n"
        # Check for missing slots (time intervals) for common days
        slot_mismatch_msg = ""
        for day in sorted(common_days):
            gen_slots = set(gen_df[gen_df['Slot_Date'] == day]['Slot_Time'])
            cons_slots = set(cons_df[cons_df['Slot_Date'] == day]['Slot_Time'])
            missing_slots_in_gen = cons_slots - gen_slots
            missing_slots_in_cons = gen_slots - cons_slots
            if missing_slots_in_gen:
                slot_mismatch_msg += f"Day {day}: Slots in CONSUMED but missing in GENERATED: {', '.join(sorted(missing_slots_in_gen))}\n"
            if missing_slots_in_cons:
                slot_mismatch_msg += f"Day {day}: Slots in GENERATED but missing in CONSUMED: {', '.join(sorted(missing_slots_in_cons))}\n"

        # If there are missing days or slots, prompt user to proceed
        # Always proceed with only the matching days and slots (intersection)
        # Remove any slots not present in both files (already handled by merge)
        warning_msg = ''
        if missing_days_msg or slot_mismatch_msg:
            warning_msg = missing_days_msg + slot_mismatch_msg
            warning_msg += "\nProceeding with only the matching days and slots (intersection)." 
        # Do not return here; warning_msg will be included in the PDF output

        # Calculate excess: a = After_Loss, b = Energy_kWh (consumed)
        merged['Excess'] = merged['After_Loss'] - merged['Energy_kWh_cons']
        merged['Excess'] = merged['Excess'].apply(lambda x: x if x > 0 else 0)
        # Sort merged data chronologically by Slot_Date and Slot_Time (slot start)
        def slot_time_to_minutes(slot_time):
            # Extract start time from 'HH:MM - HH:MM' and convert to minutes since midnight
            try:
                start = slot_time.split('-')[0].strip()
                h, m = map(int, start.split(':'))
                return h * 60 + m
            except Exception:
                return 0
        merged['Slot_Date_dt'] = pd.to_datetime(merged['Slot_Date'], format='%d/%m/%Y', errors='coerce')
        merged['Slot_Time_min'] = merged['Slot_Time'].apply(slot_time_to_minutes)
        merged = merged.sort_values(['Slot_Date_dt', 'Slot_Time_min']).reset_index(drop=True)
        merged.drop(['Slot_Date_dt', 'Slot_Time_min'], axis=1, inplace=True)
        # Totals
        sum_injection = merged['Energy_kWh_gen'].sum()  # Generated before loss
        total_generated_after_loss = merged['After_Loss'].sum()
        total_consumed = merged['Energy_kWh_cons'].sum()
        total_excess = merged['Excess'].sum()
        comparison = sum_injection - total_generated_after_loss
        # For PDF, show all slots or only excess slots
        merged_excess = merged[merged['Excess'] > 0].copy()
        merged_all = merged.copy()
        sum_injection_excess = merged_excess['Energy_kWh_gen'].sum()
        total_generated_after_loss_excess = merged_excess['After_Loss'].sum()
        total_consumed_excess = merged_excess['Energy_kWh_cons'].sum()
        total_excess_excess = merged_excess['Excess'].sum()
        sum_injection_all = merged_all['Energy_kWh_gen'].sum()
        total_generated_after_loss_all = merged_all['After_Loss'].sum()
        total_consumed_all = merged_all['Energy_kWh_cons'].sum()
        total_excess_all = merged_all['Excess'].sum()
        excess_status = 'Excess' if total_excess > 0 else 'No Excess'
        comparison_excess = sum_injection_excess - total_generated_after_loss_excess
        comparison_all = sum_injection_all - total_generated_after_loss_all

        # Count unique days for generated and consumed (in merged data)
        unique_days_gen = merged['Slot_Date'].nunique()
        unique_days_cons = merged['Slot_Date'].nunique()  # Same as gen in merged, but for full data:
        unique_days_gen_full = gen_df['Slot_Date'].nunique()
        unique_days_cons_full = cons_df['Slot_Date'].nunique()

        def format_time(t):
            # Remove leading zeros and show as HH:MM (or as in input)
            if isinstance(t, str):
                t = t.strip()
                if len(t) >= 5 and t[:2] == '00':
                    t = t[2:]
                if len(t) == 4 and t[1] == ':':
                    t = '0' + t
                return t
            return str(t)

        def safe_date_str(d):
            if pd.isnull(d):
                return ''
            if isinstance(d, str):
                return d
            return d.strftime('%d/%m/%Y')
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
            # 15-min slot-wise table (existing)
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
                pdf.cell(30, 8, safe_date_str(row['Slot_Date']), 1)
                pdf.cell(25, 8, format_time(row['Slot_Time']), 1)
                pdf.cell(35, 8, f"{row['After_Loss']:.4f}", 1)
                pdf.cell(35, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
                pdf.cell(35, 8, f"{row['Excess']:.4f}", 1)
                # Reduce width of Missing Info column for better alignment
                pdf.cell(30, 8, row.get('Missing_Info', ''), 1)
                pdf.ln()
            pdf.ln(2)
            # Show error/warning message in PDF if present
            if 'error_message' in globals() and error_message:
                pdf.set_font('Arial', 'B', 10)
                pdf.multi_cell(0, 8, f"Warnings/Errors:\n{error_message}")

            # No day-wise summary table in slot-wise PDFs
            pdf.ln(2)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {sum_injection:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {total_generated_after_loss:.4f}', ln=True)
            pdf.cell(0, 10, f'Comparison (Injection - After Loss): {comparison:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {total_consumed:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Excess Energy (kWh): {total_excess:.4f}', ln=True)
            pdf.cell(0, 10, f'Unique Days Used (Generated): {unique_days_gen_full}', ln=True)
            pdf.cell(0, 10, f'Unique Days Used (Consumed): {unique_days_cons_full}', ln=True)
            pdf.cell(0, 10, f'Status: {excess_status}', ln=True)
            pdf_output = io.BytesIO()
            pdf_bytes = pdf.output(dest='S')
            if isinstance(pdf_bytes, str):
                pdf_bytes = pdf_bytes.encode('latin1')
            pdf_output.write(pdf_bytes)
            pdf_output.seek(0)
            return pdf_output

        # Generate PDFs as per user selection
        pdfs = []
        def generate_daywise_pdf(pdf_data, month, year, filename):
            # This function generates a PDF with only the day-wise summary table (all days in month, even if missing)
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
            import pandas as pd
            # Determine full date range for the selected month
            if month and year:
                month_int = int(month)
                year_int = int(year)
                start_date = datetime(year_int, month_int, 1)
                if month_int == 12:
                    end_date = datetime(year_int + 1, 1, 1)
                else:
                    end_date = datetime(year_int, month_int + 1, 1)
            else:
                # Fallback: use min/max date in data
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
            # Add calculation summary at the end
            sum_injection = pdf_data['Energy_kWh_gen'].sum() if 'Energy_kWh_gen' in pdf_data.columns else 0
            total_generated_after_loss = pdf_data['After_Loss'].sum() if 'After_Loss' in pdf_data.columns else 0
            comparison = sum_injection - total_generated_after_loss
            total_consumed = pdf_data['Energy_kWh_cons'].sum() if 'Energy_kWh_cons' in pdf_data.columns else 0
            total_excess = pdf_data['Excess'].sum() if 'Excess' in pdf_data.columns else 0
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, f'Sum of Injection (Generated before Loss, kWh): {sum_injection:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Generated Energy after Loss (kWh): {total_generated_after_loss:.4f}', ln=True)
            pdf.cell(0, 10, f'Comparison (Injection - After Loss): {comparison:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Consumed Energy (kWh, after multiplication): {total_consumed:.4f}', ln=True)
            pdf.cell(0, 10, f'Total Excess Energy (kWh): {total_excess:.4f}', ln=True)
            pdf_output = io.BytesIO()
            pdf_bytes = pdf.output(dest='S')
            if isinstance(pdf_bytes, str):
                pdf_bytes = pdf_bytes.encode('latin1')
            pdf_output.write(pdf_bytes)
            pdf_output.seek(0)
            return pdf_output

        # Defensive: If no PDF options are selected, default to generating 'all slots' PDF
        if not (show_excess_only or show_all_slots or show_daywise):
            show_all_slots = True
        import traceback
        if show_excess_only:
            try:
                print('DEBUG: Generating excess only PDF...')
                pdf_obj = generate_pdf(
                    merged_excess, sum_injection_excess, total_generated_after_loss_excess, comparison_excess, total_consumed_excess, total_excess_excess, excess_status, 'energy_adjustment_excess_only.pdf')
                print('DEBUG: generate_pdf (excess only) returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append(('energy_adjustment_excess_only.pdf', pdf_obj))
            except Exception as e:
                print('ERROR in generate_pdf (excess only):', e)
                traceback.print_exc()
        if show_all_slots:
            try:
                print('DEBUG: Generating all slots PDF...')
                pdf_obj = generate_pdf(
                    merged_all, sum_injection_all, total_generated_after_loss_all, comparison_all, total_consumed_all, total_excess_all, excess_status, 'energy_adjustment_all_slots.pdf')
                print('DEBUG: generate_pdf (all slots) returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append(('energy_adjustment_all_slots.pdf', pdf_obj))
            except Exception as e:
                print('ERROR in generate_pdf (all slots):', e)
                traceback.print_exc()
        if show_daywise:
            try:
                print('DEBUG: Generating daywise PDF...')
                pdf_obj = generate_daywise_pdf(
                    merged_all, month, year, 'energy_adjustment_daywise.pdf')
                print('DEBUG: generate_daywise_pdf returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append(('energy_adjustment_daywise.pdf', pdf_obj))
            except Exception as e:
                print('ERROR in generate_daywise_pdf:', e)
                traceback.print_exc()

        # If both PDFs, zip and send, else send single
        if len(pdfs) >= 2:
            print('DEBUG: Returning ZIP file to client')
            import zipfile
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                for fname, pdf_io in pdfs:
                    zf.writestr(fname, pdf_io.getvalue())
            zip_buffer.seek(0)
            return send_file(zip_buffer, as_attachment=True, download_name='energy_adjustment_reports.zip', mimetype='application/zip')
        elif len(pdfs) == 1:
            fname, pdf_io = pdfs[0]
            print(f'DEBUG: Returning PDF file to client: {fname}')
            return send_file(pdf_io, as_attachment=True, download_name=fname, mimetype='application/pdf')
        else:
            print('DEBUG: No PDF generated, returning error page')
            return render_template('index.html', error="Please select at least one PDF output option.")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
