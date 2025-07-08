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
        
        # Debug PDF options
        print(f"DEBUG: Raw PDF options from form: show_excess_only={request.form.get('show_excess_only')}, show_all_slots={request.form.get('show_all_slots')}, show_daywise={request.form.get('show_daywise')}")
        print(f"DEBUG: Processed PDF options: show_excess_only={show_excess_only}, show_all_slots={show_all_slots}, show_daywise={show_daywise}")

        # Get form data
        t_and_d_loss = float(request.form['t_and_d_loss'])
        consumer_number = request.form['consumer_number']
        consumer_name = request.form['consumer_name']
        multiplication_factor = float(request.form.get('multiplication_factor', 1))
        auto_detect_month = request.form.get('auto_detect_month') == '1'
        month = request.form.get('month', '')
        year = request.form.get('year', '')
        date_filter = request.form.get('date', '').strip()  # Optional single date filter (format: dd/mm/yyyy)

        # Get uploaded files
        generated_files = request.files.getlist('generated_files')
        consumed_files = request.files.getlist('consumed_files')
        
        if not generated_files:
            return render_template('index.html', error="No generated energy Excel files were uploaded.")
        if not consumed_files:
            return render_template('index.html', error="No consumed energy Excel files were uploaded.")
            
        # Process multiple generated energy Excel files
        gen_dfs = []
        for gen_file in generated_files:
            try:
                temp_df = pd.read_excel(gen_file, header=0)
                if temp_df.shape[1] < 3:
                    return render_template('index.html', error=f"Generated energy Excel file '{gen_file.filename}' must have at least 3 columns: Date, Time, and Energy in MW.")
                
                # Add filename to help with debugging
                temp_df['Source_File'] = gen_file.filename
                gen_dfs.append(temp_df)
            except Exception as e:
                return render_template('index.html', error=f"Error reading generated energy Excel file '{gen_file.filename}': {str(e)}")
        
        # Combine all generated energy dataframes
        if not gen_dfs:
            return render_template('index.html', error="No valid generated energy Excel files were found.")
        gen_df = pd.concat(gen_dfs, ignore_index=True)
        gen_df = gen_df.iloc[:, :3]
        gen_df.columns = ['Date', 'Time', 'Energy_MW']
        # Strip whitespace from Date and Time columns
        gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
        gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
        # Standardize date format to yyyy-mm-dd for robust filtering
        gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
        
        # Function to convert month number to month name
        def get_month_name(month_num):
            month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                          'July', 'August', 'September', 'October', 'November', 'December']
            try:
                return month_names[int(month_num) - 1]
            except (ValueError, IndexError):
                return str(month_num)
        
        # Auto-detect month and year if enabled
        if auto_detect_month and not (month and year):
            # Extract unique months and years from the data
            unique_months = gen_df['Date'].dt.month.unique()
            unique_years = gen_df['Date'].dt.year.unique()
            
            if len(unique_months) == 1 and not month:
                month = str(unique_months[0])
                print(f"Auto-detected month: {month} ({get_month_name(month)})")
            elif len(unique_months) > 1 and not month:
                # If multiple months, use the most frequent one
                month = str(gen_df['Date'].dt.month.value_counts().idxmax())
                print(f"Multiple months detected, using most frequent: {month} ({get_month_name(month)})")
            
            if len(unique_years) == 1 and not year:
                year = str(unique_years[0])
                print(f"Auto-detected year: {year}")
            elif len(unique_years) > 1 and not year:
                # If multiple years, use the most frequent one
                year = str(gen_df['Date'].dt.year.value_counts().idxmax())
                print(f"Multiple years detected, using most frequent: {year}")
                
            # Add information to be displayed in PDF
            auto_detect_info = f"Auto-detected from {len(generated_files)} generated and {len(consumed_files)} consumed files"
        
        # Filter by year/month with custom slot logic (handle slot ranges in Time column)
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
            # Ensure each day has exactly 96 slots (15-minute intervals from 00:00 to 23:45)
            print(f"Filtered generated data: {len(filtered_gen)} rows")
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

        # Process multiple consumed energy Excel files
        cons_dfs = []
        for cons_file in consumed_files:
            try:
                temp_df = pd.read_excel(cons_file, header=0)
                if temp_df.shape[1] < 3:
                    return render_template('index.html', error=f"Consumed energy Excel file '{cons_file.filename}' must have at least 3 columns: Date, Time, and Energy in kWh.")
                
                # Add filename to help with debugging
                temp_df['Source_File'] = cons_file.filename
                cons_dfs.append(temp_df)
            except Exception as e:
                return render_template('index.html', error=f"Error reading consumed energy Excel file '{cons_file.filename}': {str(e)}")
        
        # Combine all consumed energy dataframes
        if not cons_dfs:
            return render_template('index.html', error="No valid consumed energy Excel files were found.")
        cons_df = pd.concat(cons_dfs, ignore_index=True)
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
            # Ensure each day has exactly 96 slots (15-minute intervals from 00:00 to 23:45)
            print(f"Filtered consumed data: {len(filtered_cons)} rows")
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

        # Validate that each day has exactly 96 slots (15-minute intervals from 00:00 to 23:45)
        def validate_day_slots(df, date_col='Slot_Date', time_col='Slot_Time'):
            """Check if each day has exactly 96 slots and identify missing slots."""
            # Expected slots for a complete day (00:00 to 23:45 in 15-minute intervals)
            expected_slots = []
            for hour in range(24):
                for minute in [0, 15, 30, 45]:
                    start_time = f"{hour:02d}:{minute:02d}"
                    end_minute = (minute + 15) % 60
                    end_hour = hour + 1 if end_minute == 0 else hour
                    if end_hour == 24:  # Handle the case of 23:45-00:00
                        end_hour = 0
                    end_time = f"{end_hour:02d}:{end_minute:02d}"
                    expected_slots.append(f"{start_time} - {end_time}")
            
            # Check each day in the dataframe
            day_slot_counts = df.groupby(date_col)[time_col].count()
            incomplete_days = day_slot_counts[day_slot_counts != 96].index.tolist()
            
            if incomplete_days:
                print(f"Warning: The following days do not have exactly 96 slots: {incomplete_days}")
                # For each incomplete day, identify which slots are missing
                for day in incomplete_days:
                    day_slots = set(df[df[date_col] == day][time_col])
                    missing_slots = [slot for slot in expected_slots if slot not in day_slots]
                    print(f"Day {day} is missing {len(missing_slots)} slots: {missing_slots[:5]}...")
            
            return incomplete_days, expected_slots
        
        # Check for incomplete days in both datasets
        gen_incomplete_days, expected_slots = validate_day_slots(gen_df)
        cons_incomplete_days, _ = validate_day_slots(cons_df)
        
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
        
        merged['TOD_Category'] = merged['Slot_Time'].apply(classify_tod)
        
        # Group excess energy by TOD category
        tod_excess = merged.groupby('TOD_Category')['Excess'].sum().reset_index()
        
        # Calculate financial values
        total_excess = merged['Excess'].sum()
        
        # Base rate for all excess energy
        base_rate = 7.25  # rupees per kWh
        base_amount = total_excess * base_rate
        
        # Additional charges for specific TOD categories
        c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
        c1_c2_additional = c1_c2_excess * 1.81  # rupees per kWh
        
        c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
        c5_additional = c5_excess * 0.3625  # rupees per kWh
        
        # Calculate total amount
        total_amount = base_amount + c1_c2_additional + c5_additional
        
        # Calculate E-Tax (5% of total amount)
        etax = total_amount * 0.05
        
        # Calculate total amount with E-Tax
        total_with_etax = total_amount + etax
        
        # Calculate negative factors
        etax_on_iex = total_excess * 0.1
        cross_subsidy_surcharge = total_excess * 1.92
        wheeling_charges = cross_subsidy_surcharge * 1.04
        
        # Calculate final amount to be collected
        final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
        
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
        # Use the total consumed energy from all slots for consistency across all PDFs
        total_consumed_excess = merged['Energy_kWh_cons'].sum()  # Changed from merged_excess to merged
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
        def generate_pdf(pdf_data, sum_injection, total_generated_after_loss, comparison, total_consumed, total_excess, excess_status, filename, auto_detect=auto_detect_month, gen_files=generated_files, cons_files=consumed_files):
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
            # Define TOD descriptions for reference (time ranges removed for better readability)
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
            pdf.cell(60, 8, 'TOD Category & Description', 1)
            pdf.cell(25, 8, 'Gen. After Loss', 1)
            pdf.cell(25, 8, 'Consumed', 1)
            pdf.cell(25, 8, 'Excess', 1)
            pdf.cell(10, 8, 'Miss', 1)
            pdf.ln()
            pdf.set_font('Arial', '', 9)  # Slightly smaller font for better fit
            for idx, row in pdf_data.iterrows():
                pdf.cell(20, 8, safe_date_str(row['Slot_Date']), 1)
                pdf.cell(25, 8, format_time(row['Slot_Time']), 1)
                
                # Combine TOD category and description
                tod_cat = row.get('TOD_Category', '')
                tod_desc = tod_descriptions.get(tod_cat, '')
                combined_tod = f"{tod_cat}: {tod_desc}" if tod_cat else ""
                pdf.cell(60, 8, combined_tod, 1)
                
                pdf.cell(25, 8, f"{row['After_Loss']:.4f}", 1)
                pdf.cell(25, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
                pdf.cell(25, 8, f"{row['Excess']:.4f}", 1)
                pdf.cell(10, 8, row.get('Missing_Info', ''), 1)
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
            
            # Add financial calculations
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Financial Calculations:', ln=True)
            pdf.set_font('Arial', '', 10)
            
            # Base rate calculation
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess * base_rate
            pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess:.4f} kWh) x Rs.7.25 = Rs.{base_amount:.2f}", ln=True)
            
            # Additional charges for specific TOD categories
            c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
            c1_c2_additional = c1_c2_excess * 1.81  # rupees per kWh
            pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess:.4f} kWh) x Rs.1.81 = Rs.{c1_c2_additional:.2f}", ln=True)
            
            c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
            c5_additional = c5_excess * 0.3625  # rupees per kWh
            pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess:.4f} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
            
            # Calculate total amount
            total_amount = base_amount + c1_c2_additional + c5_additional
            pdf.cell(0, 8, f"4. Total Amount: Rs.{base_amount:.2f} + Rs.{c1_c2_additional:.2f} + Rs.{c5_additional:.2f} = Rs.{total_amount:.2f}", ln=True)
            
            # Calculate E-Tax (5% of total amount)
            etax = total_amount * 0.05
            pdf.cell(0, 8, f"5. E-Tax (5% of Total Amount): Rs.{total_amount:.2f} x 0.05 = Rs.{etax:.2f}", ln=True)
            
            # Calculate total amount with E-Tax
            total_with_etax = total_amount + etax
            pdf.cell(0, 8, f"6. Total Amount with E-Tax: Rs.{total_amount:.2f} + Rs.{etax:.2f} = Rs.{total_with_etax:.2f}", ln=True)
            
            # Calculate negative factors
            etax_on_iex = total_excess * 0.1
            pdf.cell(0, 8, f"7. E-Tax on IEX: Total Excess ({total_excess:.4f} kWh) x Rs.0.1 = Rs.{etax_on_iex:.2f}", ln=True)
            
            cross_subsidy_surcharge = total_excess * 1.92
            pdf.cell(0, 8, f"8. Cross Subsidy Surcharge: Total Excess ({total_excess:.4f} kWh) x Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
            # New formula for wheeling charges
            # Step 1: Calculate x = (Total excess energy * loss % / (100 - loss %))
            loss_percentage = t_and_d_loss  # Using the T&D loss percentage from form input
            x_value = (total_excess * loss_percentage) / (100 - loss_percentage)
            
            # Step 2: Calculate y = (total excess energy * loss % / x) + total excess energy
            y_value = (total_excess * loss_percentage / x_value) + total_excess if x_value != 0 else total_excess
            
            # Step 3: Calculate wheeling charges = y * 1.04
            wheeling_charges = y_value * 1.04
            
            pdf.cell(0, 8, f"9a. Wheeling Charges Step 1: [{total_excess:.4f} x {loss_percentage:.2f}% / (100-{loss_percentage:.2f}%)] = {x_value:.4f}", ln=True)
            pdf.cell(0, 8, f"9b. Wheeling Charges Step 2: [({total_excess:.4f} x {loss_percentage:.2f}% / {x_value:.4f}) + {total_excess:.4f}] x 1.04 = Rs.{wheeling_charges:.2f}", ln=True)
            
            # Calculate final amount to be collected
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 10, f"10. Total Amount to be Collected: Rs.{total_with_etax:.2f} - (Rs.{etax_on_iex:.2f} + Rs.{cross_subsidy_surcharge:.2f} + Rs.{wheeling_charges:.2f}) = Rs.{final_amount:.2f}", ln=True)
            
            try:
                print("DEBUG: Generating PDF output in generate_pdf function...")
                pdf_output = io.BytesIO()
                pdf_bytes = pdf.output(dest='S')
                if isinstance(pdf_bytes, str):
                    print("DEBUG: PDF bytes is a string, encoding to latin1")
                    pdf_bytes = pdf_bytes.encode('latin1')
                pdf_output.write(pdf_bytes)
                pdf_output.seek(0)
                print("DEBUG: PDF generation successful")
                return pdf_output
            except UnicodeEncodeError as e:
                print(f"ERROR: Unicode encoding error in generate_pdf: {e}")
                # Find the problematic character
                if isinstance(e, UnicodeEncodeError):
                    bad_char = e.object[e.start:e.end]
                    print(f"ERROR: Problematic character: '{bad_char}' (Unicode: U+{ord(bad_char[0]):04X})")
                    print(f"ERROR: Position in string: {e.start}")
                    # Get some context around the error
                    context_start = max(0, e.start - 20)
                    context_end = min(len(e.object), e.end + 20)
                    context = e.object[context_start:context_end]
                    print(f"ERROR: Context: '...{context}...'")
                raise

        # Generate PDFs as per user selection
        pdfs = []
        
        # Add auto-detection info to PDF if enabled
        auto_detect_info = ""
        if auto_detect_month:
            auto_detect_info = f"Auto-detected from {len(generated_files)} generated and {len(consumed_files)} consumed files"
            
        # Helper function to add month and year info to PDF
        def add_month_year_info(pdf, month, year):
            if month:
                month_display = f"{get_month_name(month)} ({month})" if month.isdigit() else month
                pdf.cell(0, 10, f'Month: {month_display}', ln=True)
            if year:
                pdf.cell(0, 10, f'Year: {year}', ln=True)
            if auto_detect_info:
                pdf.set_font('Arial', 'I', 10)
                pdf.cell(0, 10, auto_detect_info, ln=True)
                pdf.set_font('Arial', '', 12)
        def generate_daywise_pdf(pdf_data, month, year, filename, auto_detect_info=auto_detect_info):
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
                    end_date = datetime(year_int, 12, 31) + timedelta(days=1)
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
                    end_date = datetime(year_int, month_int, last_day) + timedelta(days=1)
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
            
            # Add financial calculations
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Financial Calculations:', ln=True)
            pdf.set_font('Arial', '', 10)
            
            # Base rate calculation
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess * base_rate
            pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess:.4f} kWh) x Rs.7.25 = Rs.{base_amount:.2f}", ln=True)
            
            # Additional charges for specific TOD categories
            c1_c2_excess = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Excess'].sum()
            c1_c2_additional = c1_c2_excess * 1.81  # rupees per kWh
            pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess:.4f} kWh) x Rs.1.81 = Rs.{c1_c2_additional:.2f}", ln=True)
            
            c5_excess = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Excess'].sum()
            c5_additional = c5_excess * 0.3625  # rupees per kWh
            pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess:.4f} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
            
            # Calculate total amount
            total_amount = base_amount + c1_c2_additional + c5_additional
            pdf.cell(0, 8, f"4. Total Amount: Rs.{base_amount:.2f} + Rs.{c1_c2_additional:.2f} + Rs.{c5_additional:.2f} = Rs.{total_amount:.2f}", ln=True)
            
            # Calculate E-Tax (5% of total amount)
            etax = total_amount * 0.05
            pdf.cell(0, 8, f"5. E-Tax (5% of Total Amount): Rs.{total_amount:.2f} x 0.05 = Rs.{etax:.2f}", ln=True)
            
            # Calculate total amount with E-Tax
            total_with_etax = total_amount + etax
            pdf.cell(0, 8, f"6. Total Amount with E-Tax: Rs.{total_amount:.2f} + Rs.{etax:.2f} = Rs.{total_with_etax:.2f}", ln=True)
            
            # Calculate negative factors
            etax_on_iex = total_excess * 0.1
            pdf.cell(0, 8, f"7. E-Tax on IEX: Total Excess ({total_excess:.4f} kWh) x Rs.0.1 = Rs.{etax_on_iex:.2f}", ln=True)
            
            cross_subsidy_surcharge = total_excess * 1.92
            pdf.cell(0, 8, f"8. Cross Subsidy Surcharge: Total Excess ({total_excess:.4f} kWh) x Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
            # New formula for wheeling charges
            # Step 1: Calculate x = (Total excess energy * loss % / (100 - loss %))
            loss_percentage = t_and_d_loss  # Using the T&D loss percentage from form input
            x_value = (total_excess * loss_percentage) / (100 - loss_percentage)
            
            # Step 2: Calculate y = (total excess energy * loss % / x) + total excess energy
            y_value = (total_excess * loss_percentage / x_value) + total_excess if x_value != 0 else total_excess
            
            # Step 3: Calculate wheeling charges = y * 1.04
            wheeling_charges = y_value * 1.04
            
            pdf.cell(0, 8, f"9a. Wheeling Charges Step 1: [{total_excess:.4f} x {loss_percentage:.2f}% / (100-{loss_percentage:.2f}%)] = {x_value:.4f}", ln=True)
            pdf.cell(0, 8, f"9b. Wheeling Charges Step 2: [({total_excess:.4f} x {loss_percentage:.2f}% / {x_value:.4f}) + {total_excess:.4f}] x 1.04 = Rs.{wheeling_charges:.2f}", ln=True)
            
            # Calculate final amount to be collected
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 10, f"10. Total Amount to be Collected: Rs.{total_with_etax:.2f} - (Rs.{etax_on_iex:.2f} + Rs.{cross_subsidy_surcharge:.2f} + Rs.{wheeling_charges:.2f}) = Rs.{final_amount:.2f}", ln=True)
            
            try:
                print("DEBUG: Generating PDF output in generate_daywise_pdf function...")
                pdf_output = io.BytesIO()
                pdf_bytes = pdf.output(dest='S')
                if isinstance(pdf_bytes, str):
                    print("DEBUG: PDF bytes is a string, encoding to latin1")
                    pdf_bytes = pdf_bytes.encode('latin1')
                pdf_output.write(pdf_bytes)
                pdf_output.seek(0)
                print("DEBUG: PDF generation successful")
                return pdf_output
            except UnicodeEncodeError as e:
                print(f"ERROR: Unicode encoding error in generate_daywise_pdf: {e}")
                # Find the problematic character
                if isinstance(e, UnicodeEncodeError):
                    bad_char = e.object[e.start:e.end]
                    print(f"ERROR: Problematic character: '{bad_char}' (Unicode: U+{ord(bad_char[0]):04X})")
                    print(f"ERROR: Position in string: {e.start}")
                    # Get some context around the error
                    context_start = max(0, e.start - 20)
                    context_end = min(len(e.object), e.end + 20)
                    context = e.object[context_start:context_end]
                    print(f"ERROR: Context: '...{context}...'")
                raise

        # Defensive: If no PDF options are selected, default to generating 'all slots' PDF
        if not (show_excess_only or show_all_slots or show_daywise):
            show_all_slots = True
            print("DEBUG: No PDF options selected, defaulting to show_all_slots=True")
        else:
            print(f"DEBUG: PDF options selected - show_excess_only: {show_excess_only}, show_all_slots: {show_all_slots}, show_daywise: {show_daywise}")
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
        try:
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
        except Exception as e:
            import traceback
            error_traceback = traceback.format_exc()
            print(f"ERROR in PDF generation/delivery: {e}")
            print(error_traceback)
            return render_template('index.html', error=f"Error generating PDF: {str(e)}\n\nTraceback: {error_traceback}")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, port=5002)
