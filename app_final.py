from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': '',
        'C2': '',
        'C4': '',
        'C5': '',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    
    # No TOD descriptions as requested
    pdf.ln(2)
    
    # Then add the table with just category and excess energy
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(50, 8, 'TOD Category', 1)
    pdf.cell(50, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        excess = row['Excess']
        pdf.cell(50, 8, category, 1)
        pdf.cell(50, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excess

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
        
        # Initialize auto_detect_info
        auto_detect_info = None
        
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
            
            # Check if the selected year and month exist in the data
            available_years = gen_df['Date'].dt.year.unique()
            available_months = gen_df['Date'].dt.month.unique()
            
            if year_int not in available_years or month_int not in available_months:
                print(f"WARNING: Selected year ({year_int}) and month ({month_int}) not found in data.")
                print(f"Available years: {available_years}, Available months: {available_months}")
                
                # If auto-detect is enabled, use the available data instead
                if auto_detect_month:
                    if year_int not in available_years and len(available_years) > 0:
                        year_int = available_years[0]
                        year = str(year_int)
                        print(f"Auto-corrected to available year: {year_int}")
                    
                    if month_int not in available_months and len(available_months) > 0:
                        month_int = available_months[0]
                        month = str(month_int)
                        print(f"Auto-corrected to available month: {month_int} ({get_month_name(month)})")
            
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
            # Get available years and months for user-friendly error message
            available_years = sorted(gen_df['Date'].dt.year.unique())
            available_months = sorted(gen_df['Date'].dt.month.unique())
            available_dates = ', '.join(sorted(gen_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
            
            # Create user-friendly error message
            error_msg = f"No data found for the selected filters."
            
            if year and int(year) not in available_years:
                error_msg += f"\n\nThe selected year ({year}) is not available in the data."
                error_msg += f"\nAvailable years: {', '.join(map(str, available_years))}"
            
            if month and int(month) not in available_months:
                month_name = get_month_name(month)
                error_msg += f"\n\nThe selected month ({month_name}) is not available in the data."
                error_msg += f"\nAvailable months: {', '.join([get_month_name(str(m)) for m in available_months])}"
            
            if date_filter and pd.to_datetime(date_filter, dayfirst=True).strftime('%d/%m/%Y') not in available_dates:
                error_msg += f"\n\nThe selected date ({date_filter}) is not available in the data."
            
            error_msg += f"\n\nAvailable dates: {available_dates}"
            
            # Add debug information if needed
            if 'DEBUG' in os.environ.get('FLASK_ENV', ''):
                debug_msg = []
                debug_msg.append(f"[DEBUG] Filtered gen_df is empty after applying filters.")
                debug_msg.append(f"[DEBUG] year: {year}, month: {month}, date_filter: {date_filter}")
                debug_msg.append(f"[DEBUG] gen_df['Date'] sample: {gen_df['Date'].head(10).tolist()}")
                debug_msg.append(f"[DEBUG] gen_df['Time'] sample: {gen_df['Time'].head(10).tolist()}")
                if 'DateTime' in gen_df.columns:
                    debug_msg.append(f"[DEBUG] gen_df['DateTime'] sample: {gen_df['DateTime'].head(10).tolist()}")
                debug_msg.append(f"[DEBUG] Available dates: {available_dates}")
                error_msg += "\n\n" + '\n'.join(debug_msg)
            
            return render_template('index.html', error=error_msg)
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
            # Make sure to initialize year_int and month_int for consumed data
            # Use the values that were potentially auto-corrected for gen_df
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
            # Get available years and months for user-friendly error message
            available_years = sorted(cons_df['Date'].dt.year.unique())
            available_months = sorted(cons_df['Date'].dt.month.unique())
            available_dates = ', '.join(sorted(cons_df['Date'].dt.strftime('%d/%m/%Y').dropna().unique()))
            
            # Create user-friendly error message
            error_msg = f"No data found for the selected filters in the CONSUMED file."
            
            if year and int(year) not in available_years:
                error_msg += f"\n\nThe selected year ({year}) is not available in the consumed data."
                error_msg += f"\nAvailable years: {', '.join(map(str, available_years))}"
            
            if month and int(month) not in available_months:
                month_name = get_month_name(month)
                error_msg += f"\n\nThe selected month ({month_name}) is not available in the consumed data."
                error_msg += f"\nAvailable months: {', '.join([get_month_name(str(m)) for m in available_months])}"
            
            if date_filter and pd.to_datetime(date_filter, dayfirst=True).strftime('%d/%m/%Y') not in available_dates:
                error_msg += f"\n\nThe selected date ({date_filter}) is not available in the consumed data."
            
            error_msg += f"\n\nAvailable dates in consumed data: {available_dates}"
            
            # Add debug information if needed
            if 'DEBUG' in os.environ.get('FLASK_ENV', ''):
                debug_msg = []
                debug_msg.append(f"[DEBUG] Filtered cons_df is empty after applying filters.")
                debug_msg.append(f"[DEBUG] year: {year}, month: {month}, date_filter: {date_filter}")
                debug_msg.append(f"[DEBUG] cons_df['Date'] sample: {cons_df['Date'].head(10).tolist()}")
                debug_msg.append(f"[DEBUG] cons_df['Time'] sample: {cons_df['Time'].head(10).tolist()}")
                if 'DateTime' in cons_df.columns:
                    debug_msg.append(f"[DEBUG] cons_df['DateTime'] sample: {cons_df['DateTime'].head(10).tolist()}")
                debug_msg.append(f"[DEBUG] Available dates: {available_dates}")
                error_msg += "\n\n" + '\n'.join(debug_msg)
            
            return render_template('index.html', error=error_msg)
        cons_df = filtered_cons
        cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce')
        # Apply multiplication factor to consumed energy
        cons_df['Energy_kWh'] = cons_df['Energy_kWh'] * multiplication_factor
        # Standardize slot time to 'HH:MM - HH:MM' format
        cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1)
        # Fix slot time: change '23:45 - 24:00' to '23:45 - 00:00'
        cons_df['Slot_Time'] = cons_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
        cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')
        
        # Merge generated and consumed energy data
        gen_df_for_merge = gen_df[['Slot_Date', 'Slot_Time', 'After_Loss']]
        cons_df_for_merge = cons_df[['Slot_Date', 'Slot_Time', 'Energy_kWh']]
        cons_df_for_merge = cons_df_for_merge.rename(columns={'Energy_kWh': 'Energy_kWh_cons'})
        
        # Merge on Slot_Date and Slot_Time
        merged = pd.merge(gen_df_for_merge, cons_df_for_merge, on=['Slot_Date', 'Slot_Time'], how='outer')
        
        # Fill NaN values with 0
        merged['After_Loss'] = merged['After_Loss'].fillna(0)
        merged['Energy_kWh_cons'] = merged['Energy_kWh_cons'].fillna(0)
        
        # Calculate excess energy
        merged['Excess'] = merged['After_Loss'] - merged['Energy_kWh_cons']
        
        # Add a column to indicate missing data
        merged['Missing_Info'] = ''
        merged.loc[merged['After_Loss'] == 0, 'Missing_Info'] = 'G'  # Missing generated data
        merged.loc[merged['Energy_kWh_cons'] == 0, 'Missing_Info'] = 'C'  # Missing consumed data
        merged.loc[(merged['After_Loss'] == 0) & (merged['Energy_kWh_cons'] == 0), 'Missing_Info'] = 'B'  # Both missing
        
        # Filter to show only excess energy slots if requested
        if show_excess_only:
            merged = merged[merged['Excess'] > 0]
        
        # Filter to show only slots with data in both files
        if not show_all_slots:
            merged = merged[(merged['After_Loss'] > 0) & (merged['Energy_kWh_cons'] > 0)]
        
        # Sort by date and time
        merged = merged.sort_values(['Slot_Date', 'Slot_Time'])
        
        # Calculate total generated and consumed energy
        total_generated_after_loss = merged['After_Loss'].sum()
        total_consumed = merged['Energy_kWh_cons'].sum()
        total_excess = merged['Excess'].sum()
        
        # Calculate sum of injection (before loss)
        sum_injection = gen_df['Energy_kWh'].sum()
        comparison = sum_injection - total_generated_after_loss
        
        # Get unique days in each dataset
        unique_days_gen = gen_df['Slot_Date'].nunique()
        unique_days_cons = cons_df['Slot_Date'].nunique()
        unique_days_gen_full = f"{unique_days_gen} ({', '.join(sorted(gen_df['Slot_Date'].unique()))})"
        unique_days_cons_full = f"{unique_days_cons} ({', '.join(sorted(cons_df['Slot_Date'].unique()))})"
        
        # Determine if there's excess energy
        excess_status = "Excess Energy Available" if total_excess > 0 else "No Excess Energy"
        
        # Function to classify TOD category based on time slot
        def classify_tod(time_slot):
            try:
                # Extract start time from the slot
                if '-' in time_slot:
                    start_time = time_slot.split('-')[0].strip()
                else:
                    start_time = time_slot.strip()
                
                # Convert to datetime.time object
                start_dt = pd.to_datetime(start_time, format='%H:%M').time()
                hour = start_dt.hour
                
                # Classify based on TOD categories
                if 6 <= hour < 10:
                    return 'C1'
                elif 18 <= hour < 22:
                    return 'C2'
                elif 5 <= hour < 6 or 10 <= hour < 18:
                    return 'C4'
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
        
        # Helper function to format time slots for better display
        def format_time(time_str):
            if pd.isna(time_str):
                return ""
            return str(time_str).strip()
        
        # Helper function to safely format dates
        def safe_date_str(date_str):
            if pd.isna(date_str):
                return ""
            return str(date_str).strip()
        
        # Generate PDF report
        def generate_pdf(pdf_data, month, year, filename, auto_detect_info=None):
            # Initialize error_message variable
            error_message = None
            
            class PDF(FPDF):
                def header(self):
                    # Logo
                    # self.image('logo.png', 10, 8, 33)
                    # Arial bold 15
                    self.set_font('Arial', 'B', 15)
                    # Move to the right
                    self.cell(80)
                    # Title
                    self.cell(30, 10, 'Energy Adjustment Report', 0, 0, 'C')
                    # Line break
                    self.ln(20)
                
                def footer(self):
                    # Position at 1.5 cm from bottom
                    self.set_y(-15)
                    # Arial italic 8
                    self.set_font('Arial', 'I', 8)
                    # Page number
                    self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')
            
            # Create PDF object
            pdf = PDF()
            pdf.alias_nb_pages()
            pdf.add_page()
            pdf.set_font('Arial', 'B', 12)
            
            # Add title and consumer information
            pdf.cell(0, 10, f'Energy Adjustment Report - {get_month_name(month)} {year}', ln=True, align='C')
            pdf.set_font('Arial', '', 10)
            if auto_detect_info:
                pdf.cell(0, 8, f"Note: {auto_detect_info}", ln=True)
            pdf.cell(0, 8, f"Consumer Number: {consumer_number}", ln=True)
            pdf.cell(0, 8, f"Consumer Name: {consumer_name}", ln=True)
            pdf.cell(0, 8, f"T&D Loss: {t_and_d_loss}%", ln=True)
            pdf.cell(0, 8, f"Multiplication Factor: {multiplication_factor}", ln=True)
            
            # Add summary information
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Summary:', ln=True)
            pdf.set_font('Arial', '', 11)
            pdf.cell(0, 8, f"Total Generated Energy (after {t_and_d_loss}% T&D loss): {total_generated_after_loss:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Total Consumed Energy (after multiplication factor): {total_consumed:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Total Excess Energy: {total_excess:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Status: {excess_status}", ln=True)
            
            # Add calculation explanation
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Step-by-step Calculation:', ln=True)
            pdf.set_font('Arial', '', 11)
            pdf.multi_cell(0, 8, "1. For each 15-minute slot, generated energy (MW) is converted to kWh (MW * 250).\n2. T&D loss is deducted: After_Loss = Generated_kWh * (1 - T&D loss / 100).\n3. Consumed energy is multiplied by the entered multiplication factor.\n4. For each slot, Excess = Generated_After_Loss - Consumed_Energy.\n5. The table below shows the slot-wise calculation and excess.")
            pdf.ln(2)
            # Define TOD descriptions for reference
            tod_descriptions = {
                'C1': '',
                'C2': '',
                'C4': '',
                'C5': '',
                'Unknown': 'Unknown Time Slot'
            }
            
            # No TOD descriptions as requested
            pdf.ln(2)
            
            # 15-min slot-wise table with improved layout
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(20, 8, 'Date', 1)
            pdf.cell(25, 8, 'Time', 1)
            pdf.cell(25, 8, 'TOD Category', 1)
            pdf.cell(30, 8, 'Gen. After Loss', 1)
            pdf.cell(30, 8, 'Consumed', 1)
            pdf.cell(30, 8, 'Excess', 1)
            pdf.cell(10, 8, 'Miss', 1)
            pdf.ln()
            pdf.set_font('Arial', '', 9)  # Slightly smaller font for better fit
            for idx, row in pdf_data.iterrows():
                pdf.cell(20, 8, safe_date_str(row['Slot_Date']), 1)
                pdf.cell(25, 8, format_time(row['Slot_Time']), 1)
                pdf.cell(25, 8, row.get('TOD_Category', ''), 1)
                pdf.cell(30, 8, f"{row['After_Loss']:.4f}", 1)
                pdf.cell(30, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
                pdf.cell(30, 8, f"{row['Excess']:.4f}", 1)
                pdf.cell(10, 8, row.get('Missing_Info', ''), 1)
                pdf.ln()
            pdf.ln(2)
            # Show error/warning message in PDF if present
            if error_message:
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
            
            # Use the helper function to add TOD-wise excess energy breakdown
            tod_excess = add_tod_summary_table(pdf, pdf_data)
            
            # Add financial calculation
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Financial Calculation:', ln=True)
            pdf.set_font('Arial', '', 10)
            
            # Base rate for all excess energy
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
                # Save PDF to a BytesIO object
                pdf_output = io.BytesIO()
                pdf.output(dest=pdf_output)
                pdf_output.seek(0)
                return pdf_output
            except Exception as e:
                error_message = f"Error generating PDF: {str(e)}"
                print(f"ERROR: {error_message}")
                # Try to identify problematic characters
                try:
                    for idx, row in pdf_data.iterrows():
                        for col in row.index:
                            try:
                                str(row[col]).encode('latin-1')
                            except UnicodeEncodeError:
                                print(f"Problematic character in row {idx}, column {col}: {row[col]}")
                except Exception as e2:
                    print(f"Error while debugging: {str(e2)}")
                return None
        
        # Generate day-wise PDF report
        def generate_daywise_pdf(pdf_data, month, year, filename, auto_detect_info=None):
            # Initialize error_message variable
            error_message = None
            
            class PDF(FPDF):
                def header(self):
                    # Logo
                    # self.image('logo.png', 10, 8, 33)
                    # Arial bold 15
                    self.set_font('Arial', 'B', 15)
                    # Move to the right
                    self.cell(80)
                    # Title
                    self.cell(30, 10, 'Energy Adjustment Report (Day-wise)', 0, 0, 'C')
                    # Line break
                    self.ln(20)
                
                def footer(self):
                    # Position at 1.5 cm from bottom
                    self.set_y(-15)
                    # Arial italic 8
                    self.set_font('Arial', 'I', 8)
                    # Page number
                    self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')
            
            # Create PDF object
            pdf = PDF()
            pdf.alias_nb_pages()
            pdf.add_page()
            pdf.set_font('Arial', 'B', 12)
            
            # Add title and consumer information
            pdf.cell(0, 10, f'Energy Adjustment Report (Day-wise) - {get_month_name(month)} {year}', ln=True, align='C')
            pdf.set_font('Arial', '', 10)
            if auto_detect_info:
                pdf.cell(0, 8, f"Note: {auto_detect_info}", ln=True)
            pdf.cell(0, 8, f"Consumer Number: {consumer_number}", ln=True)
            pdf.cell(0, 8, f"Consumer Name: {consumer_name}", ln=True)
            pdf.cell(0, 8, f"T&D Loss: {t_and_d_loss}%", ln=True)
            pdf.cell(0, 8, f"Multiplication Factor: {multiplication_factor}", ln=True)
            
            # Add summary information
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Summary:', ln=True)
            pdf.set_font('Arial', '', 11)
            pdf.cell(0, 8, f"Total Generated Energy (after {t_and_d_loss}% T&D loss): {total_generated_after_loss:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Total Consumed Energy (after multiplication factor): {total_consumed:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Total Excess Energy: {total_excess:.4f} kWh", ln=True)
            pdf.cell(0, 8, f"Status: {excess_status}", ln=True)
            
            # Add day-wise summary table
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Day-wise Summary:', ln=True)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(40, 8, 'Date', 1)
            pdf.cell(50, 8, 'Generated (kWh)', 1)
            pdf.cell(50, 8, 'Consumed (kWh)', 1)
            pdf.cell(50, 8, 'Excess (kWh)', 1)
            pdf.ln()
            pdf.set_font('Arial', '', 10)
            
            # Group by date
            daywise = pdf_data.groupby('Slot_Date').agg({
                'After_Loss': 'sum',
                'Energy_kWh_cons': 'sum',
                'Excess': 'sum'
            })
            
            # Ensure all days of the month are included
            if year and month:
                year_int = int(year)
                month_int = int(month)
                
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
                
                # Create a date range for all days in the month
                all_days = pd.date_range(start=f"{year_int}-{month_int:02d}-01", end=f"{year_int}-{month_int:02d}-{last_day}")
            else:
                # If no month/year specified, use the dates in the data
                all_days = pd.to_datetime(pdf_data['Slot_Date'].unique(), format='%d/%m/%Y')
            
            # Reindex to include all days and fill missing values with 0
            daywise = daywise.reindex(all_days.strftime('%d/%m/%Y'), fill_value=0).reset_index()
            daywise = daywise.rename(columns={'index': 'Slot_Date'})
            for idx, row in daywise.iterrows():
                pdf.cell(40, 8, row['Slot_Date'], 1)
                pdf.cell(50, 8, f"{row['After_Loss']:.4f}", 1)
                pdf.cell(50, 8, f"{row['Energy_kWh_cons']:.4f}", 1)
                pdf.cell(50, 8, f"{row['Excess']:.4f}", 1)
                pdf.ln()
            pdf.ln(2)
            
            # Use the helper function to add TOD-wise excess energy breakdown
            tod_excess = add_tod_summary_table(pdf, pdf_data)
            
            # Add financial calculation
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'Financial Calculation:', ln=True)
            pdf.set_font('Arial', '', 10)
            
            # Base rate for all excess energy
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
                # Save PDF to a BytesIO object
                pdf_output = io.BytesIO()
                pdf.output(dest=pdf_output)
                pdf_output.seek(0)
                return pdf_output
            except Exception as e:
                error_message = f"Error generating PDF: {str(e)}"
                print(f"ERROR: {error_message}")
                # Try to identify problematic characters
                try:
                    for idx, row in pdf_data.iterrows():
                        for col in row.index:
                            try:
                                str(row[col]).encode('latin-1')
                            except UnicodeEncodeError:
                                print(f"Problematic character in row {idx}, column {col}: {row[col]}")
                except Exception as e2:
                    print(f"Error while debugging: {str(e2)}")
                return None
        
        # Generate PDF based on selected options
        if show_daywise:
            pdf_output = generate_daywise_pdf(merged, month, year, 'energy_adjustment_daywise.pdf')
            if pdf_output:
                return send_file(
                    pdf_output,
                    as_attachment=True,
                    download_name='energy_adjustment_daywise.pdf',
                    mimetype='application/pdf'
                )
            else:
                return render_template('index.html', error="Error generating PDF. Check server logs for details.")
        else:
            pdf_output = generate_pdf(merged, month, year, 'energy_adjustment.pdf')
            if pdf_output:
                return send_file(
                    pdf_output,
                    as_attachment=True,
                    download_name='energy_adjustment.pdf',
                    mimetype='application/pdf'
                )
            else:
                return render_template('index.html', error="Error generating PDF. Check server logs for details.")
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, port=5002)