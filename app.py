from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os
import math

PDF_AUTHOR_NAME = "Er.Aravind MRT VREDC"


class AuthorPDF(FPDF):
    def __init__(self, author_name=PDF_AUTHOR_NAME, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.author_name = author_name

    def header(self):
        if self.author_name:
            self.set_font('Arial', 'I', 10)
            self.set_text_color(80, 80, 80)
            self.cell(0, 10, self.author_name, 0, 0, 'R')
            self.ln(5)
            self.set_text_color(0, 0, 0)


app = Flask(__name__)


WHEELING_RATE_PER_KWH = 2.34


def compute_wheeling_components(total_excess_kwh, t_and_d_loss_percent):
    """Return (reference_kwh, charges) for wheeling deduction."""
    try:
        loss_pct = float(t_and_d_loss_percent or 0)
    except Exception:
        loss_pct = 0.0

    reference_kwh = 0.0
    if loss_pct > 0 and loss_pct < 100:
        reference_kwh = (total_excess_kwh * loss_pct) / (100 - loss_pct)

    charges = reference_kwh * WHEELING_RATE_PER_KWH
    return reference_kwh, charges

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
        enable_iex = request.form.get('enable_iex') == '1'
        enable_cpp = request.form.get('enable_cpp') == '1'
        t_and_d_loss = request.form.get('t_and_d_loss')
        cpp_t_and_d_loss = request.form.get('cpp_t_and_d_loss')
        consumer_number = request.form['consumer_number']
        consumer_name = request.form['consumer_name']
        multiplication_factor = float(request.form.get('multiplication_factor', 1))
        auto_detect_month = request.form.get('auto_detect_month') == '1'
        month = request.form.get('month', '')
        year = request.form.get('year', '')
        date_filter = request.form.get('date', '').strip()  # Optional single date filter (format: dd/mm/yyyy)

        # Get uploaded files
        generated_files = request.files.getlist('generated_files')
        cpp_files = request.files.getlist('cpp_files')
        consumed_files = request.files.getlist('consumed_files')
        
        # Filter out empty file inputs and check if sources are enabled
        generated_files = [f for f in generated_files if f.filename] if enable_iex else []
        cpp_files = [f for f in cpp_files if f.filename] if enable_cpp else []
        
        # Validate that at least one generation source is enabled
        if not enable_iex and not enable_cpp:
            return render_template('index.html', error="Please enable at least one generation source (I.E.X or C.P.P).")
        
        # Validate files and T&D loss values for enabled generation sources
        if enable_iex:
            if not generated_files:
                return render_template('index.html', error="Please select I.E.X generation files since I.E.X is enabled.")
            if not t_and_d_loss or t_and_d_loss.strip() == '':
                return render_template('index.html', error="Please enter T&D Loss for I.E.X generation since I.E.X is enabled.")
        
        if enable_cpp:
            if not cpp_files:
                return render_template('index.html', error="Please select C.P.P generation files since C.P.P is enabled.")
            if not cpp_t_and_d_loss or cpp_t_and_d_loss.strip() == '':
                return render_template('index.html', error="Please enter T&D Loss for C.P.P generation since C.P.P is enabled.")
        
        # Convert T&D loss values to float
        t_and_d_loss = float(t_and_d_loss) if t_and_d_loss and t_and_d_loss.strip() else 0
        cpp_t_and_d_loss = float(cpp_t_and_d_loss) if cpp_t_and_d_loss and cpp_t_and_d_loss.strip() else 0
        
        if not consumed_files:
            return render_template('index.html', error="No consumed energy Excel files were uploaded.")
            
        # Process I.E.X generated energy Excel files (if provided)
        gen_df = None
        if generated_files:
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
            if gen_dfs:
                gen_df = pd.concat(gen_dfs, ignore_index=True)
                gen_df = gen_df.iloc[:, :3]
                gen_df.columns = ['Date', 'Time', 'Energy_MW']
                # Strip whitespace from Date and Time columns
                gen_df['Date'] = gen_df['Date'].astype(str).str.strip()
                gen_df['Time'] = gen_df['Time'].astype(str).str.strip()
                
                # Convert Energy_MW to numeric, handling string values
                gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
                nan_count = gen_df['Energy_MW'].isna().sum()
                if nan_count > 0:
                    print(f"WARNING: {nan_count} non-numeric Energy_MW values found in I.E.X files and converted to NaN")
                
                # Standardize date format to yyyy-mm-dd for robust filtering
                gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
                gen_df['Source_Type'] = 'I.E.X'
        
        # Process C.P.P (Captive Power Purchase) files (if provided)
        cpp_df = None
        if cpp_files:
            cpp_dfs = []
            for cpp_file in cpp_files:
                try:
                    temp_df = pd.read_excel(cpp_file, header=0)
                    if temp_df.shape[1] < 3:
                        return render_template('index.html', error=f"C.P.P energy Excel file '{cpp_file.filename}' must have at least 3 columns: Date, Time, and Energy in MW.")
                    
                    # Add filename to help with debugging
                    temp_df['Source_File'] = cpp_file.filename
                    temp_df['Source_Type'] = 'C.P.P'
                    cpp_dfs.append(temp_df)
                except Exception as e:
                    return render_template('index.html', error=f"Error reading C.P.P energy file '{cpp_file.filename}': {str(e)}")
            
            # Process C.P.P data if files were uploaded
            if cpp_dfs:
                cpp_df = pd.concat(cpp_dfs, ignore_index=True)
                cpp_df = cpp_df.iloc[:, :3]
                cpp_df.columns = ['Date', 'Time', 'Energy_MW']
                cpp_df['Date'] = cpp_df['Date'].astype(str).str.strip()
                cpp_df['Time'] = cpp_df['Time'].astype(str).str.strip()
                
                # Convert Energy_MW to numeric, handling string values
                cpp_df['Energy_MW'] = pd.to_numeric(cpp_df['Energy_MW'], errors='coerce')
                nan_count = cpp_df['Energy_MW'].isna().sum()
                if nan_count > 0:
                    print(f"WARNING: {nan_count} non-numeric Energy_MW values found in C.P.P files and converted to NaN")
                
                cpp_df['Date'] = pd.to_datetime(cpp_df['Date'], errors='coerce', dayfirst=True)
                cpp_df['Source_Type'] = 'C.P.P'
        
        # Combine I.E.X and C.P.P data if both exist
        if gen_df is not None and cpp_df is not None:
            combined_gen_df = pd.concat([gen_df, cpp_df], ignore_index=True)
            print(f"Combined {len(gen_df)} I.E.X records with {len(cpp_df)} C.P.P records")
        elif gen_df is not None:
            combined_gen_df = gen_df
            print(f"Using {len(gen_df)} I.E.X records only")
        elif cpp_df is not None:
            combined_gen_df = cpp_df
            print(f"Using {len(cpp_df)} C.P.P records only")
        else:
            return render_template('index.html', error="No valid generation energy files were found.")
        
        gen_df = combined_gen_df
        
        # Debug: Check combined data totals
        print(f"\n=== COMBINED DATA DEBUG ===")
        print(f"Total combined records: {len(gen_df)}")
        print(f"Energy_MW column data type: {gen_df['Energy_MW'].dtype}")
        print(f"Sample Energy_MW values: {gen_df['Energy_MW'].head(10).tolist()}")
        
        # Convert Energy_MW to numeric for debugging (handle string values)
        gen_df['Energy_MW'] = pd.to_numeric(gen_df['Energy_MW'], errors='coerce')
        nan_count = gen_df['Energy_MW'].isna().sum()
        if nan_count > 0:
            print(f"WARNING: {nan_count} Energy_MW values could not be converted to numbers!")
            print(f"Sample problematic values: {gen_df[gen_df['Energy_MW'].isna()][['Date', 'Time', 'Energy_MW']].head()}")
        
        if 'Source_Type' in gen_df.columns:
            iex_count = len(gen_df[gen_df['Source_Type'] == 'I.E.X'])
            cpp_count = len(gen_df[gen_df['Source_Type'] == 'C.P.P'])
            print(f"I.E.X records in combined: {iex_count}")
            print(f"C.P.P records in combined: {cpp_count}")
            if iex_count > 0:
                iex_total_mw = gen_df[gen_df['Source_Type'] == 'I.E.X']['Energy_MW'].sum()
                print(f"Total I.E.X MW in combined data: {iex_total_mw:.4f} MW")
            if cpp_count > 0:
                cpp_total_mw = gen_df[gen_df['Source_Type'] == 'C.P.P']['Energy_MW'].sum()
                print(f"Total C.P.P MW in combined data: {cpp_total_mw:.4f} MW")
        print("=== END COMBINED DATA DEBUG ===\n")
        
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
                month = str(int(unique_months[0]))
                print(f"Auto-detected month: {month} ({get_month_name(month)})")
            elif len(unique_months) > 1 and not month:
                # If multiple months, use the most frequent one
                month = str(int(gen_df['Date'].dt.month.value_counts().idxmax()))
                print(f"Multiple months detected, using most frequent: {month} ({get_month_name(month)})")
            
            if len(unique_years) == 1 and not year:
                year = str(int(unique_years[0]))
                print(f"Auto-detected year: {year}")
            elif len(unique_years) > 1 and not year:
                # If multiple years, use the most frequent one
                year = str(int(gen_df['Date'].dt.year.value_counts().idxmax()))
                print(f"Multiple years detected, using most frequent: {year}")
                
            # Add information to be displayed in PDF
            cpp_count = len(cpp_files) if cpp_files else 0
            iex_count = len(generated_files) if generated_files else 0
            auto_detect_info = f"Auto-detected from {iex_count} I.E.X, {cpp_count} C.P.P, and {len(consumed_files)} consumed files"
        
        # Debug: Check data before filtering
        print(f"\n=== BEFORE FILTERING DEBUG ===")
        print(f"Total records before filtering: {len(gen_df)}")
        if 'Source_Type' in gen_df.columns:
            iex_before = len(gen_df[gen_df['Source_Type'] == 'I.E.X'])
            cpp_before = len(gen_df[gen_df['Source_Type'] == 'C.P.P'])
            print(f"I.E.X records before filtering: {iex_before}")
            print(f"C.P.P records before filtering: {cpp_before}")
            if iex_before > 0:
                iex_mw_before = gen_df[gen_df['Source_Type'] == 'I.E.X']['Energy_MW'].sum()
                print(f"I.E.X MW before filtering: {iex_mw_before:.4f} MW")
            if cpp_before > 0:
                cpp_mw_before = gen_df[gen_df['Source_Type'] == 'C.P.P']['Energy_MW'].sum()
                print(f"C.P.P MW before filtering: {cpp_mw_before:.4f} MW")
        print(f"Date range in data: {gen_df['Date'].min()} to {gen_df['Date'].max()}")
        print(f"Filter parameters: year={year}, month={month}, date_filter={date_filter}")
        print("=== END BEFORE FILTERING DEBUG ===\n")
        
        # Filter by year/month with custom slot logic (handle slot ranges in Time column)
        filtered_gen = gen_df.copy()
        if year and month:
            try:
                # Handle potential float strings by converting to float first, then int
                year_int = int(float(year))
                month_int = int(float(month))
            except ValueError as e:
                return render_template('index.html', error=f"Invalid year or month value. Year: '{year}', Month: '{month}'. Error: {str(e)}")
            
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
                try:
                    year_int = int(float(year))
                    filtered_gen = filtered_gen[filtered_gen['Date'].dt.year == year_int]
                except ValueError:
                    return render_template('index.html', error=f"Invalid year value: '{year}'")
            if month:
                try:
                    month_int = int(float(month))
                    filtered_gen = filtered_gen[filtered_gen['Date'].dt.month == month_int]
                except ValueError:
                    return render_template('index.html', error=f"Invalid month value: '{month}'")
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
        # Debug: Check MW to kWh conversion
        print(f"\n=== MW TO kWH CONVERSION DEBUG ===")
        print(f"Number of slots in gen_df: {len(gen_df)}")
        print(f"MW sample values: {gen_df['Energy_MW'].head(10).tolist()}")
        print(f"MW data type: {gen_df['Energy_MW'].dtype}")
        print(f"Total MW in gen_df: {gen_df['Energy_MW'].sum():.4f} MW")
        
        gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
        
        print(f"Total kWh after conversion: {gen_df['Energy_kWh'].sum():.4f} kWh")
        print(f"Manual check: {gen_df['Energy_MW'].sum():.4f} MW * 250 = {gen_df['Energy_MW'].sum() * 250:.4f} kWh")
        print(f"Any NaN values in Energy_MW? {gen_df['Energy_MW'].isna().sum()}")
        print(f"Any zero values in Energy_MW? {(gen_df['Energy_MW'] == 0).sum()}")
        print("=== END MW TO kWH DEBUG ===\n")
        
        # Apply separate T&D losses based on source type
        def apply_td_loss(row):
            if row['Source_Type'] == 'I.E.X':
                return row['Energy_kWh'] * (1 - t_and_d_loss / 100) if t_and_d_loss > 0 else row['Energy_kWh']
            elif row['Source_Type'] == 'C.P.P':
                return row['Energy_kWh'] * (1 - cpp_t_and_d_loss / 100) if cpp_t_and_d_loss > 0 else row['Energy_kWh']
            else:
                return row['Energy_kWh']  # Fallback, no loss applied
        
        gen_df['After_Loss'] = gen_df.apply(apply_td_loss, axis=1)
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
            try:
                year_int = int(float(year))
                month_int = int(float(month))
            except ValueError as e:
                return render_template('index.html', error=f"Invalid year or month value in consumption data. Year: '{year}', Month: '{month}'. Error: {str(e)}")
            
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
                try:
                    year_int = int(float(year))
                    filtered_cons = filtered_cons[filtered_cons['Date'].dt.year == year_int]
                except ValueError:
                    return render_template('index.html', error=f"Invalid year value in consumption filtering: '{year}'")
            if month:
                try:
                    month_int = int(float(month))
                    filtered_cons = filtered_cons[filtered_cons['Date'].dt.month == month_int]
                except ValueError:
                    return render_template('index.html', error=f"Invalid month value in consumption filtering: '{month}'")
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
        
        # Sequential adjustment logic: First I.E.X, then C.P.P
        # Separate I.E.X and C.P.P data for sequential adjustment
        iex_df = gen_df[gen_df['Source_Type'] == 'I.E.X'].copy() if enable_iex else pd.DataFrame()
        cpp_df_only = gen_df[gen_df['Source_Type'] == 'C.P.P'].copy() if enable_cpp else pd.DataFrame()
        
        # Debug: Check data separation for I.E.X and C.P.P
        print(f"\n=== DATA SEPARATION DEBUG ===")
        if enable_iex:
            print(f"I.E.X records: {len(iex_df)}")
            print(f"I.E.X total MW: {iex_df['Energy_MW'].sum():.4f} MW")
            print(f"I.E.X total kWh (before loss): {iex_df['Energy_kWh'].sum():.4f} kWh")
        else:
            print("I.E.X is disabled")
            
        if enable_cpp:
            print(f"C.P.P records: {len(cpp_df_only)}")
            print(f"C.P.P total MW: {cpp_df_only['Energy_MW'].sum():.4f} MW")
            print(f"C.P.P total kWh (before loss): {cpp_df_only['Energy_kWh'].sum():.4f} kWh")
        else:
            print("C.P.P is disabled")
        print("=== END DATA SEPARATION DEBUG ===\n")
        
        # Build all slot combinations from consumption and both generation sources
        cons_slots_set = set((d, t) for d, t in zip(cons_df['Slot_Date'], cons_df['Slot_Time']))
        iex_slots_set = set((d, t) for d, t in zip(iex_df['Slot_Date'], iex_df['Slot_Time'])) if not iex_df.empty else set()
        cpp_slots_set = set((d, t) for d, t in zip(cpp_df_only['Slot_Date'], cpp_df_only['Slot_Time'])) if not cpp_df_only.empty else set()
        
        all_slots = pd.DataFrame(
            list(cons_slots_set | iex_slots_set | cpp_slots_set),
            columns=['Slot_Date', 'Slot_Time'])
        
        # Merge consumption data first
        merged = pd.merge(all_slots, cons_df[['Slot_Date', 'Slot_Time', 'Energy_kWh']], on=['Slot_Date', 'Slot_Time'], how='left')
        merged['Energy_kWh_cons'] = merged['Energy_kWh'].fillna(0)
        merged.drop('Energy_kWh', axis=1, inplace=True)
        
        # Merge I.E.X data
        if not iex_df.empty:
            iex_merge = iex_df[['Slot_Date', 'Slot_Time', 'After_Loss', 'Energy_kWh']].copy()
            iex_merge.columns = ['Slot_Date', 'Slot_Time', 'IEX_After_Loss', 'IEX_Energy_kWh']
            merged = pd.merge(merged, iex_merge, on=['Slot_Date', 'Slot_Time'], how='left')
            merged['IEX_After_Loss'] = merged['IEX_After_Loss'].fillna(0)
            merged['IEX_Energy_kWh'] = merged['IEX_Energy_kWh'].fillna(0)
        else:
            merged['IEX_After_Loss'] = 0
            merged['IEX_Energy_kWh'] = 0
        
        # Merge C.P.P data
        if not cpp_df_only.empty:
            cpp_merge = cpp_df_only[['Slot_Date', 'Slot_Time', 'After_Loss', 'Energy_kWh']].copy()
            cpp_merge.columns = ['Slot_Date', 'Slot_Time', 'CPP_After_Loss', 'CPP_Energy_kWh']
            merged = pd.merge(merged, cpp_merge, on=['Slot_Date', 'Slot_Time'], how='left')
            merged['CPP_After_Loss'] = merged['CPP_After_Loss'].fillna(0)
            merged['CPP_Energy_kWh'] = merged['CPP_Energy_kWh'].fillna(0)
        else:
            merged['CPP_After_Loss'] = 0
            merged['CPP_Energy_kWh'] = 0
        
        # Sequential Adjustment Calculation
        # Step 1: I.E.X adjustment first
        merged['IEX_Adjustment'] = merged[['IEX_After_Loss', 'Energy_kWh_cons']].min(axis=1)
        merged['IEX_Excess'] = (merged['IEX_After_Loss'] - merged['Energy_kWh_cons']).apply(lambda x: max(0, x))
        
        # Step 2: Calculate remaining consumption after I.E.X adjustment
        merged['Remaining_Consumption'] = (merged['Energy_kWh_cons'] - merged['IEX_Adjustment']).apply(lambda x: max(0, x))
        
        # Step 3: C.P.P adjustment with remaining consumption
        merged['CPP_Adjustment'] = merged[['CPP_After_Loss', 'Remaining_Consumption']].min(axis=1)
        merged['CPP_Excess'] = (merged['CPP_After_Loss'] - merged['Remaining_Consumption']).apply(lambda x: max(0, x))
        
        # Step 4: Total calculations
        merged['Total_Excess'] = merged['IEX_Excess'] + merged['CPP_Excess']
        merged['Total_Generated_After_Loss'] = merged['IEX_After_Loss'] + merged['CPP_After_Loss']
        merged['Total_Generated_Before_Loss'] = merged['IEX_Energy_kWh'] + merged['CPP_Energy_kWh']
        
        # Debug: Print sequential calculation summary
        print("\n=== SEQUENTIAL ADJUSTMENT DEBUG ===")
        print(f"Enable I.E.X: {enable_iex}, Enable C.P.P: {enable_cpp}")
        if enable_iex and enable_cpp:
            print(f"Total I.E.X Before Loss: {merged['IEX_Energy_kWh'].sum():.4f} kWh")
            print(f"Total I.E.X After Loss: {merged['IEX_After_Loss'].sum():.4f} kWh")
            print(f"Total C.P.P Before Loss: {merged['CPP_Energy_kWh'].sum():.4f} kWh")
            print(f"Total C.P.P After Loss: {merged['CPP_After_Loss'].sum():.4f} kWh")
            print(f"Total Consumption: {merged['Energy_kWh_cons'].sum():.4f} kWh")
            print(f"Total I.E.X Excess: {merged['IEX_Excess'].sum():.4f} kWh")
            print(f"Total C.P.P Excess: {merged['CPP_Excess'].sum():.4f} kWh")
            print(f"Total Combined Excess: {merged['Total_Excess'].sum():.4f} kWh")
            print(f"Remaining Consumption Total: {merged['Remaining_Consumption'].sum():.4f} kWh")
        print("=== END DEBUG ===\n")
        
        # For backward compatibility with existing PDF code
        merged['After_Loss'] = merged['Total_Generated_After_Loss']
        merged['Energy_kWh_gen'] = merged['Total_Generated_Before_Loss']
        merged['Excess'] = merged['Total_Excess']
        # Track missing slots for reporting
        merged['Missing_Info'] = ''
        merged['is_missing_iex'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in iex_slots_set, axis=1)
        merged['is_missing_cpp'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in cpp_slots_set, axis=1)
        merged['is_missing_cons'] = ~merged.apply(lambda row: (row['Slot_Date'], row['Slot_Time']) in cons_slots_set, axis=1)
        
        if enable_iex:
            merged.loc[merged['is_missing_iex'], 'Missing_Info'] += '[Missing in I.E.X] '
        if enable_cpp:
            merged.loc[merged['is_missing_cpp'], 'Missing_Info'] += '[Missing in C.P.P] '
        merged.loc[merged['is_missing_cons'], 'Missing_Info'] += '[Missing in CONSUMED] '
        merged.drop(['is_missing_iex', 'is_missing_cpp', 'is_missing_cons'], axis=1, inplace=True)
        # Compose error/warning message for PDF
        error_message = ''
        local_missing_days_msg = globals().get('missing_days_msg', '')
        local_slot_mismatch_msg = globals().get('slot_mismatch_msg', '')
        if local_missing_days_msg or local_slot_mismatch_msg:
            error_message = local_missing_days_msg + local_slot_mismatch_msg + '\nProceeding with only the matching days and slots (missing slots filled with zero).'

        # Check for missing days in either file (priority: match Slot_Date between files)
        cons_days = set(cons_df['Slot_Date'].unique())
        iex_days = set(iex_df['Slot_Date'].unique()) if not iex_df.empty else set()
        cpp_days = set(cpp_df_only['Slot_Date'].unique()) if not cpp_df_only.empty else set()
        
        all_gen_days = iex_days | cpp_days
        common_days = cons_days & all_gen_days
        
        missing_in_gen = sorted(list(cons_days - all_gen_days))
        missing_in_cons = sorted(list(all_gen_days - cons_days))
        
        missing_days_msg = ""
        if missing_in_gen:
            missing_days_msg += f"Warning: The following days are present in CONSUMED but missing in GENERATION: {', '.join(missing_in_gen)}\n"
        if missing_in_cons:
            missing_days_msg += f"Warning: The following days are present in GENERATION but missing in CONSUMED: {', '.join(missing_in_cons)}\n"
        
        # Check for missing slots (time intervals) for common days
        slot_mismatch_msg = ""
        for day in sorted(common_days):
            cons_slots = set(cons_df[cons_df['Slot_Date'] == day]['Slot_Time'])
            iex_slots = set(iex_df[iex_df['Slot_Date'] == day]['Slot_Time']) if not iex_df.empty else set()
            cpp_slots = set(cpp_df_only[cpp_df_only['Slot_Date'] == day]['Slot_Time']) if not cpp_df_only.empty else set()
            all_gen_slots = iex_slots | cpp_slots
            
            missing_slots_in_gen = cons_slots - all_gen_slots
            missing_slots_in_cons = all_gen_slots - cons_slots
            
            if missing_slots_in_gen:
                slot_mismatch_msg += f"Day {day}: Slots in CONSUMED but missing in GENERATION: {', '.join(sorted(missing_slots_in_gen))}\n"
            if missing_slots_in_cons:
                slot_mismatch_msg += f"Day {day}: Slots in GENERATION but missing in CONSUMED: {', '.join(sorted(missing_slots_in_cons))}\n"

        # If there are missing days or slots, add warning message
        warning_msg = ''
        if missing_days_msg or slot_mismatch_msg:
            warning_msg = missing_days_msg + slot_mismatch_msg
            warning_msg += "\nProceeding with only the matching days and slots (intersection)."
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
                
                # Morning peak: 6:00 AM - 10:00 AM (C1) - EXCLUDES 10:00-10:15 slot
                if 6 <= hour < 10:
                    return 'C1'
                # Evening peak: 6:00 PM - 10:00 PM (C2) - EXCLUDES 22:00-22:15 slot
                elif 18 <= hour < 22:
                    return 'C2'
                # Normal hours: 5:00 AM - 6:00 AM + 10:00 AM to 6:00 PM (C4) - INCLUDES 10:00 and 22:00 slots
                elif (5 <= hour < 6) or (10 <= hour < 18):
                    return 'C4'
                # Night hours: 22:00 PM to 5:00 AM (C5)
                elif (hour >= 22) or (hour < 5):
                    return 'C5'
                else:
                    return 'Unknown'
            except Exception:
                return 'Unknown'
        
        merged['TOD_Category'] = merged['Slot_Time'].apply(classify_tod)
        
        # Debug: Print some TOD classifications to verify fix
        print(f"\n=== TOD CLASSIFICATION DEBUG ===")
        test_times = merged[merged['Slot_Time'].str.contains('09:45|10:00|10:15|21:45|22:00|22:15', na=False)]
        if not test_times.empty:
            for _, row in test_times.head(10).iterrows():
                print(f"Time: {row['Slot_Time']} -> TOD: {row['TOD_Category']}")
        print("=== END TOD DEBUG ===\n")
        
        # Group excess energy by TOD category using sequential adjustment totals
        tod_excess = merged.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
        
        # Calculate financial values using sequential adjustment total with rounded values for consistency
        total_excess_financial = merged['Total_Excess'].sum()
        
        # Helper function for consistent rounding throughout the application
        def round_kwh_financial(value):
            return int(value + 0.5) if value >= 0 else int(value - 0.5)
        
        # Round the total for financial calculations to match table display values
        total_excess_financial_rounded = round_kwh_financial(total_excess_financial)
        
        # Base rate for all excess energy using rounded values
        base_rate = 7.25  # rupees per kWh
        base_amount = total_excess_financial_rounded * base_rate
        
        # Additional charges for specific TOD categories using rounded values
        c1_c2_excess_raw = tod_excess.loc[tod_excess['TOD_Category'].isin(['C1', 'C2']), 'Total_Excess'].sum()
        c1_c2_excess = round_kwh_financial(c1_c2_excess_raw)
        c1_c2_additional = c1_c2_excess * 1.8125  # rupees per kWh
        
        c5_excess_raw = tod_excess.loc[tod_excess['TOD_Category'] == 'C5', 'Total_Excess'].sum()
        c5_excess = round_kwh_financial(c5_excess_raw)
        c5_additional = c5_excess * 0.3625  # rupees per kWh
        
        # Calculate total amount
        total_amount = base_amount + c1_c2_additional + c5_additional
        
        # Calculate E-Tax (5% of total amount)
        etax = total_amount * 0.05
        
        # Calculate total amount with E-Tax
        total_with_etax = total_amount + etax
        
        # Calculate IEX excess for specific charges using rounded values
        iex_excess_financial_raw = merged['IEX_Excess'].sum()
        iex_excess_financial = round_kwh_financial(iex_excess_financial_raw)
        
        # Calculate negative factors using rounded values
        etax_on_iex = total_excess_financial_rounded * 0.1
        cross_subsidy_surcharge = iex_excess_financial * 1.92  # Only for IEX excess
        
        wheeling_reference_kwh, wheeling_charges = compute_wheeling_components(
            total_excess_financial_rounded,
            t_and_d_loss,
        )
        
        # Calculate final amount to be collected
        final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
        
        # Round up final amount to next highest value
        final_amount_rounded = math.ceil(final_amount)
        
        merged.drop(['Slot_Date_dt', 'Slot_Time_min'], axis=1, inplace=True)
        # Totals using sequential adjustment calculations
        sum_injection = merged['Energy_kWh_gen'].sum()  # Generated before loss
        total_generated_after_loss = merged['After_Loss'].sum()
        total_consumed = merged['Energy_kWh_cons'].sum()
        
        # Use the new sequential adjustment totals instead of old combined logic
        total_excess = merged['Total_Excess'].sum()  # Use Total_Excess from sequential calculation
        comparison = sum_injection - total_generated_after_loss
        
        # For PDF, show all slots or only excess slots (using Total_Excess)
        merged_excess = merged[merged['Total_Excess'] > 0].copy()  # Filter by Total_Excess
        merged_all = merged.copy()
        
        # DEBUG: Show difference between full totals and excess-only totals
        print(f"\n=== EXCESS VS ALL TOTALS DEBUG ===")
        excess_iex_total = merged_excess['IEX_Energy_kWh'].sum() if 'IEX_Energy_kWh' in merged_excess.columns else 0
        excess_cpp_total = merged_excess['CPP_Energy_kWh'].sum() if 'CPP_Energy_kWh' in merged_excess.columns else 0
        excess_generation_total = excess_iex_total + excess_cpp_total
        
        all_iex_total = merged['IEX_Energy_kWh'].sum()
        all_cpp_total = merged['CPP_Energy_kWh'].sum()
        all_generation_total = all_iex_total + all_cpp_total
        
        print(f"EXCESS SLOTS ONLY - I.E.X: {excess_iex_total:.4f} kWh, C.P.P: {excess_cpp_total:.4f} kWh, Total: {excess_generation_total:.4f} kWh")
        print(f"ALL SLOTS - I.E.X: {all_iex_total:.4f} kWh, C.P.P: {all_cpp_total:.4f} kWh, Total: {all_generation_total:.4f} kWh")
        print(f"Excess rows: {len(merged_excess)}, All rows: {len(merged)}")
        print("=== END EXCESS VS ALL TOTALS DEBUG ===\n")
        
        # CORRECTED: For excess PDF, use only excess slot totals; for all PDF, use sequential totals
        sum_injection_excess = excess_generation_total  # Only excess slots
        total_generated_after_loss_excess = merged_excess['IEX_After_Loss'].sum() + merged_excess['CPP_After_Loss'].sum()  # Only excess slots
        # Use the total consumed energy from all slots for consistency across all PDFs
        total_consumed_excess = merged['Energy_kWh_cons'].sum()  # Total consumption from all slots
        total_excess_excess = merged_excess['Total_Excess'].sum()  # Use Total_Excess
        
        sum_injection_all = all_generation_total  # All sequential totals
        total_generated_after_loss_all = merged_all['IEX_After_Loss'].sum() + merged_all['CPP_After_Loss'].sum()  # All sequential totals
        total_consumed_all = merged_all['Energy_kWh_cons'].sum()
        total_excess_all = merged_all['Total_Excess'].sum()  # Use Total_Excess
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
        def generate_pdf(pdf_data, sum_injection, total_generated_after_loss, comparison, total_consumed, total_excess, excess_status, filename, auto_detect=auto_detect_month, gen_files=generated_files, cpp_files=cpp_files, cons_files=consumed_files, full_totals=None):
            # Debug: Check what data PDF generation receives
            print(f"\n=== PDF GENERATION DEBUG ===")
            print(f"PDF received total_excess parameter: {total_excess:.4f} kWh")
            print(f"PDF received total_generated_after_loss: {total_generated_after_loss:.4f} kWh")
            print(f"PDF received total_consumed: {total_consumed:.4f} kWh")
            print(f"PDF data shape: {pdf_data.shape}")
            if 'IEX_Energy_kWh' in pdf_data.columns:
                print(f"PDF data IEX_Energy_kWh sum: {pdf_data['IEX_Energy_kWh'].sum():.4f} kWh")
            if 'CPP_Energy_kWh' in pdf_data.columns:
                print(f"PDF data CPP_Energy_kWh sum: {pdf_data['CPP_Energy_kWh'].sum():.4f} kWh")
            if 'Total_Excess' in pdf_data.columns:
                print(f"PDF data Total_Excess sum: {pdf_data['Total_Excess'].sum():.4f} kWh")
            if 'Excess' in pdf_data.columns:
                print(f"PDF data Excess sum: {pdf_data['Total_Excess'].sum():.4f} kWh")
            if 'IEX_Excess' in pdf_data.columns:
                print(f"PDF data IEX_Excess sum: {pdf_data['IEX_Excess'].sum():.4f} kWh")
            if 'CPP_Excess' in pdf_data.columns:
                print(f"PDF data CPP_Excess sum: {pdf_data['CPP_Excess'].sum():.4f} kWh")
            if full_totals:
                print(f"Full totals provided: IEX_Before={full_totals.get('iex_before', 0):.4f}, CPP_Before={full_totals.get('cpp_before', 0):.4f}")
                print(f"Full totals provided: IEX_After={full_totals.get('iex_after', 0):.4f}, CPP_After={full_totals.get('cpp_after', 0):.4f}")
                print(f"Full totals provided: IEX_Excess={full_totals.get('iex_excess', 0):.4f}, CPP_Excess={full_totals.get('cpp_excess', 0):.4f}")
            print("=== END PDF GENERATION DEBUG ===\n")
            
            # Import datetime for timestamp
            from datetime import datetime
            
            pdf = AuthorPDF()
            pdf.set_margins(20, 20, 20)  # Set proper margins: left, top, right (20mm each)
            pdf.set_auto_page_break(auto=True, margin=20)  # Auto page break with bottom margin
            pdf.add_page()
            
            # FIRST PAGE - DESCRIPTION AND INFORMATION ONLY
            pdf.set_font('Arial', 'B', 16)  # Larger title font
            pdf.cell(0, 15, 'Energy Adjustment Report', ln=True, align='C')
            pdf.ln(10)
            
            # Consumer Information Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Consumer Information:', ln=True)
            pdf.set_font('Arial', '', 12)
            pdf.cell(0, 8, f'Consumer Number: {consumer_number}', ln=True)
            pdf.cell(0, 8, f'Consumer Name: {consumer_name}', ln=True)
            pdf.cell(0, 8, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
            pdf.ln(5)
            
            # Technical Parameters Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Technical Parameters:', ln=True)
            pdf.set_font('Arial', '', 12)
            
            # Display T&D losses based on what sources are used
            if generated_files and cpp_files:
                pdf.cell(0, 8, f'I.E.X T&D Loss (%): {t_and_d_loss}', ln=True)
                pdf.cell(0, 8, f'C.P.P T&D Loss (%): {cpp_t_and_d_loss}', ln=True)
            elif generated_files:
                pdf.cell(0, 8, f'I.E.X T&D Loss (%): {t_and_d_loss}', ln=True)
            elif cpp_files:
                pdf.cell(0, 8, f'C.P.P T&D Loss (%): {cpp_t_and_d_loss}', ln=True)
            pdf.ln(5)
            
            # Data Sources Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Data Sources:', ln=True)
            pdf.set_font('Arial', '', 12)
            if generated_files:
                pdf.cell(0, 8, f'Generated Energy Files: {len(generated_files)} file(s)', ln=True)
            if cpp_files:
                pdf.cell(0, 8, f'C.P.P Energy Files: {len(cpp_files)} file(s)', ln=True)
            if consumed_files:
                pdf.cell(0, 8, f'Consumed Energy Files: {len(consumed_files)} file(s)', ln=True)
            pdf.ln(5)
            
            # Report Information Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Report Information:', ln=True)
            pdf.set_font('Arial', '', 12)
            if auto_detect:
                pdf.cell(0, 8, f'Data Period: Auto-detected from uploaded files', ln=True)
            if auto_detect_info:
                pdf.cell(0, 8, f'Period Details: {auto_detect_info}', ln=True)
            pdf.cell(0, 8, f'Report Generated: {datetime.now().strftime("%d/%m/%Y at %H:%M:%S")}', ln=True)
            pdf.ln(10)
            
            # Calculation Methodology Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Calculation Methodology:', ln=True)
            pdf.set_font('Arial', '', 12)
            
            methodology_text = [
                "1. Energy Adjustment = Net Generated Energy - Adjusted Consumed Energy",
                "2. Adjusted Consumed Energy = Consumed Energy × Multiplication Factor",
                "3. Net Generated Energy = Generated Energy × (1 - T&D Loss %)",
                "4. Financial calculations use Time of Day (TOD) rates",
                "5. All calculations maintain precision and consistency across tables and summaries"
            ]
            
            for line in methodology_text:
                pdf.cell(0, 6, line, ln=True)
            
            pdf.ln(10)
            
            # Add page break before tables
            pdf.add_page()
            
            # SECOND PAGE AND BEYOND - TABLES START HERE
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Energy Adjustment Tables', ln=True, align='C')
            pdf.ln(10)
            
            # Set appropriate font for table headers based on energy sources
            
            # Function to add table headers with proper text wrapping
            def add_table_headers():
                if generated_files and cpp_files:
                    # Sequential adjustment table with detailed columns - improved headers
                    pdf.set_font('Arial', 'B', 8)  # Multi-source headers: Font size 8
                    
                    # First row - main headers
                    pdf.cell(16, 8, 'Date', 1, 0, 'C')
                    pdf.cell(20, 8, 'Time', 1, 0, 'C')
                    pdf.cell(12, 8, 'TOD', 1, 0, 'C')
                    pdf.cell(18, 8, 'Consumed', 1, 0, 'C')
                    pdf.cell(18, 8, 'IEX After', 1, 0, 'C')
                    pdf.cell(16, 8, 'IEX', 1, 0, 'C')
                    pdf.cell(18, 8, 'CPP After', 1, 0, 'C')
                    pdf.cell(16, 8, 'CPP', 1, 0, 'C')
                    pdf.cell(18, 8, 'Total', 1, 0, 'C')
                    pdf.cell(12, 8, 'Missing', 1, 0, 'C')
                    pdf.ln()
                    
                    # Second row - specifications
                    pdf.cell(16, 8, '', 1, 0, 'C')
                    pdf.cell(20, 8, '', 1, 0, 'C')
                    pdf.cell(12, 8, '', 1, 0, 'C')
                    pdf.cell(18, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(18, 8, 'Loss (kWh)', 1, 0, 'C')
                    pdf.cell(16, 8, 'Excess', 1, 0, 'C')
                    pdf.cell(18, 8, 'Loss (kWh)', 1, 0, 'C')
                    pdf.cell(16, 8, 'Excess', 1, 0, 'C')
                    pdf.cell(18, 8, 'Excess', 1, 0, 'C')
                    pdf.cell(12, 8, 'Info', 1, 0, 'C')
                    pdf.ln()
                    
                    # Third row - units
                    pdf.cell(16, 8, '', 1, 0, 'C')
                    pdf.cell(20, 8, '', 1, 0, 'C')
                    pdf.cell(12, 8, '', 1, 0, 'C')
                    pdf.cell(18, 8, '', 1, 0, 'C')
                    pdf.cell(18, 8, '', 1, 0, 'C')
                    pdf.cell(16, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(18, 8, '', 1, 0, 'C')
                    pdf.cell(16, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(18, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(12, 8, '', 1, 0, 'C')
                    pdf.ln()
                    
                    # Reset font to table data font after headers
                    pdf.set_font('Arial', '', 9)  # Multi-source table content font size
                    
                else:
                    # Standard table for single source with properly formatted headers
                    pdf.set_font('Arial', 'B', 9)  # Single-source headers: Font size 9
                    
                    # First row - main headers with borders
                    pdf.cell(20, 8, 'Date', 1, 0, 'C')
                    pdf.cell(25, 8, 'Time', 1, 0, 'C')  
                    pdf.cell(15, 8, 'TOD', 1, 0, 'C')  
                    pdf.cell(25, 8, 'Generated', 1, 0, 'C')
                    pdf.cell(25, 8, 'Consumed', 1, 0, 'C') 
                    pdf.cell(25, 8, 'Excess', 1, 0, 'C')
                    pdf.cell(15, 8, 'Missing', 1, 0, 'C')
                    pdf.ln()
                    
                    # Second row for "After Loss", "Energy", "Energy", "Info"
                    pdf.cell(20, 8, '', 1, 0, 'C')  # Border for Date column
                    pdf.cell(25, 8, '', 1, 0, 'C')  # Border for Time column
                    pdf.cell(15, 8, '', 1, 0, 'C')  # Border for TOD column
                    pdf.cell(25, 8, 'After Loss', 1, 0, 'C')
                    pdf.cell(25, 8, 'Energy', 1, 0, 'C')
                    pdf.cell(25, 8, 'Energy', 1, 0, 'C')
                    pdf.cell(15, 8, 'Info', 1, 0, 'C')
                    pdf.ln()
                    
                    # Third row for units "(kWh)"
                    pdf.cell(20, 8, '', 1, 0, 'C')  # Border for Date column
                    pdf.cell(25, 8, '', 1, 0, 'C')  # Border for Time column
                    pdf.cell(15, 8, '', 1, 0, 'C')  # Border for TOD column
                    pdf.cell(25, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(25, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(25, 8, '(kWh)', 1, 0, 'C')
                    pdf.cell(15, 8, '', 1, 0, 'C')  # Border for Missing Info column
                    pdf.ln()
                
                # Reset font to table data font after headers
                pdf.set_font('Arial', '', 8)  # Single-source table content font size
            
            # Add initial table headers
            add_table_headers()
            
            # Set table content font size based on source type
            if generated_files and cpp_files:
                pdf.set_font('Arial', '', 9)  # Multi-source: table content font size 9
            else:
                pdf.set_font('Arial', '', 8)  # Single-source: table content font size 8
            table_complete = False  # Flag to track if table data is finished
            
            # Helper function for proper rounding (≥0.5 rounds up) to avoid scope issues
            def round_excess(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            for idx, row in pdf_data.iterrows():
                # Check if we need a new page (leaving space for summary)
                if pdf.get_y() > 250:  # Near bottom of page
                    pdf.add_page()
                    # Only add headers if we're still in the table data section
                    if not table_complete:
                        add_table_headers()  # Add headers on new page only for table data
                
                if generated_files and cpp_files:
                    # Sequential adjustment table data with adjusted column widths
                    pdf.cell(16, 7, safe_date_str(row['Slot_Date']), 1, 0, 'C')
                    pdf.cell(20, 7, format_time(row['Slot_Time']), 1, 0, 'C')  # Increased time column width
                    
                    tod_cat = row.get('TOD_Category', '')
                    pdf.cell(12, 7, tod_cat, 1, 0, 'C')
                    
                    pdf.cell(18, 7, f"{round_excess(row['Energy_kWh_cons'])}", 1, 0, 'C')
                    pdf.cell(18, 7, f"{round_excess(row.get('IEX_After_Loss', 0))}", 1, 0, 'C')
                    # Show rounded excess values instead of decimals
                    iex_excess_rounded = round_excess(row.get('IEX_Excess', 0))
                    pdf.cell(16, 7, f"{iex_excess_rounded}", 1, 0, 'C')
                    pdf.cell(18, 7, f"{round_excess(row.get('CPP_After_Loss', 0))}", 1, 0, 'C')
                    cpp_excess_rounded = round_excess(row.get('CPP_Excess', 0))
                    pdf.cell(16, 7, f"{cpp_excess_rounded}", 1, 0, 'C')
                    total_excess_rounded = round_excess(row['Total_Excess'])
                    pdf.cell(18, 7, f"{total_excess_rounded}", 1, 0, 'C')
                    pdf.cell(12, 7, row.get('Missing_Info', '')[:3], 1, 0, 'C')  # Truncate missing info
                else:
                    # Standard table data for single source
                    pdf.cell(20, 7, safe_date_str(row['Slot_Date']), 1, 0, 'C')
                    pdf.cell(25, 7, format_time(row['Slot_Time']), 1, 0, 'C')  # Increased time column width
                    
                    tod_cat = row.get('TOD_Category', '')
                    pdf.cell(15, 7, tod_cat, 1, 0, 'C')
                    
                    pdf.cell(25, 7, f"{row['After_Loss']:.2f}", 1, 0, 'C')
                    pdf.cell(25, 7, f"{row['Energy_kWh_cons']:.2f}", 1, 0, 'C')
                    # Show rounded excess values instead of decimals
                    total_excess_rounded = round_excess(row['Total_Excess'])
                    pdf.cell(25, 7, f"{total_excess_rounded}", 1, 0, 'C')
                    pdf.cell(15, 7, row.get('Missing_Info', '')[:4], 1, 0, 'C')
                
                pdf.ln()
            
            # Mark table as complete - no more headers needed for subsequent pages
            table_complete = True
            
            pdf.ln(2)
            # Show error/warning message in PDF if present
            if 'error_message' in globals() and error_message:
                pdf.set_font('Arial', 'B', 10)
                pdf.multi_cell(0, 8, f"Warnings/Errors:\n{error_message}")

            # Enhanced summary for sequential adjustment
            # Check if we need a new page for summary (but don't add table headers)
            if pdf.get_y() > 220:  # Need more space for summary
                pdf.add_page()
            
            pdf.ln(2)
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'DETAILED CALCULATION SUMMARY:', ln=True)
            pdf.set_font('Arial', '', 11)
            
            # Helper function for proper rounding (≥0.5 rounds up) to match table values
            def round_kwh_summary(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            if generated_files and cpp_files:
                # Sequential adjustment summary - use rounded totals from table data for precision
                if full_totals:
                    # Use full totals from all slots but round them to match table display
                    total_iex_before_loss_raw = full_totals.get('iex_before', 0)
                    total_cpp_before_loss_raw = full_totals.get('cpp_before', 0) 
                    total_iex_after_loss_raw = full_totals.get('iex_after', 0)
                    total_cpp_after_loss_raw = full_totals.get('cpp_after', 0)
                    total_iex_excess_raw = full_totals.get('iex_excess', 0)
                    total_cpp_excess_raw = full_totals.get('cpp_excess', 0)
                else:
                    # Calculate from pdf_data but round to match table display
                    total_iex_before_loss_raw = pdf_data['IEX_Energy_kWh'].sum() if 'IEX_Energy_kWh' in pdf_data.columns else 0
                    total_cpp_before_loss_raw = pdf_data['CPP_Energy_kWh'].sum() if 'CPP_Energy_kWh' in pdf_data.columns else 0
                    total_iex_after_loss_raw = pdf_data['IEX_After_Loss'].sum() if 'IEX_After_Loss' in pdf_data.columns else 0
                    total_cpp_after_loss_raw = pdf_data['CPP_After_Loss'].sum() if 'CPP_After_Loss' in pdf_data.columns else 0
                    total_iex_excess_raw = pdf_data['IEX_Excess'].sum() if 'IEX_Excess' in pdf_data.columns else 0
                    total_cpp_excess_raw = pdf_data['CPP_Excess'].sum() if 'CPP_Excess' in pdf_data.columns else 0
                
                # Round all values to match table display (this is what users see in the detailed table)
                total_iex_before_loss_rounded = round_kwh_summary(total_iex_before_loss_raw)
                total_iex_after_loss_rounded = round_kwh_summary(total_iex_after_loss_raw)
                total_cpp_before_loss_rounded = round_kwh_summary(total_cpp_before_loss_raw)
                total_cpp_after_loss_rounded = round_kwh_summary(total_cpp_after_loss_raw)
                total_iex_excess_rounded = round_kwh_summary(total_iex_excess_raw)
                total_cpp_excess_rounded = round_kwh_summary(total_cpp_excess_raw)
                total_excess_rounded = total_iex_excess_rounded + total_cpp_excess_rounded
                
                iex_adjustment_rounded = round_kwh_summary(total_iex_after_loss_raw - total_iex_excess_raw)
                cpp_adjustment_rounded = round_kwh_summary(total_cpp_after_loss_raw - total_cpp_excess_raw)
                
                pdf.cell(0, 8, f'I.E.X Generation (before T&D loss): {total_iex_before_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Generation (after {t_and_d_loss}% T&D loss): {total_iex_after_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Adjustment with Consumption: {iex_adjustment_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Excess Energy (rounded): {total_iex_excess_rounded} kWh', ln=True)
                pdf.ln(3)
                
                pdf.cell(0, 8, f'C.P.P Generation (before T&D loss): {total_cpp_before_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Generation (after {cpp_t_and_d_loss}% T&D loss): {total_cpp_after_loss_rounded} kWh', ln=True)
                remaining_consumption_total_raw = pdf_data['Remaining_Consumption'].sum() if 'Remaining_Consumption' in pdf_data.columns else 0
                remaining_consumption_total_rounded = round_kwh_summary(remaining_consumption_total_raw)
                pdf.cell(0, 8, f'Remaining Consumption (after I.E.X adjustment): {remaining_consumption_total_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Adjustment with Remaining Consumption: {cpp_adjustment_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Excess Energy (rounded): {total_cpp_excess_rounded} kWh', ln=True)
                pdf.ln(3)
                
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, f'TOTAL CALCULATIONS:', ln=True)
                pdf.set_font('Arial', '', 11)
                total_generation_before_rounded = round_kwh_summary(total_iex_before_loss_raw + total_cpp_before_loss_raw)
                total_generation_after_rounded = round_kwh_summary(total_iex_after_loss_raw + total_cpp_after_loss_raw)
                total_consumed_rounded = round_kwh_summary(total_consumed)
                comparison_rounded = round_kwh_summary((total_iex_before_loss_raw + total_cpp_before_loss_raw) - (total_iex_after_loss_raw + total_cpp_after_loss_raw))
                
                pdf.cell(0, 8, f'Total Generation (before loss): {total_generation_before_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Generation (after loss): {total_generation_after_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Consumed Energy: {total_consumed_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Comparison (Total Generation before loss - after loss): {comparison_rounded} kWh', ln=True)
            else:
                # Standard summary for single source - use rounded totals to match table
                total_excess_raw = total_excess
                total_excess_rounded = round_kwh_summary(total_excess_raw)
                
                if enable_iex:
                    if full_totals:
                        total_iex_before_loss_raw = full_totals.get('iex_before', 0)
                        total_iex_after_loss_raw = full_totals.get('iex_after', 0)
                    else:
                        total_iex_before_loss_raw = pdf_data['IEX_Energy_kWh'].sum() if 'IEX_Energy_kWh' in pdf_data.columns else 0
                        total_iex_after_loss_raw = pdf_data['IEX_After_Loss'].sum() if 'IEX_After_Loss' in pdf_data.columns else 0
                    
                    total_iex_before_loss_rounded = round_kwh_summary(total_iex_before_loss_raw)
                    total_iex_after_loss_rounded = round_kwh_summary(total_iex_after_loss_raw)
                    pdf.cell(0, 8, f'I.E.X Generation (before T&D loss): {total_iex_before_loss_rounded} kWh', ln=True)
                    pdf.cell(0, 8, f'I.E.X Generation (after {t_and_d_loss}% T&D loss): {total_iex_after_loss_rounded} kWh', ln=True)
                
                if enable_cpp:
                    if full_totals:
                        total_cpp_before_loss_raw = full_totals.get('cpp_before', 0)
                        total_cpp_after_loss_raw = full_totals.get('cpp_after', 0)
                    else:
                        total_cpp_before_loss_raw = pdf_data['CPP_Energy_kWh'].sum() if 'CPP_Energy_kWh' in pdf_data.columns else 0
                        total_cpp_after_loss_raw = pdf_data['CPP_After_Loss'].sum() if 'CPP_After_Loss' in pdf_data.columns else 0
                    
                    total_cpp_before_loss_rounded = round_kwh_summary(total_cpp_before_loss_raw)
                    total_cpp_after_loss_rounded = round_kwh_summary(total_cpp_after_loss_raw)
                    pdf.cell(0, 8, f'C.P.P Generation (before T&D loss): {total_cpp_before_loss_rounded} kWh', ln=True)
                    pdf.cell(0, 8, f'C.P.P Generation (after {cpp_t_and_d_loss}% T&D loss): {total_cpp_after_loss_rounded} kWh', ln=True)
                
                total_consumed_rounded = round_kwh_summary(total_consumed)
                pdf.cell(0, 8, f'Total Consumed Energy (after multiplication): {total_consumed_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
            
            pdf.cell(0, 8, f'Unique Days Used (Generated): {unique_days_gen_full}', ln=True)
            pdf.cell(0, 8, f'Unique Days Used (Consumed): {unique_days_cons_full}', ln=True)
            pdf.cell(0, 8, f'Status: {excess_status}', ln=True)
            
            # Add TOD-wise excess energy breakdown
            # Check if we need a new page for TOD breakdown (but don't add table headers)
            if pdf.get_y() > 220:  # Need space for TOD breakdown
                pdf.add_page()
            
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 14)  # Standardized heading font size
            pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
            pdf.set_font('Arial', 'B', 10)  # Consistent with table data font size
            pdf.cell(20, 10, 'TOD', 1)
            pdf.cell(50, 10, 'Excess Energy (kWh)', 1)
            
            pdf.ln()
            pdf.set_font('Arial', '', 10)  # Consistent with table data font size
            
            # Get TOD-wise excess from the dataframe using rounded values to match table display
            tod_excess_raw = pdf_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
            
            # Apply rounding to match table values (what users see in the detailed table)
            def round_excess_breakdown(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            tod_descriptions = {
                'C1': 'Morning Peak',
                'C2': 'Evening Peak',
                'C4': 'Normal Hours',
                'C5': 'Night Hours',
                'Unknown': 'Unknown Time Slot'
            }
            
            # Calculate C category (sum of C1, C2, C4, C5) using rounded values
            c_categories = ['C1', 'C2', 'C4', 'C5']
            c_total_rounded = 0
            
            for _, row in tod_excess_raw.iterrows():
                category = row['TOD_Category']
                excess_raw = row['Total_Excess']
                excess_rounded = round_excess_breakdown(excess_raw)
                
                pdf.cell(20, 10, category, 1)
                pdf.cell(50, 10, f"{excess_rounded}", 1)  # Show rounded values to match table
                pdf.ln()
                
                # Add to C category total if applicable
                if category in c_categories:
                    c_total_rounded += excess_rounded
            
            # Add C category total (sum of rounded individual values)
            pdf.cell(20, 10, 'C', 1)
            pdf.cell(50, 10, f"{c_total_rounded}", 1)
            pdf.ln()
            
            # Add financial calculations on a dedicated page
            pdf.add_page()
            pdf.set_font('Arial', 'B', 14)  # Standardized heading font size
            pdf.cell(0, 10, 'Financial Calculations:', ln=True)
            pdf.set_font('Arial', '', 10)  # Consistent with table data font size
            
            # Helper function for proper rounding (≥0.5 rounds up)
            def round_kwh(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            # Get IEX excess for cross subsidy surcharge calculation
            if 'IEX_Excess' in pdf_data.columns:
                iex_excess_total_raw = pdf_data['IEX_Excess'].sum()
            else:
                iex_excess_total_raw = 0
            
            # Use rounded values for financial calculations to match table display
            if generated_files and cpp_files:
                # For multi-source, use the rounded total excess calculated earlier
                total_excess_rounded_fin = total_excess_rounded  # Already calculated above
            else:
                # For single source, calculate rounded total excess
                total_excess_rounded_fin = round_kwh(total_excess)
            
            iex_excess_rounded = round_kwh(iex_excess_total_raw)
            
            # Base rate calculation using rounded total excess
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess_rounded_fin * base_rate
            pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess_rounded_fin} kWh) x Rs.7.25 = Rs.{base_amount:.2f}", ln=True)
            
            # Additional charges for specific TOD categories using rounded values from breakdown
            # Calculate rounded C1+C2 and C5 totals from the rounded TOD breakdown
            c1_c2_excess_rounded = 0
            c5_excess_rounded = 0
            
            for _, row in tod_excess_raw.iterrows():
                category = row['TOD_Category']
                excess_raw = row['Total_Excess']
                excess_rounded = round_excess_breakdown(excess_raw)
                
                if category in ['C1', 'C2']:
                    c1_c2_excess_rounded += excess_rounded
                elif category == 'C5':
                    c5_excess_rounded += excess_rounded
            
            c1_c2_additional = c1_c2_excess_rounded * 1.8125  # rupees per kWh
            pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess_rounded} kWh) x Rs.1.8125 = Rs.{c1_c2_additional:.2f}", ln=True)
            
            c5_additional = c5_excess_rounded * 0.3625  # rupees per kWh
            pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess_rounded} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
            
            # Calculate total amount
            total_amount = base_amount + c1_c2_additional + c5_additional
            pdf.cell(0, 8, f"4. Total Amount: Rs.{base_amount:.2f} + Rs.{c1_c2_additional:.2f} + Rs.{c5_additional:.2f} = Rs.{total_amount:.2f}", ln=True)
            
            # Calculate E-Tax (5% of total amount)
            etax = total_amount * 0.05
            pdf.cell(0, 8, f"5. E-Tax (5% of Total Amount): Rs.{total_amount:.2f} x 0.05 = Rs.{etax:.2f}", ln=True)
            
            # Calculate total amount with E-Tax
            total_with_etax = total_amount + etax
            pdf.cell(0, 8, f"6. Total Amount with E-Tax: Rs.{total_amount:.2f} + Rs.{etax:.2f} = Rs.{total_with_etax:.2f}", ln=True)
            
            # Calculate negative factors using rounded values for consistency
            etax_on_iex = total_excess_rounded * 0.1  # Use rounded total from summary
            pdf.cell(0, 8, f"7. E-Tax on IEX: Total Excess ({total_excess_rounded} kWh) x Rs.0.1 = Rs.{etax_on_iex:.2f}", ln=True)
            
            cross_subsidy_surcharge = iex_excess_rounded * 1.92
            pdf.cell(0, 8, f"8. Cross Subsidy Surcharge: IEX Excess ({iex_excess_rounded} kWh) x Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
            wheeling_reference_kwh, wheeling_charges = compute_wheeling_components(
                total_excess_rounded_fin,
                t_and_d_loss,
            )

            pdf.cell(0, 8, f"9. Wheeling Charges: Adj. Loss Component ({wheeling_reference_kwh:.2f} kWh) x Rs.{WHEELING_RATE_PER_KWH:.2f} = Rs.{wheeling_charges:.2f}", ln=True)

            # Calculate final amount to be collected with detailed breakdown
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)

            # Break down the calculation like wheeling charges
            pdf.cell(0, 8, f"10a. Total Amount to be Collected - Step 1:", ln=True)
            pdf.cell(0, 8, f"     Rs.{total_with_etax:.2f} - (Rs.{etax_on_iex:.2f} + Rs.{cross_subsidy_surcharge:.2f} + Rs.{wheeling_charges:.2f})", ln=True)
            pdf.cell(0, 8, f"10b. Total Amount to be Collected - Step 2:", ln=True)
            pdf.cell(0, 8, f"     Rs.{total_with_etax:.2f} - Rs.{etax_on_iex + cross_subsidy_surcharge + wheeling_charges:.2f} = Rs.{final_amount:.2f}", ln=True)

            # Round up final amount to next highest value
            final_amount_rounded = math.ceil(final_amount)

            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 10, f"11. Final Amount (Rounded Up): Rs.{final_amount_rounded}", ln=True)

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
            cpp_count = len(cpp_files) if cpp_files else 0
            iex_count = len(generated_files) if generated_files else 0
            auto_detect_info = f"Auto-detected from {iex_count} I.E.X, {cpp_count} C.P.P, and {len(consumed_files)} consumed files"
            
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
        def generate_daywise_pdf(pdf_data, month, year, filename, enable_iex=enable_iex, enable_cpp=enable_cpp, t_and_d_loss=t_and_d_loss, cpp_t_and_d_loss=cpp_t_and_d_loss, auto_detect_info=auto_detect_info):
            # This function generates a PDF with only the day-wise summary table (all days in month, even if missing)
            # Import datetime for timestamp
            from datetime import datetime, timedelta
            import pandas as pd
            
            pdf = AuthorPDF()
            pdf.set_margins(20, 20, 20)  # Set proper margins: left, top, right (20mm each)
            pdf.set_auto_page_break(auto=True, margin=20)  # Auto page break with bottom margin
            pdf.add_page()
            
            # FIRST PAGE - DESCRIPTION AND INFORMATION ONLY
            pdf.set_font('Arial', 'B', 16)  # Larger title font
            pdf.cell(0, 15, 'Energy Adjustment Day-wise Summary Report', ln=True, align='C')
            pdf.ln(10)
            
            # Consumer Information Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Consumer Information:', ln=True)
            pdf.set_font('Arial', '', 12)
            pdf.cell(0, 8, f'Consumer Number: {consumer_number}', ln=True)
            pdf.cell(0, 8, f'Consumer Name: {consumer_name}', ln=True)
            pdf.cell(0, 8, f'Multiplication Factor (Consumed Energy): {multiplication_factor}', ln=True)
            pdf.ln(5)
            
            # Technical Parameters Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Technical Parameters:', ln=True)
            pdf.set_font('Arial', '', 12)
            
            # Display T&D losses based on what sources are used
            if enable_iex and enable_cpp:
                pdf.cell(0, 8, f'I.E.X T&D Loss (%): {t_and_d_loss}', ln=True)
                pdf.cell(0, 8, f'C.P.P T&D Loss (%): {cpp_t_and_d_loss}', ln=True)
            elif enable_iex:
                pdf.cell(0, 8, f'I.E.X T&D Loss (%): {t_and_d_loss}', ln=True)
            elif enable_cpp:
                pdf.cell(0, 8, f'C.P.P T&D Loss (%): {cpp_t_and_d_loss}', ln=True)
            pdf.ln(5)
            
            # Report Information Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Report Information:', ln=True)
            pdf.set_font('Arial', '', 12)
            if month and year:
                pdf.cell(0, 8, f'Report Period: {month}/{year}', ln=True)
            if auto_detect_info:
                pdf.cell(0, 8, f'Period Details: {auto_detect_info}', ln=True)
            pdf.cell(0, 8, f'Report Generated: {datetime.now().strftime("%d/%m/%Y at %H:%M:%S")}', ln=True)
            pdf.ln(10)
            
            # Report Description Section
            pdf.set_font('Arial', 'B', 14)
            pdf.cell(0, 10, 'Report Description:', ln=True)
            pdf.set_font('Arial', '', 11)
            
            description_text = [
                "This day-wise summary report consolidates energy adjustment calculations",
                "into daily totals, showing the overall energy balance for each day",
                "of the selected month. The report includes all days in the month,",
                "even those with no energy transactions, for complete visibility."
            ]
            
            for line in description_text:
                pdf.cell(0, 6, line, ln=True)
            
            pdf.ln(10)
            
            # Add page break before tables
            pdf.add_page()
            
            # SECOND PAGE AND BEYOND - TABLES START HERE
            pdf.set_font('Arial', 'B', 16)
            pdf.cell(0, 10, 'Day-wise Summary Table', ln=True, align='C')
            pdf.ln(10)
            
            # Set up day-wise table headers
            pdf.set_font('Arial', 'B', 9)  # Consistent with single-source table headers
            pdf.cell(40, 10, 'Date', 1)
            pdf.cell(50, 10, 'Total Gen. After Loss', 1)
            pdf.cell(50, 10, 'Total Consumed', 1)
            pdf.cell(50, 10, 'Total Excess', 1)
            pdf.ln()
            pdf.set_font('Arial', '', 8)  # Reduced font size for day-wise table content
            # Determine full date range for the selected month
            if month and year:
                try:
                    month_int = int(float(month))
                    year_int = int(float(year))
                except ValueError:
                    return render_template('index.html', error=f"Invalid month or year value in daywise PDF generation. Month: '{month}', Year: '{year}'")
                
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
                'Total_Excess': 'sum'
            })
            daywise = daywise.reindex(all_days.strftime('%d/%m/%Y'), fill_value=0).reset_index()
            daywise = daywise.rename(columns={'index': 'Slot_Date'})
            for idx, row in daywise.iterrows():
                pdf.cell(40, 10, row['Slot_Date'], 1)
                pdf.cell(50, 10, f"{row['After_Loss']:.4f}", 1)
                pdf.cell(50, 10, f"{row['Energy_kWh_cons']:.4f}", 1)
                # Round excess values for display using proper rounding (≥0.5 rounds up)
                total_excess_rounded = int(row['Total_Excess'] + 0.5) if row['Total_Excess'] >= 0 else int(row['Total_Excess'] - 0.5)
                pdf.cell(50, 10, f"{total_excess_rounded}", 1)
                pdf.ln()
            pdf.ln(2)
            
            # Skip detailed slot-wise data for day-wise PDF - go directly to summaries
            
            # Add missing DETAILED CALCULATION SUMMARY heading
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(0, 10, 'DETAILED CALCULATION SUMMARY:', ln=True)
            pdf.set_font('Arial', '', 11)  # Standardized font size to match regular PDF
            
            # Add calculation summary at the end using rounded values to match table display
            sum_injection = pdf_data['Energy_kWh_gen'].sum() if 'Energy_kWh_gen' in pdf_data.columns else 0
            total_generated_after_loss = pdf_data['After_Loss'].sum() if 'After_Loss' in pdf_data.columns else 0
            comparison = sum_injection - total_generated_after_loss
            total_consumed = pdf_data['Energy_kWh_cons'].sum() if 'Energy_kWh_cons' in pdf_data.columns else 0
            total_excess = pdf_data['Total_Excess'].sum() if 'Total_Excess' in pdf_data.columns else 0
            
            # Helper function for proper rounding to match table values
            def round_kwh_daywise_summary(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            # Determine if this is multi-source (IEX + CPP) or single source
            is_multi_source = generated_files and cpp_files
            
            if is_multi_source:
                # Sequential adjustment summary - use rounded totals from table data for precision
                if full_totals:
                    # Use full totals from all slots but round them to match table display
                    total_iex_before_loss_raw = full_totals.get('iex_before', 0)
                    total_cpp_before_loss_raw = full_totals.get('cpp_before', 0) 
                    total_iex_after_loss_raw = full_totals.get('iex_after', 0)
                    total_cpp_after_loss_raw = full_totals.get('cpp_after', 0)
                    total_iex_excess_raw = full_totals.get('iex_excess', 0)
                    total_cpp_excess_raw = full_totals.get('cpp_excess', 0)
                else:
                    # Calculate from pdf_data but round to match table display
                    total_iex_before_loss_raw = pdf_data['IEX_Energy_kWh'].sum() if 'IEX_Energy_kWh' in pdf_data.columns else 0
                    total_cpp_before_loss_raw = pdf_data['CPP_Energy_kWh'].sum() if 'CPP_Energy_kWh' in pdf_data.columns else 0
                    total_iex_after_loss_raw = pdf_data['IEX_After_Loss'].sum() if 'IEX_After_Loss' in pdf_data.columns else 0
                    total_cpp_after_loss_raw = pdf_data['CPP_After_Loss'].sum() if 'CPP_After_Loss' in pdf_data.columns else 0
                    total_iex_excess_raw = pdf_data['IEX_Excess'].sum() if 'IEX_Excess' in pdf_data.columns else 0
                    total_cpp_excess_raw = pdf_data['CPP_Excess'].sum() if 'CPP_Excess' in pdf_data.columns else 0
                
                # Round all values to match table display (this is what users see in the detailed table)
                total_iex_before_loss_rounded = round_kwh_daywise_summary(total_iex_before_loss_raw)
                total_iex_after_loss_rounded = round_kwh_daywise_summary(total_iex_after_loss_raw)
                total_cpp_before_loss_rounded = round_kwh_daywise_summary(total_cpp_before_loss_raw)
                total_cpp_after_loss_rounded = round_kwh_daywise_summary(total_cpp_after_loss_raw)
                total_iex_excess_rounded = round_kwh_daywise_summary(total_iex_excess_raw)
                total_cpp_excess_rounded = round_kwh_daywise_summary(total_cpp_excess_raw)
                total_excess_rounded = total_iex_excess_rounded + total_cpp_excess_rounded
                
                iex_adjustment_rounded = round_kwh_daywise_summary(total_iex_after_loss_raw - total_iex_excess_raw)
                cpp_adjustment_rounded = round_kwh_daywise_summary(total_cpp_after_loss_raw - total_cpp_excess_raw)
                
                pdf.cell(0, 8, f'I.E.X Generation (before T&D loss): {total_iex_before_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Generation (after {t_and_d_loss}% T&D loss): {total_iex_after_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Adjustment with Consumption: {iex_adjustment_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'I.E.X Excess Energy (rounded): {total_iex_excess_rounded} kWh', ln=True)
                pdf.ln(3)
                
                pdf.cell(0, 8, f'C.P.P Generation (before T&D loss): {total_cpp_before_loss_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Generation (after {cpp_t_and_d_loss}% T&D loss): {total_cpp_after_loss_rounded} kWh', ln=True)
                remaining_consumption_total_raw = pdf_data['Remaining_Consumption'].sum() if 'Remaining_Consumption' in pdf_data.columns else 0
                remaining_consumption_total_rounded = round_kwh_daywise_summary(remaining_consumption_total_raw)
                pdf.cell(0, 8, f'Remaining Consumption (after I.E.X adjustment): {remaining_consumption_total_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Adjustment with Remaining Consumption: {cpp_adjustment_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'C.P.P Excess Energy (rounded): {total_cpp_excess_rounded} kWh', ln=True)
                pdf.ln(3)
                
                pdf.set_font('Arial', 'B', 11)
                pdf.cell(0, 8, f'TOTAL CALCULATIONS:', ln=True)
                pdf.set_font('Arial', '', 11)
                total_generation_before_rounded = round_kwh_daywise_summary(total_iex_before_loss_raw + total_cpp_before_loss_raw)
                total_generation_after_rounded = round_kwh_daywise_summary(total_iex_after_loss_raw + total_cpp_after_loss_raw)
                total_consumed_rounded = round_kwh_daywise_summary(total_consumed)
                comparison_rounded = round_kwh_daywise_summary((total_iex_before_loss_raw + total_cpp_before_loss_raw) - (total_iex_after_loss_raw + total_cpp_after_loss_raw))
                
                pdf.cell(0, 8, f'Total Generation (before loss): {total_generation_before_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Generation (after loss): {total_generation_after_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Consumed Energy: {total_consumed_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Comparison (Total Generation before loss - after loss): {comparison_rounded} kWh', ln=True)
            else:
                # Standard summary for single source - use rounded totals to match table
                total_excess_raw = total_excess
                total_excess_rounded = round_kwh_daywise_summary(total_excess_raw)
                
                if enable_iex:
                    if full_totals:
                        total_iex_before_loss_raw = full_totals.get('iex_before', 0)
                        total_iex_after_loss_raw = full_totals.get('iex_after', 0)
                    else:
                        total_iex_before_loss_raw = pdf_data['IEX_Energy_kWh'].sum() if 'IEX_Energy_kWh' in pdf_data.columns else 0
                        total_iex_after_loss_raw = pdf_data['IEX_After_Loss'].sum() if 'IEX_After_Loss' in pdf_data.columns else 0
                    
                    total_iex_before_loss_rounded = round_kwh_daywise_summary(total_iex_before_loss_raw)
                    total_iex_after_loss_rounded = round_kwh_daywise_summary(total_iex_after_loss_raw)
                    pdf.cell(0, 8, f'I.E.X Generation (before T&D loss): {total_iex_before_loss_rounded} kWh', ln=True)
                    pdf.cell(0, 8, f'I.E.X Generation (after {t_and_d_loss}% T&D loss): {total_iex_after_loss_rounded} kWh', ln=True)
                
                if enable_cpp:
                    if full_totals:
                        total_cpp_before_loss_raw = full_totals.get('cpp_before', 0)
                        total_cpp_after_loss_raw = full_totals.get('cpp_after', 0)
                    else:
                        total_cpp_before_loss_raw = pdf_data['CPP_Energy_kWh'].sum() if 'CPP_Energy_kWh' in pdf_data.columns else 0
                        total_cpp_after_loss_raw = pdf_data['CPP_After_Loss'].sum() if 'CPP_After_Loss' in pdf_data.columns else 0
                    
                    total_cpp_before_loss_rounded = round_kwh_daywise_summary(total_cpp_before_loss_raw)
                    total_cpp_after_loss_rounded = round_kwh_daywise_summary(total_cpp_after_loss_raw)
                    pdf.cell(0, 8, f'C.P.P Generation (before T&D loss): {total_cpp_before_loss_rounded} kWh', ln=True)
                    pdf.cell(0, 8, f'C.P.P Generation (after {cpp_t_and_d_loss}% T&D loss): {total_cpp_after_loss_rounded} kWh', ln=True)
                
                total_consumed_rounded = round_kwh_daywise_summary(total_consumed)
                pdf.cell(0, 8, f'Total Consumed Energy (after multiplication): {total_consumed_rounded} kWh', ln=True)
                pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
            
            pdf.cell(0, 8, f'Unique Days Used (Generated): {unique_days_gen_full}', ln=True)
            pdf.cell(0, 8, f'Unique Days Used (Consumed): {unique_days_cons_full}', ln=True)
            pdf.cell(0, 8, f'Status: {excess_status}', ln=True)
            
            # Check if we need a new page for TOD breakdown 
            if pdf.get_y() > 220:  # Need space for TOD breakdown
                pdf.add_page()
            
            # Add TOD-wise excess energy breakdown
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 14)  # Standardized heading font size
            pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
            pdf.set_font('Arial', 'B', 10)  # Consistent with table data font size
            pdf.cell(20, 10, 'TOD', 1)
            pdf.cell(50, 10, 'Excess Energy (kWh)', 1)
            
            pdf.ln()
            pdf.set_font('Arial', '', 10)  # Consistent with table data font size
            
            # Get TOD-wise excess from the dataframe using rounded values to match table display
            tod_excess = pdf_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
            
            # Apply rounding to match table values (what users see in the detailed table)
            def round_excess_daywise(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            # Define TOD category order with C at the top as requested
            tod_order = ['C', 'C1', 'C2', 'C4', 'C5', 'Unknown']
            tod_descriptions = {
                'C1': 'Morning Peak',
                'C2': 'Evening Peak',
                'C4': 'Normal Hours',
                'C5': 'Night Hours',
                'Unknown': 'Unknown Time Slot'
            }
            
            # Calculate C category (sum of C1, C2, C4, C5) using rounded values
            c_categories = ['C1', 'C2', 'C4', 'C5']
            c_total_rounded_daywise = 0
            
            # Store individual category values for ordered display
            tod_values = {}
            for _, row in tod_excess.iterrows():
                category = row['TOD_Category']
                excess_raw = row['Total_Excess']
                excess_rounded = round_excess_daywise(excess_raw)
                tod_values[category] = excess_rounded
                
                # Add to C category total if applicable
                if category in c_categories:
                    c_total_rounded_daywise += excess_rounded
            
            # Display TOD breakdown in proper order with C at the top
            pdf.cell(20, 10, 'C', 1)
            pdf.cell(50, 10, f"{c_total_rounded_daywise}", 1)
            pdf.ln()
            
            # Display individual categories
            for category in ['C1', 'C2', 'C4', 'C5']:
                if category in tod_values:
                    pdf.cell(20, 10, category, 1)
                    pdf.cell(50, 10, f"{tod_values[category]}", 1)
                    pdf.ln()
            
            # Display any Unknown categories
            if 'Unknown' in tod_values:
                pdf.cell(20, 10, 'Unknown', 1)
                pdf.cell(50, 10, f"{tod_values['Unknown']}", 1)
                pdf.ln()
            
            # Add financial calculations on a dedicated page
            pdf.add_page()
            pdf.set_font('Arial', 'B', 14)  # Standardized heading font size
            pdf.cell(0, 10, 'Financial Calculations:', ln=True)
            pdf.set_font('Arial', '', 10)  # Consistent with table data font size
            
            # Helper function for proper rounding (≥0.5 rounds up)
            def round_kwh_daywise(value):
                return int(value + 0.5) if value >= 0 else int(value - 0.5)
            
            # Get IEX excess for cross subsidy surcharge calculation
            if 'IEX_Excess' in pdf_data.columns:
                iex_excess_total_raw = pdf_data['IEX_Excess'].sum()
            else:
                iex_excess_total_raw = 0
            
            # Use rounded values for financial calculations to match table display
            total_excess_rounded_daywise = round_kwh_daywise(total_excess)
            iex_excess_rounded = round_kwh_daywise(iex_excess_total_raw)
            
            # Base rate calculation using rounded total excess
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess_rounded_daywise * base_rate
            pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess_rounded_daywise} kWh) x Rs.7.25 = Rs.{base_amount:.2f}", ln=True)
            
            # Additional charges for specific TOD categories using rounded values from breakdown
            # Calculate rounded C1+C2 and C5 totals from the rounded TOD breakdown
            c1_c2_excess_rounded_daywise = tod_values.get('C1', 0) + tod_values.get('C2', 0)
            c5_excess_rounded_daywise = tod_values.get('C5', 0)
            
            c1_c2_additional = c1_c2_excess_rounded_daywise * 1.8125  # rupees per kWh
            pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess_rounded_daywise} kWh) x Rs.1.8125 = Rs.{c1_c2_additional:.2f}", ln=True)
            
            c5_additional = c5_excess_rounded_daywise * 0.3625  # rupees per kWh
            pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess_rounded_daywise} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
            
            # Calculate partial amount for base rates and additional charges
            partial_amount = base_amount + c1_c2_additional + c5_additional
            pdf.cell(0, 8, f"4. Partial Total: Rs.{base_amount:.2f} + Rs.{c1_c2_additional:.2f} + Rs.{c5_additional:.2f} = Rs.{partial_amount:.2f}", ln=True)
            
            # Calculate E-Tax (5% of partial amount)
            etax = partial_amount * 0.05
            pdf.cell(0, 8, f"5. E-Tax (5% of Partial Total): Rs.{partial_amount:.2f} x 0.05 = Rs.{etax:.2f}", ln=True)
            
            # Calculate subtotal with E-Tax
            subtotal_with_etax = partial_amount + etax
            pdf.cell(0, 8, f"6. Subtotal with E-Tax: Rs.{partial_amount:.2f} + Rs.{etax:.2f} = Rs.{subtotal_with_etax:.2f}", ln=True)
            
            # Calculate negative factors (deductions)
            etax_on_iex = total_excess_rounded_daywise * 0.1  # Use rounded value for consistency
            pdf.cell(0, 8, f"7. E-Tax on IEX (Deduction): Total Excess ({total_excess_rounded_daywise} kWh) x Rs.0.1 = Rs.{etax_on_iex:.2f}", ln=True)
            
            cross_subsidy_surcharge = iex_excess_rounded * 1.92
            pdf.cell(0, 8, f"8. Cross Subsidy Surcharge (Deduction): IEX Excess ({iex_excess_rounded} kWh) x Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
            wheeling_reference_kwh, wheeling_charges = compute_wheeling_components(
                total_excess_rounded_daywise,
                t_and_d_loss,
            )

            pdf.cell(0, 8, f"9. Wheeling Charges: Adj. Loss Component ({wheeling_reference_kwh:.2f} kWh) x Rs.{WHEELING_RATE_PER_KWH:.2f} = Rs.{wheeling_charges:.2f}", ln=True)

            # Calculate final amount to be collected with detailed breakdown
            final_amount = subtotal_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)

            # Break down the calculation like in regular PDF
            pdf.cell(0, 8, f"10a. Total Amount to be Collected - Step 1:", ln=True)
            pdf.cell(0, 8, f"     Rs.{subtotal_with_etax:.2f} - (Rs.{etax_on_iex:.2f} + Rs.{cross_subsidy_surcharge:.2f} + Rs.{wheeling_charges:.2f})", ln=True)
            pdf.cell(0, 8, f"10b. Total Amount to be Collected - Step 2:", ln=True)
            pdf.cell(0, 8, f"     Rs.{subtotal_with_etax:.2f} - Rs.{etax_on_iex + cross_subsidy_surcharge + wheeling_charges:.2f} = Rs.{final_amount:.2f}", ln=True)

            # Round up final amount to next highest value
            final_amount_rounded = math.ceil(final_amount)

            pdf.set_font('Arial', 'B', 10)  # Consistent with table data font size
            pdf.cell(0, 8, f"11. Final Amount (Rounded Up): Rs.{final_amount_rounded}", ln=True)

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

        # Generate custom filename based on service number and name
        def generate_custom_filename(base_name, consumer_number, consumer_name, month=None, year=None):
            # Get last 3 digits of service number
            last_3_digits = str(consumer_number)[-3:] if len(str(consumer_number)) >= 3 else str(consumer_number)
            
            # Clean consumer name for filename (remove special characters)
            clean_name = "".join(c for c in consumer_name if c.isalnum() or c in (' ', '-', '_')).strip()
            clean_name = clean_name.replace(' ', '_')
            
            # Format month and year for filename (short date format)
            date_suffix = ""
            if month and year:
                try:
                    month_num = int(float(month))
                    year_num = int(float(year))
                    # Format as MM_YY (e.g., 07_25 for July 2025)
                    date_suffix = f"_{month_num:02d}_{year_num % 100:02d}"
                except:
                    date_suffix = ""  # If conversion fails, skip date suffix
            
            # Generate filename: last3digits_servicename_MM_YY.pdf
            base_filename = f"{last_3_digits}_{clean_name}{date_suffix}"
            
            # Add base name prefix for different PDF types
            if 'excess_only' in base_name:
                filename = f"{base_filename}_excess_only.pdf"
            elif 'all_slots' in base_name:
                filename = f"{base_filename}_all_slots.pdf"
            elif 'daywise' in base_name:
                filename = f"{base_filename}_daywise.pdf"
            else:
                filename = f"{base_filename}.pdf"
            
            return filename
            
            return filename

        # Defensive: If no PDF options are selected, default to generating 'all slots' PDF
        if not (show_excess_only or show_all_slots or show_daywise):
            show_all_slots = True
            print("DEBUG: No PDF options selected, defaulting to show_all_slots=True")
        else:
            print(f"DEBUG: PDF options selected - show_excess_only: {show_excess_only}, show_all_slots: {show_all_slots}, show_daywise: {show_daywise}")
        
        # Prepare full totals for PDF generation (always use totals from all data, not just excess)
        full_totals = {
            'iex_before': merged['IEX_Energy_kWh'].sum(),
            'cpp_before': merged['CPP_Energy_kWh'].sum(),
            'iex_after': merged['IEX_After_Loss'].sum(),
            'cpp_after': merged['CPP_After_Loss'].sum(),
            'iex_excess': merged['IEX_Excess'].sum(),
            'cpp_excess': merged['CPP_Excess'].sum()
        }
        
        import traceback
        if show_excess_only:
            try:
                print('DEBUG: Generating excess only PDF...')
                custom_filename = generate_custom_filename('energy_adjustment_excess_only.pdf', consumer_number, consumer_name, month, year)
                pdf_obj = generate_pdf(
                    merged_excess, sum_injection_excess, total_generated_after_loss_excess, comparison_excess, total_consumed_excess, total_excess_excess, excess_status, custom_filename, full_totals=full_totals)
                print('DEBUG: generate_pdf (excess only) returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append((custom_filename, pdf_obj))
            except Exception as e:
                print('ERROR in generate_pdf (excess only):', e)
                traceback.print_exc()
        if show_all_slots:
            try:
                print('DEBUG: Generating all slots PDF...')
                custom_filename = generate_custom_filename('energy_adjustment_all_slots.pdf', consumer_number, consumer_name, month, year)
                pdf_obj = generate_pdf(
                    merged_all, sum_injection_all, total_generated_after_loss_all, comparison_all, total_consumed_all, total_excess_all, excess_status, custom_filename, full_totals=full_totals)
                print('DEBUG: generate_pdf (all slots) returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append((custom_filename, pdf_obj))
            except Exception as e:
                print('ERROR in generate_pdf (all slots):', e)
                traceback.print_exc()
        if show_daywise:
            try:
                print('DEBUG: Generating daywise PDF...')
                custom_filename = generate_custom_filename('energy_adjustment_daywise.pdf', consumer_number, consumer_name, month, year)
                pdf_obj = generate_daywise_pdf(
                    merged_all, month, year, custom_filename)
                print('DEBUG: generate_daywise_pdf returned:', type(pdf_obj))
                if pdf_obj is not None:
                    pdfs.append((custom_filename, pdf_obj))
            except Exception as e:
                print('ERROR in generate_daywise_pdf:', e)
                traceback.print_exc()

        # If both PDFs, zip and send, else send single
        try:
            if len(pdfs) >= 2:
                print('DEBUG: Returning ZIP file to client')
                import zipfile
                zip_buffer = io.BytesIO()
                
                # Generate custom ZIP filename
                last_3_digits = str(consumer_number)[-3:] if len(str(consumer_number)) >= 3 else str(consumer_number)
                clean_name = "".join(c for c in consumer_name if c.isalnum() or c in (' ', '-', '_')).strip()
                clean_name = clean_name.replace(' ', '_')
                zip_filename = f"{last_3_digits}_{clean_name}_energy_adjustment_reports.zip"
                
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for fname, pdf_io in pdfs:
                        zf.writestr(fname, pdf_io.getvalue())
                zip_buffer.seek(0)
                return send_file(zip_buffer, as_attachment=True, download_name=zip_filename, mimetype='application/zip')
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
