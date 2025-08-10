import streamlit as st
import pandas as pd
import io
import os
import math
import json
from fpdf import FPDF
from datetime import datetime, timedelta
from pathlib import Path

# Import auto-updater
try:
    from auto_updater import initialize_updater, manual_update_check, show_update_settings
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False

# Set page config
st.set_page_config(
    page_title="Energy Adjustment Calculator",
    page_icon="⚡",
    layout="wide"
)

# Title and author
col1, col2 = st.columns([3, 1])
with col1:
    st.title("Energy Adjustment Calculator")
with col2:
    st.markdown("**Author: Er.Aravind MRT VREDC**")

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'error_message' not in st.session_state:
    st.session_state.error_message = None

def process_energy_data(generated_files, cpp_files, consumed_files,
                       enable_iex, enable_cpp, t_and_d_loss, cpp_t_and_d_loss,
                       consumer_number, consumer_name, multiplication_factor,
                       auto_detect_month, month, year, date_filter):
    """Process energy data files and return calculation results"""
    try:
        # Ensure file lists are not None
        if not enable_iex:
            generated_files = []
        if not enable_cpp:
            cpp_files = []
            
        # Process I.E.X generated energy Excel files (if provided)
        gen_df = None
        if generated_files and enable_iex and len(generated_files) > 0:
            gen_dfs = []
            for gen_file in generated_files:
                temp_df = pd.read_excel(gen_file, header=0)
                if temp_df.shape[1] < 3:
                    return {'success': False, 'error': f"Generated energy Excel file '{gen_file.name}' must have at least 3 columns: Date, Time, and Energy in MW."}
                
                # Add filename to help with debugging
                temp_df['Source_File'] = gen_file.name
                gen_dfs.append(temp_df)
            
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
                    st.warning(f"{nan_count} non-numeric Energy_MW values found in I.E.X files and converted to NaN")
                
                # Standardize date format to yyyy-mm-dd for robust filtering
                gen_df['Date'] = pd.to_datetime(gen_df['Date'], errors='coerce', dayfirst=True)
                gen_df['Source_Type'] = 'I.E.X'

        # Process C.P.P (Captive Power Purchase) files (if provided)
        cpp_df = None
        if cpp_files and enable_cpp and len(cpp_files) > 0:
            cpp_dfs = []
            for cpp_file in cpp_files:
                temp_df = pd.read_excel(cpp_file, header=0)
                if temp_df.shape[1] < 3:
                    return {'success': False, 'error': f"C.P.P energy Excel file '{cpp_file.name}' must have at least 3 columns: Date, Time, and Energy in MW."}
                
                # Add filename to help with debugging
                temp_df['Source_File'] = cpp_file.name
                temp_df['Source_Type'] = 'C.P.P'
                cpp_dfs.append(temp_df)
            
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
                    st.warning(f"{nan_count} non-numeric Energy_MW values found in C.P.P files and converted to NaN")
                
                cpp_df['Date'] = pd.to_datetime(cpp_df['Date'], errors='coerce', dayfirst=True)
                cpp_df['Source_Type'] = 'C.P.P'

        # Combine I.E.X and C.P.P data if both exist
        if gen_df is not None and cpp_df is not None:
            combined_gen_df = pd.concat([gen_df, cpp_df], ignore_index=True)
        elif gen_df is not None:
            combined_gen_df = gen_df
        elif cpp_df is not None:
            combined_gen_df = cpp_df
        else:
            return {'success': False, 'error': "No valid generation energy files were found."}
        
        gen_df = combined_gen_df

        # Function to convert month number to month name
        def get_month_name(month_num):
            month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                          'July', 'August', 'September', 'October', 'November', 'December']
            try:
                return month_names[int(month_num) - 1]
            except (ValueError, IndexError):
                return str(month_num)

        # Auto-detect month and year if enabled
        auto_detect_info = ""
        if auto_detect_month and not (month and year):
            # Extract unique months and years from the data
            unique_months = gen_df['Date'].dt.month.unique()
            unique_years = gen_df['Date'].dt.year.unique()
            
            if len(unique_months) == 1 and not month:
                month = str(int(unique_months[0]))
                st.info(f"Auto-detected month: {month} ({get_month_name(month)})")
            elif len(unique_months) > 1 and not month:
                # If multiple months, use the most frequent one
                month = str(int(gen_df['Date'].dt.month.value_counts().idxmax()))
                st.info(f"Multiple months detected, using most frequent: {month} ({get_month_name(month)})")
            
            if len(unique_years) == 1 and not year:
                year = str(int(unique_years[0]))
                st.info(f"Auto-detected year: {year}")
            elif len(unique_years) > 1 and not year:
                # If multiple years, use the most frequent one
                year = str(int(gen_df['Date'].dt.year.value_counts().idxmax()))
                st.info(f"Multiple years detected, using most frequent: {year}")
                
            # Add information to be displayed in PDF
            cpp_count = len(cpp_files) if cpp_files else 0
            iex_count = len(generated_files) if generated_files else 0
            auto_detect_info = f"Auto-detected from {iex_count} I.E.X, {cpp_count} C.P.P, and {len(consumed_files)} consumed files"

        # Process multiple consumed energy Excel files
        cons_dfs = []
        for cons_file in consumed_files:
            temp_df = pd.read_excel(cons_file, header=0)
            if temp_df.shape[1] < 3:
                return {'success': False, 'error': f"Consumed energy Excel file '{cons_file.name}' must have at least 3 columns: Date, Time, and Energy in kWh."}
            
            # Add filename to help with debugging
            temp_df['Source_File'] = cons_file.name
            cons_dfs.append(temp_df)
        
        # Combine all consumed energy dataframes
        if not cons_dfs:
            return {'success': False, 'error': "No valid consumed energy Excel files were found."}
        
        cons_df = pd.concat(cons_dfs, ignore_index=True)
        cons_df = cons_df.iloc[:, :3]
        cons_df.columns = ['Date', 'Time', 'Energy_kWh']
        # Strip whitespace from Date and Time columns
        cons_df['Date'] = cons_df['Date'].astype(str).str.strip()
        cons_df['Time'] = cons_df['Time'].astype(str).str.strip()
        # Standardize date format to yyyy-mm-dd for robust filtering
        cons_df['Date'] = pd.to_datetime(cons_df['Date'], errors='coerce', dayfirst=True)
        
        # Apply date filtering logic (simplified version)
        filtered_gen = gen_df.copy()
        filtered_cons = cons_df.copy()
        
        if year and month:
            try:
                year_int = int(float(year))
                month_int = int(float(month))
                
                # Filter generation data
                filtered_gen = filtered_gen[
                    (filtered_gen['Date'].dt.year == year_int) & 
                    (filtered_gen['Date'].dt.month == month_int)
                ]
                
                # Filter consumption data
                filtered_cons = filtered_cons[
                    (filtered_cons['Date'].dt.year == year_int) & 
                    (filtered_cons['Date'].dt.month == month_int)
                ]
                
            except ValueError:
                return {'success': False, 'error': f"Invalid year or month value. Year: '{year}', Month: '{month}'"}
        
        if date_filter:
            try:
                date_obj = pd.to_datetime(date_filter, dayfirst=True)
                filtered_gen = filtered_gen[filtered_gen['Date'] == date_obj]
                filtered_cons = filtered_cons[filtered_cons['Date'] == date_obj]
            except Exception:
                return {'success': False, 'error': f"Invalid date format for filter: {date_filter}. Use dd/mm/yyyy."}
        
        if (year or month or date_filter) and (filtered_gen.empty or filtered_cons.empty):
            return {'success': False, 'error': "No data found for the selected filters."}
        
        gen_df = filtered_gen
        cons_df = filtered_cons
        
        # Convert MW to kWh (MW * 250 for 15-minute intervals)
        gen_df['Energy_kWh'] = gen_df['Energy_MW'] * 250
        
        # Apply T&D losses based on source type
        def apply_td_loss(row):
            if row['Source_Type'] == 'I.E.X':
                return row['Energy_kWh'] * (1 - t_and_d_loss / 100) if t_and_d_loss > 0 else row['Energy_kWh']
            elif row['Source_Type'] == 'C.P.P':
                return row['Energy_kWh'] * (1 - cpp_t_and_d_loss / 100) if cpp_t_and_d_loss > 0 else row['Energy_kWh']
            else:
                return row['Energy_kWh']
        
        gen_df['After_Loss'] = gen_df.apply(apply_td_loss, axis=1)
        
        # Create slot time and date columns
        def slot_time_range(row):
            t = str(row['Time']).strip()
            if '-' in t:
                return t
            try:
                start = pd.to_datetime(t, format='%H:%M').time()
                end_dt = (pd.Timestamp.combine(pd.Timestamp.today(), start) + pd.Timedelta(minutes=15)).time()
                return f"{start.strftime('%H:%M')} - {end_dt.strftime('%H:%M')}"
            except Exception:
                return t
        
        gen_df['Slot_Time'] = gen_df.apply(slot_time_range, axis=1)
        gen_df['Slot_Time'] = gen_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
        gen_df['Slot_Date'] = gen_df['Date'].dt.strftime('%d/%m/%Y')
        
        # Apply same processing to consumption data
        cons_df['Energy_kWh'] = pd.to_numeric(cons_df['Energy_kWh'], errors='coerce') * multiplication_factor
        cons_df['Slot_Time'] = cons_df.apply(slot_time_range, axis=1)
        cons_df['Slot_Time'] = cons_df['Slot_Time'].replace({'23:45 - 24:00': '23:45 - 00:00'})
        cons_df['Slot_Date'] = cons_df['Date'].dt.strftime('%d/%m/%Y')

        # Sequential adjustment logic: First I.E.X, then C.P.P
        # Separate I.E.X and C.P.P data for sequential adjustment
        iex_df = gen_df[gen_df['Source_Type'] == 'I.E.X'].copy() if enable_iex else pd.DataFrame()
        cpp_df_only = gen_df[gen_df['Source_Type'] == 'C.P.P'].copy() if enable_cpp else pd.DataFrame()
        
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
        
        # Sort merged data chronologically by Slot_Date and Slot_Time
        def slot_time_to_minutes(slot_time):
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
            try:
                start_time = slot_time.split('-')[0].strip()
                hour, minute = map(int, start_time.split(':'))
                
                # Morning peak: 6:00 AM - 10:00 AM (C1)
                if 6 <= hour < 10:
                    return 'C1'
                # Evening peak: 6:00 PM - 10:00 PM (C2)
                elif 18 <= hour < 22:
                    return 'C2'
                # Normal hours: 5:00 AM - 6:00 AM + 10:00 AM to 6:00 PM (C4)
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
        
        # Clean up temporary columns
        merged.drop(['Slot_Date_dt', 'Slot_Time_min'], axis=1, inplace=True)
        
        # Calculate totals
        sum_injection = merged['Energy_kWh_gen'].sum()
        total_generated_after_loss = merged['After_Loss'].sum()
        total_consumed = merged['Energy_kWh_cons'].sum()
        total_excess = merged['Total_Excess'].sum()
        
        # For PDF, show all slots or only excess slots
        merged_excess = merged[merged['Total_Excess'] > 0].copy()
        merged_all = merged.copy()
        
        # Count unique days
        unique_days_gen = merged['Slot_Date'].nunique()
        unique_days_cons = merged['Slot_Date'].nunique()
        
        excess_status = 'Excess' if total_excess > 0 else 'No Excess'
        
        # Calculate TOD-wise excess for financial calculations
        tod_excess = merged.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
        
        # Helper function for consistent rounding throughout the application
        def round_kwh_financial(value):
            return int(value + 0.5) if value >= 0 else int(value - 0.5)
        
        # Round the total for financial calculations to match table display values
        total_excess_financial_rounded = round_kwh_financial(total_excess)
        
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
        
        # Calculate E-Tax (5%)
        etax = total_amount * 0.05
        
        # Calculate total amount with E-Tax
        total_with_etax = total_amount + etax
        
        # Calculate IEX excess for specific charges using rounded values
        iex_excess_financial_raw = merged['IEX_Excess'].sum()
        iex_excess_financial = round_kwh_financial(iex_excess_financial_raw)
        
        # Calculate negative factors using rounded values
        etax_on_iex = total_excess_financial_rounded * 0.1
        cross_subsidy_surcharge = iex_excess_financial * 1.92  # Only for IEX excess
        
        # Enhanced wheeling charges calculation for multiple sources
        wheeling_charges = 0
        
        # Get individual source excesses (rounded for financial calculations)
        iex_excess_for_wheeling = round_kwh_financial(merged['IEX_Excess'].sum()) if enable_iex else 0
        cpp_excess_for_wheeling = round_kwh_financial(merged['CPP_Excess'].sum()) if enable_cpp else 0
        
        # Calculate wheeling charges for IEX source (if enabled and has excess)
        if enable_iex and iex_excess_for_wheeling > 0:
            iex_loss_percentage = t_and_d_loss
            iex_x_value = (iex_excess_for_wheeling * iex_loss_percentage) / (100 - iex_loss_percentage) if iex_loss_percentage < 100 else 0
            iex_y_value = (iex_excess_for_wheeling * iex_loss_percentage / iex_x_value) + iex_excess_for_wheeling if iex_x_value != 0 else iex_excess_for_wheeling
            wheeling_charges += iex_y_value * 1.04
        
        # Calculate wheeling charges for CPP source (if enabled and has excess)
        if enable_cpp and cpp_excess_for_wheeling > 0:
            cpp_loss_percentage = cpp_t_and_d_loss
            cpp_x_value = (cpp_excess_for_wheeling * cpp_loss_percentage) / (100 - cpp_loss_percentage) if cpp_loss_percentage < 100 else 0
            cpp_y_value = (cpp_excess_for_wheeling * cpp_loss_percentage / cpp_x_value) + cpp_excess_for_wheeling if cpp_x_value != 0 else cpp_excess_for_wheeling
            wheeling_charges += cpp_y_value * 1.04
        
        # Calculate final amount to be collected
        final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
        
        # Round up final amount to next highest value
        final_amount_rounded = math.ceil(final_amount)

        return {'success': True, 'data': {
            'merged_all': merged_all,
            'merged_excess': merged_excess,
            'sum_injection': sum_injection,
            'total_generated_after_loss': total_generated_after_loss,
            'total_consumed': total_consumed,
            'total_excess': total_excess,
            'excess_status': excess_status,
            'unique_days_gen': unique_days_gen,
            'unique_days_cons': unique_days_cons,
            'month': month,
            'year': year,
            'auto_detect_info': auto_detect_info,
            'enable_iex': enable_iex,
            'enable_cpp': enable_cpp,
            't_and_d_loss': t_and_d_loss,
            'cpp_t_and_d_loss': cpp_t_and_d_loss,
            'consumer_number': consumer_number,
            'consumer_name': consumer_name,
            'multiplication_factor': multiplication_factor,
            # Financial calculation results
            'total_excess_financial_rounded': total_excess_financial_rounded,
            'base_rate': base_rate,
            'base_amount': base_amount,
            'c1_c2_excess': c1_c2_excess,
            'c1_c2_additional': c1_c2_additional,
            'c5_excess': c5_excess,
            'c5_additional': c5_additional,
            'total_amount': total_amount,
            'etax': etax,
            'total_with_etax': total_with_etax,
            'iex_excess_financial': iex_excess_financial,
            'etax_on_iex': etax_on_iex,
            'cross_subsidy_surcharge': cross_subsidy_surcharge,
            'wheeling_charges': wheeling_charges,
            'final_amount': final_amount,
            'final_amount_rounded': final_amount_rounded,
            'message': "Data processing completed successfully!"
        }}
        
    except Exception as e:
        return {'success': False, 'error': str(e)}

def generate_custom_filename(base_name, consumer_number, consumer_name, month=None, year=None):
    """Generate custom filename based on service number and name"""
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

def generate_detailed_pdf(data, pdf_data, pdf_type):
    """Generate detailed PDF with complete table data and calculations"""
    try:
        # Import datetime for timestamp
        from datetime import datetime
        
        # Initialize breakdown variable at function level
        wheeling_breakdown = []
        
        pdf = FPDF()
        pdf.set_margins(20, 20, 20)  # Set proper margins: left, top, right (20mm each)
        pdf.set_auto_page_break(auto=True, margin=20)  # Auto page break with bottom margin
        pdf.add_page()
        
        # FIRST PAGE - DESCRIPTION AND INFORMATION ONLY
        pdf.set_font('Arial', 'B', 16)  # Larger title font
        title = f"Energy Adjustment {pdf_type.replace('_', ' ').title()} Report"
        pdf.cell(0, 15, title, ln=True, align='C')
        pdf.ln(10)
        
        # Consumer Information Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Consumer Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f"Consumer Number: {data['consumer_number']}", ln=True)
        pdf.cell(0, 8, f"Consumer Name: {data['consumer_name']}", ln=True)
        pdf.cell(0, 8, f"Multiplication Factor (Consumed Energy): {data.get('multiplication_factor', 1)}", ln=True)
        pdf.ln(5)
        
        # Technical Parameters Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Technical Parameters:', ln=True)
        pdf.set_font('Arial', '', 12)
        
        # Display T&D losses based on what sources are used
        if data.get('enable_iex') and data.get('enable_cpp'):
            pdf.cell(0, 8, f"I.E.X T&D Loss (%): {data.get('t_and_d_loss', 0)}", ln=True)
            pdf.cell(0, 8, f"C.P.P T&D Loss (%): {data.get('cpp_t_and_d_loss', 0)}", ln=True)
        elif data.get('enable_iex'):
            pdf.cell(0, 8, f"I.E.X T&D Loss (%): {data.get('t_and_d_loss', 0)}", ln=True)
        elif data.get('enable_cpp'):
            pdf.cell(0, 8, f"C.P.P T&D Loss (%): {data.get('cpp_t_and_d_loss', 0)}", ln=True)
        pdf.ln(5)
        
        # Data Sources Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Data Sources:', ln=True)
        pdf.set_font('Arial', '', 12)
        # Add count of files used (placeholder - can be enhanced with actual file count)
        if data.get('enable_iex'):
            pdf.cell(0, 8, f'Generated Energy Files (I.E.X): Available', ln=True)
        if data.get('enable_cpp'):
            pdf.cell(0, 8, f'Generated Energy Files (C.P.P): Available', ln=True)
        pdf.cell(0, 8, f'Consumed Energy Files: Available', ln=True)
        pdf.ln(5)
        
        # Report Information Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Report Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f"Report Generated: {datetime.now().strftime('%d/%m/%Y at %H:%M:%S')}", ln=True)
        if data.get('month') and data.get('year'):
            pdf.cell(0, 8, f"Report Period: {data['month']}/{data['year']}", ln=True)
        if data.get('auto_detect_info'):
            pdf.cell(0, 8, f"Period Details: {data['auto_detect_info']}", ln=True)
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
        
        # Helper functions for safe formatting
        def safe_date_str(date_val):
            """Safely convert date to string"""
            if pd.isna(date_val):
                return ""
            return str(date_val)
        
        def format_time(time_val):
            """Format time value for display"""
            if pd.isna(time_val):
                return ""
            return str(time_val)
        
        # Function to add table headers with proper text wrapping
        def add_table_headers():
            is_dual_source = data.get('enable_iex') and data.get('enable_cpp')
            
            if is_dual_source:
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
        is_dual_source = data.get('enable_iex') and data.get('enable_cpp')
        if is_dual_source:
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
            
            if is_dual_source:
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
        
        # Calculate totals
        total_excess = data['total_excess']
        total_consumed = data['total_consumed']
        total_generated_after_loss = data['total_generated_after_loss']
        
        # Enhanced summary for sequential adjustment
        if is_dual_source:
            # Sequential adjustment summary - use rounded totals from table data for precision
            total_iex_before_loss = data.get('merged_all', pdf_data)['IEX_Energy_kWh'].sum()
            total_cpp_before_loss = data.get('merged_all', pdf_data)['CPP_Energy_kWh'].sum()
            total_iex_after_loss = data.get('merged_all', pdf_data)['IEX_After_Loss'].sum()
            total_cpp_after_loss = data.get('merged_all', pdf_data)['CPP_After_Loss'].sum()
            total_iex_excess = data.get('merged_all', pdf_data)['IEX_Excess'].sum()
            total_cpp_excess = data.get('merged_all', pdf_data)['CPP_Excess'].sum()
            
            # Round all values
            total_iex_before_loss_rounded = round_kwh_summary(total_iex_before_loss)
            total_iex_after_loss_rounded = round_kwh_summary(total_iex_after_loss)
            total_cpp_before_loss_rounded = round_kwh_summary(total_cpp_before_loss)
            total_cpp_after_loss_rounded = round_kwh_summary(total_cpp_after_loss)
            total_iex_excess_rounded = round_kwh_summary(total_iex_excess)
            total_cpp_excess_rounded = round_kwh_summary(total_cpp_excess)
            
            pdf.cell(0, 8, f'I.E.X Generation (before T&D loss): {total_iex_before_loss_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'I.E.X Generation (after {data.get("t_and_d_loss", 0)}% T&D loss): {total_iex_after_loss_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'I.E.X Excess Energy (rounded): {total_iex_excess_rounded} kWh', ln=True)
            pdf.ln(3)
            
            pdf.cell(0, 8, f'C.P.P Generation (before T&D loss): {total_cpp_before_loss_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'C.P.P Generation (after {data.get("cpp_t_and_d_loss", 0)}% T&D loss): {total_cpp_after_loss_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'C.P.P Excess Energy (rounded): {total_cpp_excess_rounded} kWh', ln=True)
            pdf.ln(3)
            
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(0, 8, 'TOTAL CALCULATIONS:', ln=True)
            pdf.set_font('Arial', '', 11)
            total_generation_before_rounded = round_kwh_summary(total_iex_before_loss + total_cpp_before_loss)
            total_generation_after_rounded = round_kwh_summary(total_iex_after_loss + total_cpp_after_loss)
            total_consumed_rounded = round_kwh_summary(total_consumed)
            total_excess_rounded = total_iex_excess_rounded + total_cpp_excess_rounded
            
            pdf.cell(0, 8, f'Total Generation (before loss): {total_generation_before_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'Total Generation (after loss): {total_generation_after_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'Total Consumed Energy: {total_consumed_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
        else:
            # Single source summary
            total_excess_rounded = round_kwh_summary(total_excess)
            total_consumed_rounded = round_kwh_summary(total_consumed)
            total_generated_after_loss_rounded = round_kwh_summary(total_generated_after_loss)
            
            if data.get('enable_iex'):
                pdf.cell(0, 8, f'I.E.X Generation (after {data.get("t_and_d_loss", 0)}% T&D loss): {total_generated_after_loss_rounded} kWh', ln=True)
            elif data.get('enable_cpp'):
                pdf.cell(0, 8, f'C.P.P Generation (after {data.get("cpp_t_and_d_loss", 0)}% T&D loss): {total_generated_after_loss_rounded} kWh', ln=True)
            
            pdf.cell(0, 8, f'Total Consumed Energy (after multiplication): {total_consumed_rounded} kWh', ln=True)
            pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
        
        pdf.cell(0, 8, f'Unique Days Used (Generated): {data["unique_days_gen"]}', ln=True)
        pdf.cell(0, 8, f'Unique Days Used (Consumed): {data["unique_days_cons"]}', ln=True)
        pdf.cell(0, 8, f'Status: {data["excess_status"]}', ln=True)
        
        # Check if we need a new page for TOD breakdown
        if pdf.get_y() > 220:
            pdf.add_page()
        
        # Add TOD-wise excess energy breakdown
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(20, 10, 'TOD', 1)
        pdf.cell(50, 10, 'Excess Energy (kWh)', 1)
        pdf.ln()
        
        pdf.set_font('Arial', '', 10)
        
        # Calculate TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
        
        # Calculate C category total (sum of C1, C2, C4, C5)
        c_categories = ['C1', 'C2', 'C4', 'C5']
        c_total = 0
        tod_values = {}
        
        for _, row in tod_excess.iterrows():
            category = row['TOD_Category']
            excess_rounded = round_kwh_summary(row['Total_Excess'])
            tod_values[category] = excess_rounded
            if category in c_categories:
                c_total += excess_rounded
        
        # Display C total first
        pdf.cell(20, 10, 'C', 1)
        pdf.cell(50, 10, f"{c_total}", 1)
        pdf.ln()
        
        # Display individual categories
        for category in ['C1', 'C2', 'C4', 'C5']:
            if category in tod_values:
                pdf.cell(20, 10, category, 1)
                pdf.cell(50, 10, f"{tod_values[category]}", 1)
                pdf.ln()
        
        if 'Unknown' in tod_values:
            pdf.cell(20, 10, 'Unknown', 1)
            pdf.cell(50, 10, f"{tod_values['Unknown']}", 1)
            pdf.ln()
        
        # Add financial calculations (using pre-calculated values from data processing)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Financial Calculations:', ln=True)
        pdf.set_font('Arial', '', 10)
        
        # Use pre-calculated financial values if available, otherwise calculate on the fly
        if 'total_excess_financial_rounded' in data:
            # Use pre-calculated values
            total_excess_financial_rounded = data['total_excess_financial_rounded']
            base_rate = data['base_rate']
            base_amount = data['base_amount']
            c1_c2_excess = data['c1_c2_excess']
            c1_c2_additional = data['c1_c2_additional']
            c5_excess = data['c5_excess']
            c5_additional = data['c5_additional']
            total_amount = data['total_amount']
            etax = data['etax']
            total_with_etax = data['total_with_etax']
            iex_excess_financial = data['iex_excess_financial']
            etax_on_iex = data['etax_on_iex']
            cross_subsidy_surcharge = data['cross_subsidy_surcharge']
            wheeling_charges = data['wheeling_charges']
            final_amount = data['final_amount']
            final_amount_rounded = data['final_amount_rounded']
            
            # Also need these variables for display purposes
            loss_percentage = data.get('t_and_d_loss', 0)
            x_value = 0
            y_value = total_excess_financial_rounded
            
            # Recalculate x_value and y_value for display
            if loss_percentage < 100:
                x_value = (total_excess_financial_rounded * loss_percentage) / (100 - loss_percentage)
                if x_value != 0:
                    y_value = (total_excess_financial_rounded * loss_percentage / x_value) + total_excess_financial_rounded
                else:
                    y_value = total_excess_financial_rounded
        else:
            # Calculate on the fly (fallback)
            total_excess_financial_rounded = round_kwh_summary(total_excess)
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess_financial_rounded * base_rate
            c1_c2_excess = tod_values.get('C1', 0) + tod_values.get('C2', 0)
            c1_c2_additional = c1_c2_excess * 1.8125  # rupees per kWh
            c5_excess = tod_values.get('C5', 0)
            c5_additional = c5_excess * 0.3625  # rupees per kWh
            total_amount = base_amount + c1_c2_additional + c5_additional
            etax = total_amount * 0.05
            total_with_etax = total_amount + etax
            
            # Calculate IEX excess for specific charges
            iex_excess_financial_raw = 0
            if 'IEX_Excess' in pdf_data.columns:
                iex_excess_financial_raw = pdf_data['IEX_Excess'].sum()
            elif data.get('enable_iex') and not data.get('enable_cpp'):
                iex_excess_financial_raw = total_excess
            
            iex_excess_financial = round_kwh_summary(iex_excess_financial_raw)
            etax_on_iex = total_excess_financial_rounded * 0.1
            cross_subsidy_surcharge = iex_excess_financial * 1.92  # Only for IEX excess
            
            # Enhanced wheeling charges calculation for multiple sources (fallback calculation)
            wheeling_charges = 0
            wheeling_breakdown = []
            
            # Get individual source excesses for PDF calculation
            iex_excess_pdf = round_kwh_summary(pdf_data['IEX_Excess'].sum()) if 'IEX_Excess' in pdf_data.columns else 0
            cpp_excess_pdf = round_kwh_summary(pdf_data['CPP_Excess'].sum()) if 'CPP_Excess' in pdf_data.columns else 0
            
            # Calculate wheeling charges for IEX source (if enabled and has excess)
            if data.get('enable_iex') and iex_excess_pdf > 0:
                iex_loss_percentage = data.get('t_and_d_loss', 0)
                iex_x_value = (iex_excess_pdf * iex_loss_percentage) / (100 - iex_loss_percentage) if iex_loss_percentage < 100 else 0
                iex_y_value = (iex_excess_pdf * iex_loss_percentage / iex_x_value) + iex_excess_pdf if iex_x_value != 0 else iex_excess_pdf
                iex_wheeling = iex_y_value * 1.04
                wheeling_charges += iex_wheeling
                wheeling_breakdown.append(('IEX', iex_excess_pdf, iex_loss_percentage, iex_x_value, iex_y_value, iex_wheeling))
            
            # Calculate wheeling charges for CPP source (if enabled and has excess)
            if data.get('enable_cpp') and cpp_excess_pdf > 0:
                cpp_loss_percentage = data.get('cpp_t_and_d_loss', 0)
                cpp_x_value = (cpp_excess_pdf * cpp_loss_percentage) / (100 - cpp_loss_percentage) if cpp_loss_percentage < 100 else 0
                cpp_y_value = (cpp_excess_pdf * cpp_loss_percentage / cpp_x_value) + cpp_excess_pdf if cpp_x_value != 0 else cpp_excess_pdf
                cpp_wheeling = cpp_y_value * 1.04
                wheeling_charges += cpp_wheeling
                wheeling_breakdown.append(('CPP', cpp_excess_pdf, cpp_loss_percentage, cpp_x_value, cpp_y_value, cpp_wheeling))
            
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
            final_amount_rounded = math.ceil(final_amount)
        
        # Display the financial calculations with proper formatting
        pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess_financial_rounded} kWh) x Rs.{base_rate} = Rs.{base_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess} kWh) x Rs.1.8125 = Rs.{c1_c2_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"4. Partial Total: Rs.{base_amount:.2f} + Rs.{c1_c2_additional:.2f} + Rs.{c5_additional:.2f} = Rs.{total_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"5. E-Tax (5% of Partial Total): Rs.{total_amount:.2f} x 0.05 = Rs.{etax:.2f}", ln=True)
        pdf.cell(0, 8, f"6. Subtotal with E-Tax: Rs.{total_amount:.2f} + Rs.{etax:.2f} = Rs.{total_with_etax:.2f}", ln=True)
        pdf.cell(0, 8, f"7. Less: E-Tax on IEX: Total Excess ({total_excess_financial_rounded} kWh) x Rs.0.1 = Rs.{etax_on_iex:.2f}", ln=True)
        pdf.cell(0, 8, f"8. Less: Cross Subsidy Surcharge: IEX Excess ({iex_excess_financial} kWh) x Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
        # Enhanced wheeling charges display with breakdown
        pdf.cell(0, 8, f"9. Wheeling Charges Calculation:", ln=True)
        
        # Initialize breakdown if not already set
        if 'wheeling_breakdown' not in locals():
            wheeling_breakdown = []
        
        if wheeling_breakdown:
            for i, (source, excess, loss_pct, x_val, y_val, wheeling) in enumerate(wheeling_breakdown):
                pdf.cell(0, 8, f"9{chr(97+i*2)}. {source} Wheeling Charges Step 1:", ln=True)
                pdf.cell(0, 8, f"    [{excess} x {loss_pct:.2f}% / (100-{loss_pct:.2f}%)] = {x_val:.4f}", ln=True)
                pdf.cell(0, 8, f"9{chr(97+i*2+1)}. {source} Wheeling Charges Step 2:", ln=True)
                pdf.cell(0, 8, f"    [({excess} x {loss_pct:.2f}% / {x_val:.4f}) + {excess}] x 1.04 = Rs.{wheeling:.2f}", ln=True)
            
            if len(wheeling_breakdown) > 1:
                total_wheeling_components = " + ".join([f"Rs.{w[5]:.2f}" for w in wheeling_breakdown])
                pdf.cell(0, 8, f"9z. Total Wheeling Charges: {total_wheeling_components} = Rs.{wheeling_charges:.2f}", ln=True)
        else:
            # Fallback to simple display if breakdown not available
            pdf.cell(0, 8, f"9. Wheeling Charges: Rs.{wheeling_charges:.2f}", ln=True)
        
        # Calculate deductions total for clarity
        deductions_total = etax_on_iex + cross_subsidy_surcharge + wheeling_charges
        pdf.cell(0, 8, f"10a. Total Amount to be Collected - Step 1:", ln=True)
        pdf.cell(0, 8, f"     Rs.{total_with_etax:.2f} - (Rs.{etax_on_iex:.2f} + Rs.{cross_subsidy_surcharge:.2f} + Rs.{wheeling_charges:.2f})", ln=True)
        pdf.cell(0, 8, f"10b. Total Amount to be Collected - Step 2:", ln=True)
        pdf.cell(0, 8, f"     Rs.{total_with_etax:.2f} - Rs.{deductions_total:.2f} = Rs.{final_amount:.2f}", ln=True)
        
        # Round up final amount to next highest value
        pdf.set_font('Arial', 'B', 10)  # Consistent with table data font size
        pdf.cell(0, 8, f"11. Final Amount (Rounded Up): Rs.{final_amount_rounded}", ln=True)
        
        # Generate PDF bytes
        pdf_output = io.BytesIO()
        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output.write(pdf_bytes)
        pdf_output.seek(0)
        return pdf_output.getvalue()
        
    except Exception as e:
        st.error(f"Error generating detailed PDF: {str(e)}")
        return None

def generate_daywise_pdf(data, pdf_data):
    """Generate day-wise summary PDF"""
    try:
        # Initialize breakdown variable at function level
        wheeling_breakdown = []
        
        pdf = FPDF()
        pdf.set_margins(20, 20, 20)
        pdf.set_auto_page_break(auto=True, margin=20)
        pdf.add_page()
        
        # FIRST PAGE - DESCRIPTION AND INFORMATION ONLY
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 15, 'Energy Adjustment Day-wise Summary Report', ln=True, align='C')
        pdf.ln(10)
        
        # Consumer Information Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Consumer Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f"Consumer Number: {data['consumer_number']}", ln=True)
        pdf.cell(0, 8, f"Consumer Name: {data['consumer_name']}", ln=True)
        pdf.cell(0, 8, f"Multiplication Factor (Consumed Energy): {data.get('multiplication_factor', 1)}", ln=True)
        pdf.ln(5)
        
        # Technical Parameters Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Technical Parameters:', ln=True)
        pdf.set_font('Arial', '', 12)
        
        if data.get('enable_iex') and data.get('enable_cpp'):
            pdf.cell(0, 8, f"I.E.X T&D Loss (%): {data.get('t_and_d_loss', 0)}", ln=True)
            pdf.cell(0, 8, f"C.P.P T&D Loss (%): {data.get('cpp_t_and_d_loss', 0)}", ln=True)
        elif data.get('enable_iex'):
            pdf.cell(0, 8, f"I.E.X T&D Loss (%): {data.get('t_and_d_loss', 0)}", ln=True)
        elif data.get('enable_cpp'):
            pdf.cell(0, 8, f"C.P.P T&D Loss (%): {data.get('cpp_t_and_d_loss', 0)}", ln=True)
        pdf.ln(5)
        
        # Report Information Section
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Report Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f"Report Generated: {datetime.now().strftime('%d/%m/%Y at %H:%M:%S')}", ln=True)
        if data.get('month') and data.get('year'):
            pdf.cell(0, 8, f"Report Period: {data['month']}/{data['year']}", ln=True)
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
        
        # Table headers
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(40, 10, 'Date', 1)
        pdf.cell(50, 10, 'Total Gen. After Loss', 1)
        pdf.cell(50, 10, 'Total Consumed', 1)
        pdf.cell(50, 10, 'Total Excess', 1)
        pdf.ln()
        
        # Calculate day-wise data
        if data.get('enable_iex') and data.get('enable_cpp'):
            daywise = pdf_data.groupby('Slot_Date').agg({
                'IEX_After_Loss': 'sum',
                'CPP_After_Loss': 'sum',
                'Energy_kWh_cons': 'sum',
                'Total_Excess': 'sum'
            }).reset_index()
            daywise['Total_After_Loss'] = daywise['IEX_After_Loss'] + daywise['CPP_After_Loss']
        else:
            daywise = pdf_data.groupby('Slot_Date').agg({
                'After_Loss': 'sum',
                'Energy_kWh_cons': 'sum',
                'Total_Excess': 'sum'
            }).reset_index()
            daywise['Total_After_Loss'] = daywise['After_Loss']
        
        pdf.set_font('Arial', '', 8)
        for idx, row in daywise.iterrows():
            pdf.cell(40, 10, str(row['Slot_Date']), 1)
            pdf.cell(50, 10, f"{row['Total_After_Loss']:.4f}", 1)
            pdf.cell(50, 10, f"{row['Energy_kWh_cons']:.4f}", 1)
            
            # Round excess values for display
            total_excess_rounded = int(row['Total_Excess'] + 0.5) if row['Total_Excess'] >= 0 else int(row['Total_Excess'] - 0.5)
            pdf.cell(50, 10, f"{total_excess_rounded}", 1)
            pdf.ln()
        
        # Add same calculation summary as detailed PDF
        pdf.ln(10)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'DETAILED CALCULATION SUMMARY:', ln=True)
        pdf.set_font('Arial', '', 11)
        
        # Helper function for proper rounding
        def round_kwh(value):
            return int(value + 0.5) if value >= 0 else int(value - 0.5)
        
        # Include same financial calculations as detailed PDF
        total_excess = data['total_excess']
        total_excess_rounded = round_kwh(total_excess)
        total_consumed = data['total_consumed']
        total_consumed_rounded = round_kwh(total_consumed)
        total_generated_after_loss = data['total_generated_after_loss']
        total_generated_after_loss_rounded = round_kwh(total_generated_after_loss)
        
        # Basic summary
        pdf.cell(0, 8, f'Total Generated (after loss): {total_generated_after_loss_rounded} kWh', ln=True)
        pdf.cell(0, 8, f'Total Consumed Energy (after multiplication): {total_consumed_rounded} kWh', ln=True)
        pdf.cell(0, 8, f'Total Excess Energy (rounded): {total_excess_rounded} kWh', ln=True)
        pdf.cell(0, 8, f'Unique Days Used (Generated): {data["unique_days_gen"]}', ln=True)
        pdf.cell(0, 8, f'Unique Days Used (Consumed): {data["unique_days_cons"]}', ln=True)
        pdf.cell(0, 8, f'Status: {data["excess_status"]}', ln=True)
        
        # Check if we need a new page for TOD breakdown
        if pdf.get_y() > 220:
            pdf.add_page()
        
        # Add TOD-wise excess energy breakdown
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(20, 10, 'TOD', 1)
        pdf.cell(50, 10, 'Excess Energy (kWh)', 1)
        pdf.ln()
        
        pdf.set_font('Arial', '', 10)
        
        # Calculate TOD-wise excess from the dataframe
        tod_excess = pdf_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
        
        # Calculate C category total (sum of C1, C2, C4, C5)
        c_categories = ['C1', 'C2', 'C4', 'C5']
        c_total = 0
        tod_values = {}
        
        for _, row in tod_excess.iterrows():
            category = row['TOD_Category']
            excess_rounded = round_kwh(row['Total_Excess'])
            tod_values[category] = excess_rounded
            if category in c_categories:
                c_total += excess_rounded
        
        # Display C total first
        pdf.cell(20, 10, 'C', 1)
        pdf.cell(50, 10, f"{c_total}", 1)
        pdf.ln()
        
        # Display individual categories
        for category in ['C1', 'C2', 'C4', 'C5']:
            if category in tod_values:
                pdf.cell(20, 10, category, 1)
                pdf.cell(50, 10, f"{tod_values[category]}", 1)
                pdf.ln()
        
        if 'Unknown' in tod_values:
            pdf.cell(20, 10, 'Unknown', 1)
            pdf.cell(50, 10, f"{tod_values['Unknown']}", 1)
            pdf.ln()
        
        # Add financial calculations (using pre-calculated values from data processing)
        pdf.ln(5)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Financial Calculations:', ln=True)
        pdf.set_font('Arial', '', 10)
        
        # Use pre-calculated financial values if available, otherwise calculate on the fly
        if 'total_excess_financial_rounded' in data:
            # Use pre-calculated values
            total_excess_financial_rounded = data['total_excess_financial_rounded']
            base_rate = data['base_rate']
            base_amount = data['base_amount']
            c1_c2_excess = data['c1_c2_excess']
            c1_c2_additional = data['c1_c2_additional']
            c5_excess = data['c5_excess']
            c5_additional = data['c5_additional']
            total_amount = data['total_amount']
            etax = data['etax']
            total_with_etax = data['total_with_etax']
            iex_excess_financial = data['iex_excess_financial']
            etax_on_iex = data['etax_on_iex']
            cross_subsidy_surcharge = data['cross_subsidy_surcharge']
            wheeling_charges = data['wheeling_charges']
            final_amount = data['final_amount']
            final_amount_rounded = data['final_amount_rounded']
            
            # Also need these variables for display purposes
            loss_percentage = data.get('t_and_d_loss', 0)
            x_value = 0
            y_value = total_excess_financial_rounded
            
            # Recalculate x_value and y_value for display
            if loss_percentage < 100:
                x_value = (total_excess_financial_rounded * loss_percentage) / (100 - loss_percentage)
                if x_value != 0:
                    y_value = (total_excess_financial_rounded * loss_percentage / x_value) + total_excess_financial_rounded
                else:
                    y_value = total_excess_financial_rounded
        else:
            # Calculate on the fly (fallback)
            total_excess_financial_rounded = round_kwh(total_excess)
            base_rate = 7.25  # rupees per kWh
            base_amount = total_excess_financial_rounded * base_rate
            c1_c2_excess = tod_values.get('C1', 0) + tod_values.get('C2', 0)
            c1_c2_additional = c1_c2_excess * 1.8125  # rupees per kWh
            c5_excess = tod_values.get('C5', 0)
            c5_additional = c5_excess * 0.3625  # rupees per kWh
            total_amount = base_amount + c1_c2_additional + c5_additional
            etax = total_amount * 0.05
            total_with_etax = total_amount + etax
            
            # Calculate IEX excess for specific charges
            iex_excess_financial_raw = 0
            if data.get('enable_iex') and data.get('enable_cpp') and 'IEX_Excess' in pdf_data.columns:
                iex_excess_financial_raw = pdf_data['IEX_Excess'].sum()
            elif data.get('enable_iex') and not data.get('enable_cpp'):
                iex_excess_financial_raw = total_excess
            
            iex_excess_financial = round_kwh(iex_excess_financial_raw)
            etax_on_iex = total_excess_financial_rounded * 0.1
            cross_subsidy_surcharge = iex_excess_financial * 1.92  # Only for IEX excess
            
            # Enhanced wheeling charges calculation for multiple sources (daywise fallback)
            wheeling_charges = 0
            wheeling_breakdown = []
            
            # Get individual source excesses for daywise calculations
            iex_excess_daywise = round_kwh(pdf_data['IEX_Excess'].sum()) if 'IEX_Excess' in pdf_data.columns else 0
            cpp_excess_daywise = round_kwh(pdf_data['CPP_Excess'].sum()) if 'CPP_Excess' in pdf_data.columns else 0
            
            # Calculate wheeling charges for IEX source (if enabled and has excess)
            if data.get('enable_iex') and iex_excess_daywise > 0:
                iex_loss_percentage = data.get('t_and_d_loss', 0)
                iex_x_value = (iex_excess_daywise * iex_loss_percentage) / (100 - iex_loss_percentage) if iex_loss_percentage < 100 else 0
                iex_y_value = (iex_excess_daywise * iex_loss_percentage / iex_x_value) + iex_excess_daywise if iex_x_value != 0 else iex_excess_daywise
                iex_wheeling = iex_y_value * 1.04
                wheeling_charges += iex_wheeling
                wheeling_breakdown.append(('IEX', iex_excess_daywise, iex_loss_percentage, iex_x_value, iex_y_value, iex_wheeling))
            
            # Calculate wheeling charges for CPP source (if enabled and has excess)
            if data.get('enable_cpp') and cpp_excess_daywise > 0:
                cpp_loss_percentage = data.get('cpp_t_and_d_loss', 0)
                cpp_x_value = (cpp_excess_daywise * cpp_loss_percentage) / (100 - cpp_loss_percentage) if cpp_loss_percentage < 100 else 0
                cpp_y_value = (cpp_excess_daywise * cpp_loss_percentage / cpp_x_value) + cpp_excess_daywise if cpp_x_value != 0 else cpp_excess_daywise
                cpp_wheeling = cpp_y_value * 1.04
                wheeling_charges += cpp_wheeling
                wheeling_breakdown.append(('CPP', cpp_excess_daywise, cpp_loss_percentage, cpp_x_value, cpp_y_value, cpp_wheeling))
            
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
            final_amount_rounded = math.ceil(final_amount)
        
        # Display simplified financial summary for daywise PDF
        pdf.cell(0, 8, f"1. Base Rate: Total Excess Energy ({total_excess_financial_rounded} kWh) x Rs.{base_rate} = Rs.{base_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"2. C1+C2 Additional: Excess in C1+C2 ({c1_c2_excess} kWh) x Rs.1.8125 = Rs.{c1_c2_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"3. C5 Additional: Excess in C5 ({c5_excess} kWh) x Rs.0.3625 = Rs.{c5_additional:.2f}", ln=True)
        pdf.cell(0, 8, f"4. Partial Total: Rs.{total_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"5. E-Tax (5%): Rs.{etax:.2f}", ln=True)
        pdf.cell(0, 8, f"6. Subtotal with E-Tax: Rs.{total_with_etax:.2f}", ln=True)
        pdf.cell(0, 8, f"7. Less: E-Tax on IEX: Rs.{etax_on_iex:.2f}", ln=True)
        pdf.cell(0, 8, f"8. Less: Cross Subsidy Surcharge: Rs.{cross_subsidy_surcharge:.2f}", ln=True)
            
        # Enhanced wheeling charges display (simplified for daywise PDF)
        pdf.cell(0, 8, f"9. Wheeling Charges Calculation:", ln=True)
        if wheeling_breakdown:
            for i, (source, excess, loss_pct, x_val, y_val, wheeling) in enumerate(wheeling_breakdown):
                pdf.cell(0, 8, f"9{chr(97+i)}. {source}: [{excess} x {loss_pct:.2f}% / (100-{loss_pct:.2f}%)] x 1.04 = Rs.{wheeling:.2f}", ln=True)
            
            if len(wheeling_breakdown) > 1:
                total_wheeling_components = " + ".join([f"Rs.{w[5]:.2f}" for w in wheeling_breakdown])
                pdf.cell(0, 8, f"9z. Total Wheeling Charges: {total_wheeling_components} = Rs.{wheeling_charges:.2f}", ln=True)
        else:
            pdf.cell(0, 8, f"9. Wheeling Charges: Rs.{wheeling_charges:.2f}", ln=True)
        
        # Calculate deductions total for clarity
        deductions_total = etax_on_iex + cross_subsidy_surcharge + wheeling_charges
        
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, f"10. Final Amount: Rs.{total_with_etax:.2f} - Rs.{deductions_total:.2f} = Rs.{final_amount:.2f}", ln=True)
        pdf.cell(0, 8, f"11. Final Amount (Rounded Up): Rs.{final_amount_rounded}", ln=True)
        
        # Generate PDF bytes
        pdf_output = io.BytesIO()
        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        pdf_output.write(pdf_bytes)
        pdf_output.seek(0)
        return pdf_output.getvalue()
        
    except Exception as e:
        st.error(f"Error generating daywise PDF: {str(e)}")
        return None

def generate_simple_pdf(data, pdf_type="excess"):
    """Generate a simple PDF report"""
    try:
        pdf = FPDF()
        pdf.set_margins(20, 20, 20)
        pdf.set_auto_page_break(auto=True, margin=20)
        pdf.add_page()
        
        # Title
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 15, 'Energy Adjustment Report', ln=True, align='C')
        pdf.ln(10)
        
        # Consumer Information
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Consumer Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f'Consumer Number: {data["consumer_number"]}', ln=True)
        pdf.cell(0, 8, f'Consumer Name: {data["consumer_name"]}', ln=True)
        pdf.cell(0, 8, f'Multiplication Factor: {data["multiplication_factor"]}', ln=True)
        pdf.ln(5)
        
        # Technical Parameters
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Technical Parameters:', ln=True)
        pdf.set_font('Arial', '', 12)
        
        if data['enable_iex'] and data['enable_cpp']:
            pdf.cell(0, 8, f'I.E.X T&D Loss (%): {data["t_and_d_loss"]}', ln=True)
            pdf.cell(0, 8, f'C.P.P T&D Loss (%): {data["cpp_t_and_d_loss"]}', ln=True)
        elif data['enable_iex']:
            pdf.cell(0, 8, f'I.E.X T&D Loss (%): {data["t_and_d_loss"]}', ln=True)
        elif data['enable_cpp']:
            pdf.cell(0, 8, f'C.P.P T&D Loss (%): {data["cpp_t_and_d_loss"]}', ln=True)
        pdf.ln(5)
        
        # Summary
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Summary:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f'Total Generated (After Loss): {data["total_generated_after_loss"]:.2f} kWh', ln=True)
        pdf.cell(0, 8, f'Total Consumed: {data["total_consumed"]:.2f} kWh', ln=True)
        pdf.cell(0, 8, f'Total Excess: {data["total_excess"]:.2f} kWh', ln=True)
        pdf.cell(0, 8, f'Status: {data["excess_status"]}', ln=True)
        pdf.ln(5)
        
        # Report Info
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, 'Report Information:', ln=True)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 8, f'Report Type: {pdf_type.title()} Slots', ln=True)
        pdf.cell(0, 8, f'Generated: {datetime.now().strftime("%d/%m/%Y at %H:%M:%S")}', ln=True)
        if data.get('auto_detect_info'):
            pdf.cell(0, 8, f'Period: {data["auto_detect_info"]}', ln=True)
        
        # Generate PDF bytes
        pdf_bytes = pdf.output(dest='S')
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode('latin1')
        
        return pdf_bytes
        
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

# First, get the checkboxes outside the form for immediate response
st.header("Input Parameters")

# PDF Output Options
st.subheader("PDF Output Options")
col1, col2, col3 = st.columns(3)
with col1:
    show_excess_only = st.checkbox("Show only slots with excess (loss)", value=True, key="show_excess_only")
with col2:
    show_all_slots = st.checkbox("Show all slots (15-min slot-wise)", key="show_all_slots")
with col3:
    show_daywise = st.checkbox("Show day-wise summary (all days in month)", key="show_daywise")

# Generation Sources Configuration (outside form for immediate response)
st.subheader("Generation Sources Configuration")
col1, col2 = st.columns(2)
with col1:
    enable_iex = st.checkbox("Enable I.E.X Generation", help="Check this to enable I.E.X generation data upload", key="enable_iex")
with col2:
    enable_cpp = st.checkbox("Enable C.P.P Generation", help="Check this to enable C.P.P generation data upload", key="enable_cpp")

# Main form for file uploads and other inputs
with st.form("energy_adjustment_form"):
    # Consumer Information
    st.subheader("Consumer Information")
    col1, col2 = st.columns(2)
    with col1:
        consumer_number = st.text_input("Consumer Number", help="Enter the consumer number")
    with col2:
        consumer_name = st.text_input("Consumer Name", help="Enter the consumer name")
    
    # I.E.X Generation (conditional based on checkbox outside form)
    generated_files = []
    t_and_d_loss = 0.0
    if enable_iex:
        st.subheader("I.E.X Generation Settings")
        generated_files = st.file_uploader(
            "Generated Energy Excel Files I.E.X (MW) From SLDC",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Select one or more Excel files containing I.E.X generation data",
            key="iex_files"
        )
        t_and_d_loss = st.number_input(
            "T&D Loss (%) for I.E.X",
            min_value=0.0,
            max_value=100.0,
            step=0.01,
            format="%.2f",
            help="Enter the transmission and distribution loss percentage for I.E.X",
            key="iex_td_loss",
            value=0.0
        )
    
    # C.P.P Generation (conditional based on checkbox outside form)
    cpp_files = []
    cpp_t_and_d_loss = 0.0
    if enable_cpp:
        st.subheader("C.P.P Generation Settings")
        cpp_files = st.file_uploader(
            "Generated Energy Excel Files C.P.P (MW) From SLDC",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Select one or more Excel files containing C.P.P generation data",
            key="cpp_files"
        )
        cpp_t_and_d_loss = st.number_input(
            "T&D Loss (%) for C.P.P",
            min_value=0.0,
            max_value=100.0,
            step=0.01,
            format="%.2f",
            help="Enter the transmission and distribution loss percentage for C.P.P",
            key="cpp_td_loss",
            value=0.0
        )
    
    # Consumed Energy
    st.subheader("Consumed Energy")
    consumed_files = st.file_uploader(
        "Consumed Energy Excel Files (kWh) From MRT",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Select one or more Excel files containing consumed energy data"
    )
    
    multiplication_factor = st.number_input(
        "Multiplication Factor (for Consumed Energy)",
        min_value=0.01,
        step=0.01,
        format="%.2f",
        help="Factor to multiply consumed energy values"
    )
    
    # Date Filters
    st.subheader("Date Filters")
    col1, col2, col3 = st.columns(3)
    with col1:
        date_filter = st.text_input(
            "Date (optional, dd/mm/yyyy)",
            placeholder="e.g. 10/10/2024",
            help="Filter data for a specific date"
        )
    with col2:
        month = st.selectbox(
            "Month (optional)",
            options=[""] + [str(i) for i in range(1, 13)],
            format_func=lambda x: "" if x == "" else datetime(2000, int(x), 1).strftime("%B")
        )
    with col3:
        year = st.number_input(
            "Year (optional)",
            min_value=2000,
            max_value=2100,
            value=None,
            help="Enter the year for filtering data"
        )
    
    auto_detect_month = st.checkbox(
        "Auto-detect month and year from files",
        value=True,
        help="Automatically detect the month and year from uploaded files"
    )
    
    # Debug information inside the form
    st.write("---")
    st.subheader("🔍 Current Configuration")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.write("**I.E.X Status:**")
        st.write(f"- Enabled: {enable_iex}")
        if enable_iex:
            st.write(f"- Files: {len(generated_files) if generated_files else 0}")
            st.write(f"- T&D Loss: {t_and_d_loss}%")
    
    with col2:
        st.write("**C.P.P Status:**")
        st.write(f"- Enabled: {enable_cpp}")
        if enable_cpp:
            st.write(f"- Files: {len(cpp_files) if cpp_files else 0}")
            st.write(f"- T&D Loss: {cpp_t_and_d_loss}%")
    
    with col3:
        st.write("**Consumption:**")
        st.write(f"- Files: {len(consumed_files) if consumed_files else 0}")
        st.write(f"- Factor: {multiplication_factor}")
    
    # Submit button
    submitted = st.form_submit_button("Generate PDF Report", type="primary")

# Remove debug information from outside the form

# Validation and processing
if submitted:
    # Clear previous errors
    st.session_state.error_message = None
    
    # Validation
    errors = []
    
    # Check PDF options
    if not (show_excess_only or show_all_slots or show_daywise):
        errors.append("Please select at least one PDF output option.")
    
    # Check consumer information
    if not consumer_number:
        errors.append("Consumer Number is required.")
    if not consumer_name:
        errors.append("Consumer Name is required.")
    
    # Check generation sources
    if not enable_iex and not enable_cpp:
        errors.append("Please enable at least one generation source (I.E.X or C.P.P).")
    
    # Check I.E.X validation
    if enable_iex:
        if not generated_files or len(generated_files) == 0:
            errors.append("Please select I.E.X generation files since I.E.X is enabled.")
        # Remove T&D loss validation to allow 0% loss
    
    # Check C.P.P validation
    if enable_cpp:
        if not cpp_files or len(cpp_files) == 0:
            errors.append("Please select C.P.P generation files since C.P.P is enabled.")
        # Remove T&D loss validation to allow 0% loss
    
    # Check consumed files
    if not consumed_files:
        errors.append("No consumed energy Excel files were uploaded.")
    
    # Check multiplication factor
    if multiplication_factor is None or multiplication_factor <= 0:
        errors.append("Multiplication factor must be greater than 0.")
    
    # Display errors if any
    if errors:
        st.session_state.error_message = "\n".join(errors)
    else:
        # Process the data
        st.success("Validation passed! Processing data...")
        
        try:
            # Process data with progress tracking
            with st.spinner("Processing uploaded files..."):
                result = process_energy_data(
                    generated_files, cpp_files, consumed_files,
                    enable_iex, enable_cpp, t_and_d_loss, cpp_t_and_d_loss,
                    consumer_number, consumer_name, multiplication_factor,
                    auto_detect_month, month, year, date_filter
                )
                
                if result['success']:
                    st.session_state.processed_data = result['data']
                    st.success("Data processed successfully!")
                else:
                    st.session_state.error_message = result['error']
            
        except Exception as e:
            st.session_state.error_message = f"Error processing data: {str(e)}"

# Display errors
if st.session_state.error_message:
    st.error(st.session_state.error_message)

# Display results if data has been processed
if st.session_state.processed_data:
    st.header("📊 Processing Results")
    
    data = st.session_state.processed_data
    
    # Display summary information
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Generated (After Loss)", f"{data['total_generated_after_loss']:.2f} kWh")
    with col2:
        st.metric("Total Consumed", f"{data['total_consumed']:.2f} kWh")
    with col3:
        st.metric("Total Excess", f"{data['total_excess']:.2f} kWh")
    with col4:
        st.metric("Status", data['excess_status'])
    
    # Additional details
    st.subheader("📈 Data Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info(f"**Unique Days (Generated):** {data['unique_days_gen']}")
    with col2:
        st.info(f"**Unique Days (Consumed):** {data['unique_days_cons']}")
    with col3:
        st.info(f"**Total Generated (Before Loss):** {data['sum_injection']:.2f} kWh")
    
    # Financial Calculations Display on Web Page
    st.subheader("💰 Financial Calculations")
    
    # Helper function for consistent rounding throughout the application
    def round_kwh_financial(value):
        return int(value + 0.5) if value >= 0 else int(value - 0.5)
    
    # Check if financial calculation data is available
    if 'total_excess_financial_rounded' in data:
        # Display financial calculations in organized columns
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Positive Charges:**")
            st.info(f"**Base Rate:** {data['total_excess_financial_rounded']} kWh × Rs.{data['base_rate']} = Rs.{data['base_amount']:.2f}")
            st.info(f"**C1+C2 Additional:** {data['c1_c2_excess']} kWh × Rs.1.8125 = Rs.{data['c1_c2_additional']:.2f}")
            st.info(f"**C5 Additional:** {data['c5_excess']} kWh × Rs.0.3625 = Rs.{data['c5_additional']:.2f}")
            st.info(f"**Subtotal:** Rs.{data['total_amount']:.2f}")
            st.info(f"**E-Tax (5%):** Rs.{data['etax']:.2f}")
            st.success(f"**Total with E-Tax:** Rs.{data['total_with_etax']:.2f}")
        
        with col2:
            st.write("**Negative Charges (Deductions):**")
            st.warning(f"**E-Tax on IEX:** Rs.{data['etax_on_iex']:.2f}")
            st.warning(f"**Cross Subsidy Surcharge:** {data['iex_excess_financial']} kWh × Rs.1.92 = Rs.{data['cross_subsidy_surcharge']:.2f}")
            st.warning(f"**Wheeling Charges:** Rs.{data['wheeling_charges']:.2f}")
            st.warning(f"**Total Deductions:** Rs.{data['etax_on_iex'] + data['cross_subsidy_surcharge'] + data['wheeling_charges']:.2f}")
            st.success(f"**Final Amount:** Rs.{data['final_amount']:.2f}")
            st.success(f"**Final Amount (Rounded Up):** Rs.{data['final_amount_rounded']}")
    else:
        # Fallback to calculating on the fly if the pre-calculated values aren't available
        # Calculate financial values using rounded values for consistency
        total_excess_financial_rounded = round_kwh_financial(data['total_excess'])
        
        # Base rate for all excess energy using rounded values
        base_rate = 7.25  # rupees per kWh
        base_amount = total_excess_financial_rounded * base_rate
        
        # Calculate TOD-wise excess for financial calculations
        merged_data = data.get('merged_all', pd.DataFrame())
        if not merged_data.empty:
            tod_excess = merged_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
            
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
            iex_excess_financial_raw = merged_data['IEX_Excess'].sum() if 'IEX_Excess' in merged_data.columns else data['total_excess']
            iex_excess_financial = round_kwh_financial(iex_excess_financial_raw)
            
            # Calculate negative factors using rounded values
            etax_on_iex = total_excess_financial_rounded * 0.1
            cross_subsidy_surcharge = iex_excess_financial * 1.92  # Only for IEX excess
            
            # Use the complex wheeling charges formula from Flask app
            loss_percentage = data.get('t_and_d_loss', 0)
            # Ensure loss_percentage is not 100 to avoid division by zero
            if loss_percentage < 100:
                x_value = (total_excess_financial_rounded * loss_percentage) / (100 - loss_percentage)
                # Ensure x_value is not zero to avoid division by zero
                if x_value != 0:
                    y_value = (total_excess_financial_rounded * loss_percentage / x_value) + total_excess_financial_rounded
                else:
                    y_value = total_excess_financial_rounded
            else:
                x_value = 0
                y_value = total_excess_financial_rounded
            wheeling_charges = y_value * 1.04
            
            # Calculate final amount to be collected
            final_amount = total_with_etax - (etax_on_iex + cross_subsidy_surcharge + wheeling_charges)
            
            # Round up final amount to next highest value
            final_amount_rounded = math.ceil(final_amount)
            
            # Display financial calculations in organized columns
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Positive Charges:**")
                st.info(f"**Base Rate:** {total_excess_financial_rounded} kWh × Rs.{base_rate} = Rs.{base_amount:.2f}")
                st.info(f"**C1+C2 Additional:** {c1_c2_excess} kWh × Rs.1.8125 = Rs.{c1_c2_additional:.2f}")
                st.info(f"**C5 Additional:** {c5_excess} kWh × Rs.0.3625 = Rs.{c5_additional:.2f}")
                st.info(f"**Subtotal:** Rs.{total_amount:.2f}")
                st.info(f"**E-Tax (5%):** Rs.{etax:.2f}")
                st.success(f"**Total with E-Tax:** Rs.{total_with_etax:.2f}")
            
            with col2:
                st.write("**Negative Charges (Deductions):**")
                st.warning(f"**E-Tax on IEX:** Rs.{etax_on_iex:.2f}")
                st.warning(f"**Cross Subsidy Surcharge:** {iex_excess_financial} kWh × Rs.1.92 = Rs.{cross_subsidy_surcharge:.2f}")
                st.warning(f"**Wheeling Charges:** Rs.{wheeling_charges:.2f}")
                st.warning(f"**Total Deductions:** Rs.{etax_on_iex + cross_subsidy_surcharge + wheeling_charges:.2f}")
                st.success(f"**Final Amount:** Rs.{final_amount:.2f}")
                st.success(f"**Final Amount (Rounded Up):** Rs.{final_amount_rounded}")
        else:
            st.warning("No merged data available for financial calculations.")
    
    # TOD-wise breakdown
    st.subheader("⏰ TOD-wise Excess Energy Breakdown")
    merged_data = data.get('merged_all', pd.DataFrame())
    if not merged_data.empty:
        tod_excess = merged_data.groupby('TOD_Category')['Total_Excess'].sum().reset_index()
        tod_display = []
        for _, row in tod_excess.iterrows():
            category = row['TOD_Category']
            excess_rounded = round_kwh_financial(row['Total_Excess'])
            tod_display.append({"TOD Category": category, "Excess Energy (kWh)": excess_rounded})
        
        if tod_display:
            tod_df = pd.DataFrame(tod_display)
            st.dataframe(tod_df, use_container_width=True)
    else:
        st.warning("No TOD data available for breakdown.")
    
    # Generate PDFs based on selected options from top
    st.header("📄 Generating PDF Reports")
    st.info("Processing your data and generating PDF reports based on your selections...")
    
    with st.spinner("Generating PDF reports..."):
        # Generate PDFs based on checkbox selections
        pdfs_generated = []
        
        if show_excess_only:
            if not data['merged_excess'].empty:
                pdf_bytes = generate_detailed_pdf(data, data['merged_excess'], "excess")
                if pdf_bytes:
                    filename = generate_custom_filename("excess_only", data['consumer_number'], data['consumer_name'], data.get('month'), data.get('year'))
                    pdfs_generated.append((filename, pdf_bytes))
        
        if show_all_slots:
            pdf_bytes = generate_detailed_pdf(data, data['merged_all'], "all_slots")
            if pdf_bytes:
                filename = generate_custom_filename("all_slots", data['consumer_number'], data['consumer_name'], data.get('month'), data.get('year'))
                pdfs_generated.append((filename, pdf_bytes))
        
        if show_daywise:
            pdf_bytes = generate_daywise_pdf(data, data['merged_all'])
            if pdf_bytes:
                filename = generate_custom_filename("daywise", data['consumer_number'], data['consumer_name'], data.get('month'), data.get('year'))
                pdfs_generated.append((filename, pdf_bytes))
    
    # Display download buttons for generated PDFs
    if pdfs_generated:
        st.success(f"✅ Successfully generated {len(pdfs_generated)} PDF report(s)!")
        
        if len(pdfs_generated) == 1:
            # Single PDF download
            filename, pdf_bytes = pdfs_generated[0]
            st.download_button(
                label=f"📥 Download {filename}",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                type="primary"
            )
        else:
            # Multiple PDFs - create a ZIP file
            import zipfile
            zip_buffer = io.BytesIO()
            last_3_digits = str(data['consumer_number'])[-3:] if len(str(data['consumer_number'])) >= 3 else str(data['consumer_number'])
            clean_name = "".join(c for c in data['consumer_name'] if c.isalnum() or c in (' ', '-', '_')).strip().replace(' ', '_')
            zip_filename = f"{last_3_digits}_{clean_name}_energy_adjustment_reports.zip"
            
            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                for fname, pdf_bytes in pdfs_generated:
                    zf.writestr(fname, pdf_bytes)
            zip_buffer.seek(0)
            
            st.download_button(
                label=f"� Download All Reports (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=zip_filename,
                mime="application/zip",
                type="primary"
            )
    else:
        st.error("❌ No PDF reports were generated. Please check your selections and try again.")

# Display errors
if st.session_state.error_message:
    st.error(st.session_state.error_message)

# Auto-Update Sidebar
if UPDATER_AVAILABLE:
    with st.sidebar:
        st.header("🔄 Auto-Update System")
        
        # Get current version
        version_file = Path(__file__).parent / "version.json"
        current_version = "1.0.0"
        if version_file.exists():
            try:
                with open(version_file, 'r') as f:
                    version_data = json.load(f)
                    current_version = version_data.get("version", "1.0.0")
            except:
                pass
        
        st.info(f"**Current Version:** {current_version}")
        
        # Initialize updater
        if 'updater' not in st.session_state:
            st.session_state.updater = initialize_updater(current_version)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔍 Check Updates", help="Check for available updates"):
                with st.spinner("Checking for updates..."):
                    update_info = st.session_state.updater.check_for_updates(show_no_updates=False)
                    if update_info:
                        st.success(f"Update available: v{update_info['version']}")
                        st.info(f"**What's New:**\n{update_info['description'][:200]}...")
                        
                        if st.button("📥 Download Update", key="download_update"):
                            with st.spinner("Downloading and applying update..."):
                                try:
                                    # Download update
                                    update_file = st.session_state.updater.download_update(update_info)
                                    if update_file:
                                        # Apply update
                                        if st.session_state.updater.apply_update(update_file, update_info):
                                            st.success("✅ Update applied successfully! Please restart the application.")
                                            st.balloons()
                                        else:
                                            st.error("❌ Failed to apply update.")
                                    else:
                                        st.error("❌ Failed to download update.")
                                except Exception as e:
                                    st.error(f"❌ Update failed: {e}")
                    else:
                        st.success("✅ You have the latest version!")
        
        with col2:
            if st.button("⚙️ Settings", help="Update settings"):
                show_update_settings(st.session_state.updater)
        
        # Auto-update status
        if st.session_state.updater.config.get("auto_check", True):
            st.success("🟢 Auto-updates enabled")
        else:
            st.warning("🟡 Auto-updates disabled")
        
        # Last check info
        last_check = st.session_state.updater.config.get("last_check")
        if last_check:
            try:
                last_check_time = datetime.fromisoformat(last_check)
                st.caption(f"Last checked: {last_check_time.strftime('%d/%m/%Y %H:%M')}")
            except:
                pass

# Footer
st.markdown("---")
st.markdown("Energy Adjustment Calculator - Streamlit Version")
if UPDATER_AVAILABLE:
    st.markdown("🔄 Auto-update system enabled")
