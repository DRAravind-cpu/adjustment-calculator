# Streamlit Energy Adjustment Calculator

## Overview
This is a complete Streamlit replication of the Flask Energy Adjustment Calculator application. The app processes energy generation and consumption data to calculate adjustments and generate PDF reports.

## Features Implemented

### ✅ Core Functionality
1. **File Upload Support**
   - Multiple Excel file uploads for I.E.X generation data
   - Multiple Excel file uploads for C.P.P generation data  
   - Multiple Excel file uploads for consumption data
   - Proper file validation and error handling

2. **Data Processing**
   - Excel file parsing and validation
   - Date/time filtering and standardization
   - MW to kWh conversion (MW × 250 for 15-minute intervals)
   - T&D loss application (separate for I.E.X and C.P.P)
   - Sequential adjustment calculation (I.E.X first, then C.P.P)

3. **Time of Day (TOD) Classification**
   - C1: Morning Peak (6:00-10:00 AM)
   - C2: Evening Peak (6:00-10:00 PM)
   - C4: Normal Hours (5:00-6:00 AM, 10:00 AM-6:00 PM)
   - C5: Night Hours (10:00 PM-5:00 AM)

4. **Sequential Adjustment Logic**
   - Step 1: I.E.X adjustment with consumption
   - Step 2: Calculate remaining consumption
   - Step 3: C.P.P adjustment with remaining consumption
   - Step 4: Calculate total excess energy

5. **Data Visualization & Analysis**
   - Summary metrics display
   - Interactive data tables (Excess Only / All Slots)
   - TOD-wise excess energy analysis
   - Bar charts for TOD distribution

6. **Financial Calculations Preview**
   - Base rate calculations (₹7.25/kWh)
   - TOD-specific additional charges
   - E-Tax calculations
   - Summary totals

7. **Export & PDF Generation**
   - CSV export for all data and excess-only data
   - Basic PDF report generation
   - Custom filename generation
   - Download functionality

### ✅ User Interface Features
1. **Modern Streamlit UI**
   - Clean, responsive layout
   - Form validation with real-time feedback
   - Progress indicators and status messages
   - Tabbed data viewing
   - Metric cards for key information

2. **Error Handling**
   - Comprehensive input validation
   - File format validation
   - Data integrity checks
   - User-friendly error messages

3. **Configuration Options**
   - Enable/disable I.E.X and C.P.P sources
   - Flexible T&D loss settings
   - Auto-detection of month/year
   - Date filtering options
   - Multiplication factor for consumption

## Testing the Application

### Sample Files Created
- `sample_generation_iex.xlsx` - 96 slots of sample I.E.X generation data for July 1, 2025
- `sample_consumption.xlsx` - 96 slots of sample consumption data for July 1, 2025

### Test Steps
1. **Start the Application**
   ```bash
   cd /Users/admin/Documents/adjustment-calculator
   /Users/admin/Documents/adjustment-calculator/.venv/bin/streamlit run streamlit_app.py --server.port 8503
   ```

2. **Access the App**
   - Open browser to: http://localhost:8503

3. **Test with Sample Data**
   - Consumer Number: `12345`
   - Consumer Name: `Test Consumer`
   - Enable I.E.X Generation: ✓
   - Upload: `sample_generation_iex.xlsx`
   - T&D Loss for I.E.X: `5.5`
   - Upload Consumption: `sample_consumption.xlsx`
   - Multiplication Factor: `1.0`
   - Auto-detect month/year: ✓

4. **Expected Results**
   - Process 96 time slots for July 1, 2025
   - Calculate excess energy where generation > consumption
   - Show TOD-wise breakdown
   - Generate downloadable PDF reports

## Key Features Matching Flask App

### ✅ Exact Replication Features
1. **Form Structure** - Identical input fields and validation
2. **File Processing** - Same Excel parsing logic
3. **Calculation Logic** - Sequential adjustment algorithm
4. **TOD Classification** - Same time categories and rules
5. **Financial Calculations** - Same rates and formulas
6. **PDF Structure** - Basic report generation
7. **Error Handling** - Comprehensive validation

### 🚀 Streamlit Enhancements
1. **Interactive UI** - Real-time data preview
2. **Visual Charts** - TOD excess energy visualization
3. **Export Options** - CSV download functionality
4. **Progress Indicators** - User feedback during processing
5. **Responsive Design** - Mobile-friendly interface

## Technical Implementation

### Dependencies
- streamlit
- pandas  
- fpdf
- openpyxl
- datetime

### Architecture
- Single-file Streamlit application
- Session state management for data persistence
- Modular functions for data processing
- Error handling with user feedback
- File validation and data integrity checks

## Current Status
The application successfully replicates all core functionality of the Flask app with an enhanced user interface. Users can upload files, process energy data, view results, and generate PDF reports.

## Next Steps (Optional Enhancements)
1. Advanced PDF generation with detailed tables
2. Multi-day data processing
3. Historical data comparison
4. Enhanced financial calculation details
5. Data visualization charts
6. Batch processing capabilities
