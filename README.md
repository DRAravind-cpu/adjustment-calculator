# Energy Adjustment Calculator

A comprehensive web application for calculating energy adjustments between generated and consumed energy, with support for multiple energy sources (I.E.X and C.P.P), T&D loss calculations, and professional PDF report generation.

## Features

- **Multi-Source Energy Support**: Process both I.E.X and C.P.P energy sources with separate T&D loss percentages
- **Sequential Adjustment Logic**: Applies I.E.X energy first, then C.P.P energy for remaining consumption
- **Multiple File Upload**: Support for multiple Excel files for each energy source
- **Flexible Date Filtering**: Filter by month/year or specific date with auto-detection capability
- **Customizable PDF Reports**: Generate three types of reports (Excess Only, All Slots, Day-wise Summary)
- **TOD Category Classification**: Categorizes time slots into TOD categories (C1, C2, C4, C5)
- **Financial Calculations**: Computes charges based on excess energy and TOD categories
- **Multiple Interfaces**: Available as both Flask web application and Streamlit application

## Implementation Options

### Flask Web Application (app.py)
- Traditional web interface with form-based input
- Suitable for server deployment
- Includes checkbox options for enabling I.E.X and C.P.P sources

### Streamlit Application (streamlit_app.py)
- Modern, interactive interface with real-time feedback
- Easier to use with drag-and-drop file uploads
- Enhanced data visualization capabilities
- Automatic updates when app.py changes (via sync utility)

## Data Processing Flow

1. **File Upload**: Accept Excel files for I.E.X generation, C.P.P generation, and consumption
2. **Data Extraction**: Read Excel files and extract date, time, and energy values
3. **Source Tracking**: Tag data with source type (I.E.X or C.P.P)
4. **Data Filtering**: Filter by month/year or specific date
5. **Energy Conversion**: Convert MW to kWh (1 MW = 250 kWh)
6. **T&D Loss Application**: Apply separate T&D loss percentages for I.E.X and C.P.P
7. **Sequential Adjustment**:
   - First apply I.E.X energy to consumption
   - Use C.P.P energy for remaining consumption
   - Calculate excess energy for each source
8. **TOD Classification**: Categorize time slots into TOD categories
9. **Financial Calculations**: Compute charges based on excess energy and TOD categories
10. **PDF Generation**: Create customized PDF reports

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/adjustment-calculator.git
cd adjustment-calculator

# Install dependencies
pip install -r requirements.txt

# Run Flask application
python app.py

# OR run Streamlit application
streamlit run streamlit_app.py
```

## Usage

### Required Inputs
- **Generated Energy Files**: Excel files with Date, Time, and Energy (MW) columns
- **Consumed Energy Files**: Excel files with Date, Time, and Energy (kWh) columns
- **Consumer Information**: Number and name
- **T&D Loss Percentages**: Separate values for I.E.X and C.P.P sources
- **Multiplication Factor**: For consumed energy adjustment (default: 1)
- **Date Range**: Month and year or enable auto-detection

### Output Options
- **Excess Only Report**: Shows only time slots with excess energy (ideal for billing)
- **All Slots Report**: Shows all 96 time slots per day (15-minute intervals)
- **Day-wise Summary**: Shows daily aggregated totals for the entire month

## File Format Requirements

- Excel files (.xlsx, .xls) with at least 3 columns:
  - Date (dd/mm/yyyy format recommended)
  - Time (HH:MM format or time range format "HH:MM - HH:MM")
  - Energy (numeric values: MW for generation, kWh for consumption)

## Recent Updates

- Added multi-source support with I.E.X and C.P.P energy sources
- Implemented sequential adjustment logic for energy calculation
- Added separate T&D loss percentages for different energy sources
- Enhanced PDF reports with source-specific breakdowns
- Added Streamlit interface as an alternative to Flask
- Improved validation and error handling
- Enhanced TOD category classification
- Removed time ranges in parentheses from TOD Category descriptions to improve PDF readability
