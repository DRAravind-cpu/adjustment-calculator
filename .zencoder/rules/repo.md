# Energy Adjustment Calculator Information

## Summary
This project is a Flask-based web application that calculates energy adjustments for power consumers. It allows users to upload Excel files for generated energy (in MW) and consumed energy (in kWh), enter T&D loss percentages, and generate PDF reports comparing energy generation and consumption with appropriate adjustments.

## Structure
- **app.py**: Main Flask application file
- **streamlit_app.py**: Alternative Streamlit implementation
- **templates/**: Contains HTML templates for the Flask web interface
- **requirements.txt**: Python dependencies
- **package.json**: Node.js package configuration

## Language & Runtime
**Language**: Python
**Version**: Python 3.x
**Web Frameworks**: Flask, Streamlit
**Package Manager**: pip

## Dependencies
**Main Dependencies**:
- Flask: Web framework
- pandas: Data manipulation and analysis
- openpyxl: Excel file handling
- fpdf: PDF generation
- streamlit: Alternative web interface

## Build & Installation
```bash
# Install dependencies
pip install -r requirements.txt

# Run Flask application
python app.py

# Alternative: Run Streamlit application
streamlit run streamlit_app.py
```

## Data Processing Flow
1. **File Upload**: Application accepts multiple Excel files for generated energy (I.E.X and C.P.P) and consumed energy
2. **Data Extraction**: Reads Excel files and extracts date, time, and energy values
3. **Data Filtering**: Filters data by month/year or specific date if provided
4. **Energy Conversion**: Converts MW to kWh (1 MW = 250 kWh)
5. **T&D Loss Application**: Applies separate T&D loss percentages for I.E.X and C.P.P generation
6. **Data Merging**: Combines generated and consumed energy data by date and time slots
7. **Excess Calculation**: Calculates excess energy (generated after loss - consumed)
8. **TOD Classification**: Categorizes time slots into TOD categories (C1, C2, C4, C5)
9. **Financial Calculations**: Computes charges based on excess energy and TOD categories
10. **PDF Generation**: Creates PDF reports based on selected output options

## Known Issues
- Discrepancy between calculated values in terminal output and values passed to PDF generation
- Pandas SettingWithCopyWarning in data processing pipeline
- Non-numeric values in energy data causing NaN conversions
- Inconsistent handling of date/time formats across different input files

## Usage
The application provides two different interfaces:

### Flask Web Interface
1. Navigate to the application URL in a browser
2. Enter consumer information (number and name)
3. Upload generated energy Excel files (MW) from SLDC
4. Upload consumed energy Excel files (kWh) from MRT
5. Enter T&D loss percentage and multiplication factor
6. Select PDF output options
7. Generate and download the PDF report

### Streamlit Interface
1. Launch the Streamlit app
2. Upload required files through the Streamlit UI
3. Fill in the parameters
4. The application will process the data and generate a downloadable PDF report

## Features
- Support for multiple input files
- Automatic month/year detection from uploaded files
- Configurable T&D loss percentages
- Multiple PDF output options (show excess only, all slots, day-wise summary)
- Support for both I.E.X and C.P.P generation sources
- Energy conversion from MW to kWh with T&D loss adjustment
- TOD (Time of Day) category classification
- Financial calculations including base rate, additional charges, and taxes