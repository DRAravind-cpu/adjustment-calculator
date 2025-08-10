# Energy Adjustment Calculator - Windows Offline Version

## For End Users

### System Requirements
- Windows 7, 8, 10, or 11 (32-bit or 64-bit)
- No additional software installation required
- Minimum 4GB RAM recommended
- 500MB free disk space

### How to Use
1. Extract the application folder to your desired location
2. Double-click `EnergyAdjustmentCalculator.exe`
3. Wait for the application to start (may take 30-60 seconds on first run)
4. Your web browser will automatically open with the application
5. Use the application as normal - upload Excel files, generate reports
6. To close the application, close the console window

### Features
- Complete offline functionality
- All original features included:
  - Excel file upload and processing
  - T&D loss calculations
  - TOD categorization
  - PDF report generation
  - Multiple output options
- No internet connection required
- No Python installation needed

### Troubleshooting
- If the browser doesn't open automatically, manually go to: http://localhost:8501
- If you get a Windows security warning, click "More info" then "Run anyway"
- Antivirus software may flag the executable - this is normal for PyInstaller apps
- For support, contact the application developer

## For Developers

### Building the Application
1. Ensure you have Python 3.8+ installed
2. Install dependencies: `pip install -r requirements.txt`
3. Run the build script:
   - On Windows: `build_windows.bat`
   - On Mac/Linux: `python build_windows_app.py`

### Distribution
- The entire `dist` folder needs to be distributed
- Users only need to run the .exe file
- Consider creating an installer using tools like NSIS or Inno Setup

### Customization
- Modify `launcher.py` to change startup behavior
- Edit `app.spec` to include additional files or change build settings
- Update the application icon by replacing `app_icon.ico`
