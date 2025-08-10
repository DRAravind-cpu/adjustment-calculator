#!/usr/bin/env python3
"""
Windows Application Builder for Energy Adjustment Calculator
Creates a standalone Windows executable using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_build_dependencies():
    """Install required packages for building Windows executable"""
    packages = [
        'pyinstaller>=5.0',
        'auto-py-to-exe',  # Optional GUI for PyInstaller
    ]
    
    for package in packages:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

def create_launcher_script():
    """Create a launcher script that starts Streamlit"""
    launcher_content = '''
import os
import sys
import subprocess
import webbrowser
import time
import socket
from pathlib import Path

def find_free_port():
    """Find a free port for Streamlit"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def main():
    # Get the directory where the executable is located
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        app_dir = Path(sys.executable).parent
    else:
        # Running as script
        app_dir = Path(__file__).parent
    
    # Change to app directory
    os.chdir(app_dir)
    
    # Find free port
    port = find_free_port()
    
    print("=" * 60)
    print("Energy Adjustment Calculator - Offline Version")
    print("=" * 60)
    print(f"Starting application on port {port}...")
    print("The application will open in your default web browser.")
    print("To stop the application, close this window.")
    print("=" * 60)
    
    # Start Streamlit
    try:
        # Wait a moment then open browser
        import threading
        def open_browser():
            time.sleep(3)
            webbrowser.open(f'http://localhost:{port}')
        
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()
        
        # Start Streamlit
        subprocess.run([
            sys.executable, '-m', 'streamlit', 'run', 
            'streamlit_app.py',
            '--server.port', str(port),
            '--server.headless', 'true',
            '--browser.gatherUsageStats', 'false',
            '--server.address', 'localhost'
        ])
    except KeyboardInterrupt:
        print("\\nApplication stopped by user.")
    except Exception as e:
        print(f"Error starting application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
'''
    
    with open('/Users/admin/Documents/adjustment-calculator/launcher.py', 'w') as f:
        f.write(launcher_content)
    
    print("Created launcher.py")

def create_pyinstaller_spec():
    """Create PyInstaller spec file for better control"""
    spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('streamlit_app.py', '.'),
        ('templates', 'templates'),
        ('.streamlit', '.streamlit'),
    ],
    hiddenimports=[
        'streamlit',
        'pandas',
        'openpyxl',
        'fpdf',
        'numpy',
        'altair',
        'plotly',
        'PIL',
        'xlsxwriter',
        'streamlit.web.cli',
        'streamlit.runtime.scriptrunner.magic_funcs',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EnergyAdjustmentCalculator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
)
'''
    
    with open('/Users/admin/Documents/adjustment-calculator/app.spec', 'w') as f:
        f.write(spec_content)
    
    print("Created app.spec")

def create_build_script():
    """Create batch script for building on Windows"""
    batch_content = '''@echo off
echo Building Energy Adjustment Calculator for Windows...
echo.

REM Install build dependencies
echo Installing PyInstaller...
pip install pyinstaller>=5.0

REM Clean previous builds
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build the application
echo Building executable...
pyinstaller app.spec

REM Check if build was successful
if exist "dist\\EnergyAdjustmentCalculator.exe" (
    echo.
    echo ========================================
    echo Build completed successfully!
    echo ========================================
    echo Executable location: dist\\EnergyAdjustmentCalculator.exe
    echo.
    echo You can now distribute the 'dist' folder to users.
    echo Users just need to run EnergyAdjustmentCalculator.exe
    echo.
) else (
    echo.
    echo ========================================
    echo Build failed!
    echo ========================================
    echo Please check the error messages above.
)

pause
'''
    
    with open('/Users/admin/Documents/adjustment-calculator/build_windows.bat', 'w') as f:
        f.write(batch_content)
    
    print("Created build_windows.bat")

def create_user_guide():
    """Create user guide for the Windows application"""
    guide_content = '''# Energy Adjustment Calculator - Windows Offline Version

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
'''
    
    with open('/Users/admin/Documents/adjustment-calculator/WINDOWS_APP_GUIDE.md', 'w') as f:
        f.write(guide_content)
    
    print("Created WINDOWS_APP_GUIDE.md")

def main():
    """Main build process"""
    print("Setting up Windows Application Build Environment...")
    print("=" * 60)
    
    # Create necessary files
    create_launcher_script()
    create_pyinstaller_spec()
    create_build_script()
    create_user_guide()
    
    print("\n" + "=" * 60)
    print("Setup Complete!")
    print("=" * 60)
    print("\nNext steps:")
    print("1. On a Windows machine, run: build_windows.bat")
    print("2. Or manually run: pyinstaller app.spec")
    print("3. The executable will be created in the 'dist' folder")
    print("4. Distribute the entire 'dist' folder to users")
    print("\nSee WINDOWS_APP_GUIDE.md for detailed instructions.")

if __name__ == "__main__":
    main()