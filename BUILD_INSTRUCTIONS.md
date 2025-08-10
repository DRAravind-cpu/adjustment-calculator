# Building Windows Offline Application

## Overview
This guide explains how to create a standalone Windows executable for the Energy Adjustment Calculator that works offline on Windows 7-11.

## Prerequisites

### On Windows (Recommended)
1. **Python 3.8-3.11** (3.12+ may have compatibility issues with PyInstaller)
2. **Git** (to clone the repository)
3. **NSIS** (optional, for creating installer)

### Installation Steps
```bash
# Clone the repository
git clone https://github.com/DRAravind-cpu/adjustment-calculator.git
cd adjustment-calculator

# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements_windows.txt
```

## Building Methods

### Method 1: Simple Build (Recommended)
```bash
# On Windows, simply run:
build_windows.bat

# This will:
# - Install PyInstaller
# - Clean previous builds
# - Create the executable
# - Show success/failure message
```

### Method 2: Manual Build
```bash
# Install PyInstaller
pip install pyinstaller>=5.0

# Build using spec file
pyinstaller app.spec

# Or build with command line options
pyinstaller --onefile --windowed --name="EnergyAdjustmentCalculator" launcher.py
```

### Method 3: Cross-Platform Build (Advanced)
```bash
# On Mac/Linux, you can prepare the build files
python build_windows_app.py

# Then transfer to Windows machine and run:
# build_windows.bat
```

## Creating an Installer (Optional)

### Using NSIS
1. Install NSIS from https://nsis.sourceforge.io/
2. After building the executable:
```bash
# Compile the installer
makensis installer.nsi
```

### Using Auto-py-to-exe (GUI Method)
```bash
# Install GUI tool
pip install auto-py-to-exe

# Launch GUI
auto-py-to-exe

# Configure settings:
# - Script Location: launcher.py
# - Onefile: One File
# - Console Window: Console Based
# - Additional Files: streamlit_app.py, templates/, .streamlit/
```

## Build Output

### Successful Build Creates:
- `dist/EnergyAdjustmentCalculator.exe` - Main executable
- `dist/` folder with all dependencies
- Optional: `EnergyAdjustmentCalculator_Setup.exe` - Installer

### Distribution
- **For simple distribution**: Share the entire `dist/` folder
- **For professional distribution**: Use the installer `.exe`

## Testing the Build

### On Windows 7-11:
1. Copy the `dist/` folder to target machine
2. Double-click `EnergyAdjustmentCalculator.exe`
3. Wait 30-60 seconds for startup
4. Browser should open automatically
5. Test all features: file upload, processing, PDF generation

## Troubleshooting

### Common Issues:

**Build Fails:**
- Ensure Python 3.8-3.11 (not 3.12+)
- Install Visual C++ Redistributable
- Try: `pip install --upgrade pyinstaller`

**Executable Won't Start:**
- Check Windows Defender/Antivirus
- Run as Administrator
- Check console output for errors

**Missing Dependencies:**
- Add to `hiddenimports` in `app.spec`
- Use `--collect-all package_name` flag

**Large File Size:**
- Use `--exclude-module` for unused packages
- Enable UPX compression: `upx=True`

### Performance Optimization:
```python
# In app.spec, add exclusions:
excludes=[
    'matplotlib',
    'scipy',
    'sklearn',
    'tensorflow',
    'torch'
]
```

## File Structure After Build
```
dist/
├── EnergyAdjustmentCalculator.exe    # Main executable
├── streamlit_app.py                  # Streamlit app
├── templates/                        # HTML templates
├── .streamlit/                       # Streamlit config
└── _internal/                        # Python runtime & dependencies
```

## Advanced Configuration

### Custom Icon
1. Create `app_icon.ico` (256x256 recommended)
2. Update `app.spec`: `icon='app_icon.ico'`

### Startup Optimization
- Modify `launcher.py` to add splash screen
- Pre-compile Python files
- Use `--optimize=2` flag

### Security Considerations
- Code signing certificate (for enterprise distribution)
- Antivirus whitelisting instructions
- User Account Control (UAC) handling

## Support
- Check `WINDOWS_APP_GUIDE.md` for user instructions
- Report issues on GitHub repository
- Test on multiple Windows versions before distribution