# Energy Adjustment Calculator - Windows Offline Version

## 🎯 Overview
This is a complete offline Windows application version of the Energy Adjustment Calculator. It works on Windows 7, 8, 10, and 11 without requiring internet connection or Python installation.

## ✨ Features
- **Complete Offline Functionality** - No internet required
- **Universal Windows Compatibility** - Works on Windows 7-11
- **No Installation Dependencies** - Python and all libraries bundled
- **User-Friendly Interface** - Same Streamlit interface as web version
- **All Original Features** - Excel upload, T&D calculations, PDF generation
- **Professional Distribution** - Optional installer for easy deployment

## 🚀 Quick Start

### For End Users
1. Download the application package
2. Extract to desired location
3. Run `EnergyAdjustmentCalculator.exe`
4. Wait for browser to open automatically
5. Use normally - upload files, generate reports

### For Developers/Distributors
1. Follow build instructions in `BUILD_INSTRUCTIONS.md`
2. Test on target Windows versions
3. Distribute `dist/` folder or use installer

## 📁 What's Included

### Build Files
- `launcher.py` - Application launcher script
- `app.spec` - PyInstaller configuration
- `build_windows.bat` - Windows build script
- `build_windows_app.py` - Cross-platform build setup

### Distribution Files
- `installer.nsi` - NSIS installer script
- `LICENSE.txt` - License agreement
- `requirements_windows.txt` - Windows-specific dependencies

### Documentation
- `BUILD_INSTRUCTIONS.md` - Detailed build guide
- `WINDOWS_APP_GUIDE.md` - User and developer guide
- `WINDOWS_OFFLINE_README.md` - This file

## 🔧 Building the Application

### Simple Method (Windows)
```bash
# Clone repository
git clone https://github.com/DRAravind-cpu/adjustment-calculator.git
cd adjustment-calculator

# Run build script
build_windows.bat
```

### Manual Method
```bash
# Setup environment
python -m venv venv
venv\Scripts\activate
pip install -r requirements_windows.txt

# Build executable
pyinstaller app.spec
```

## 📦 Distribution Options

### Option 1: Folder Distribution
- Share entire `dist/` folder
- Users run `EnergyAdjustmentCalculator.exe`
- Simple but larger download

### Option 2: Professional Installer
- Use NSIS to create installer
- Professional installation experience
- Adds to Programs list, creates shortcuts
- Smaller download, better user experience

## 🎨 Customization

### Application Icon
- Replace `app_icon.ico` with custom icon
- Update `app.spec` to reference new icon

### Startup Behavior
- Modify `launcher.py` for custom startup
- Add splash screen or loading messages
- Change default port or browser behavior

### Build Optimization
- Exclude unused modules in `app.spec`
- Enable UPX compression for smaller size
- Add/remove hidden imports as needed

## 🔍 Technical Details

### System Requirements
- **OS**: Windows 7/8/10/11 (32-bit or 64-bit)
- **RAM**: 4GB minimum, 8GB recommended
- **Storage**: 500MB free space
- **Browser**: Any modern browser (Chrome, Firefox, Edge)

### Architecture
- **Frontend**: Streamlit web interface
- **Backend**: Python Flask-like processing
- **Packaging**: PyInstaller with all dependencies
- **Runtime**: Embedded Python 3.9+ runtime

### Security Considerations
- Application may trigger antivirus warnings (normal for PyInstaller)
- Consider code signing for enterprise distribution
- No network access required (fully offline)

## 🐛 Troubleshooting

### Common Issues

**Application Won't Start**
- Check Windows Defender exclusions
- Run as Administrator
- Verify all files in `dist/` folder

**Browser Doesn't Open**
- Manually navigate to `http://localhost:8501`
- Check firewall settings
- Try different browser

**Slow Startup**
- First run takes longer (30-60 seconds)
- Subsequent runs are faster
- Consider SSD for better performance

**File Upload Issues**
- Check file permissions
- Ensure Excel files are not corrupted
- Verify file size limits (200MB max)

### Performance Tips
- Close other applications for better performance
- Use SSD storage for faster file operations
- Ensure adequate RAM (8GB+ recommended)

## 📊 Comparison with Web Version

| Feature | Web Version | Windows Offline |
|---------|-------------|-----------------|
| Internet Required | Yes | No |
| Installation | Python + Dependencies | Single Executable |
| Updates | Git pull | New executable |
| Performance | Server dependent | Local machine |
| Security | Network dependent | Fully offline |
| Distribution | Code sharing | Executable sharing |

## 🔄 Updates and Maintenance

### Updating the Application
1. Rebuild with latest code
2. Test on target systems
3. Distribute new executable
4. Optional: Create update installer

### Version Management
- Update version in `app.spec`
- Maintain changelog
- Test compatibility with new Windows versions

## 📞 Support

### For Users
- Check `WINDOWS_APP_GUIDE.md` for detailed instructions
- Report issues on GitHub repository
- Contact application developer for support

### For Developers
- See `BUILD_INSTRUCTIONS.md` for build details
- Check PyInstaller documentation for advanced options
- Test on multiple Windows versions before release

## 📄 License
See `LICENSE.txt` for license terms and conditions.

## 🔗 Links
- **Repository**: https://github.com/DRAravind-cpu/adjustment-calculator
- **Issues**: https://github.com/DRAravind-cpu/adjustment-calculator/issues
- **Releases**: https://github.com/DRAravind-cpu/adjustment-calculator/releases

---

**Note**: This offline version maintains full feature parity with the web version while providing the convenience of a standalone Windows application.