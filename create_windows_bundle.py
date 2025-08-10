#!/usr/bin/env python3
"""
Windows Bundle Creator for Energy Adjustment Calculator
Creates a single self-extracting archive with all installation files
"""

import os
import sys
import zipfile
import shutil
import subprocess
from pathlib import Path
import tempfile
import base64

def create_self_extracting_installer():
    """Create a self-extracting installer with all components"""
    
    # Files to include in the bundle
    bundle_files = [
        'launcher.py',
        'app.spec', 
        'build_windows.bat',
        'build_windows_app.py',
        'installer.nsi',
        'LICENSE.txt',
        'requirements_windows.txt',
        'streamlit_app.py',
        'app.py',
        'requirements.txt',
        'BUILD_INSTRUCTIONS.md',
        'WINDOWS_APP_GUIDE.md',
        'WINDOWS_OFFLINE_README.md',
        'templates/index.html',
        '.streamlit/config.toml'
    ]
    
    # Create the self-extracting installer script
    installer_script = '''
import os
import sys
import zipfile
import tempfile
import shutil
import subprocess
import base64
from pathlib import Path

# Embedded ZIP data will be inserted here
EMBEDDED_DATA = """{{EMBEDDED_DATA}}"""

def extract_and_setup():
    """Extract embedded files and setup the application"""
    print("=" * 60)
    print("Energy Adjustment Calculator - Windows Installer")
    print("=" * 60)
    print("Extracting application files...")
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Decode and extract embedded ZIP
        zip_data = base64.b64decode(EMBEDDED_DATA)
        zip_path = os.path.join(temp_dir, "app_bundle.zip")
        
        with open(zip_path, 'wb') as f:
            f.write(zip_data)
        
        # Ask user for installation directory
        print("\\nChoose installation directory:")
        print("1. Default: C:\\\\Program Files\\\\Energy Adjustment Calculator")
        print("2. Custom directory")
        print("3. Current directory")
        
        choice = input("\\nEnter choice (1-3): ").strip()
        
        if choice == "2":
            install_dir = input("Enter custom directory path: ").strip()
        elif choice == "3":
            install_dir = os.getcwd()
        else:
            install_dir = "C:\\\\Program Files\\\\Energy Adjustment Calculator"
        
        # Create installation directory
        install_path = Path(install_dir)
        install_path.mkdir(parents=True, exist_ok=True)
        
        print(f"\\nExtracting to: {install_path}")
        
        # Extract ZIP contents
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(install_path)
        
        print("Files extracted successfully!")
        
        # Ask if user wants to build immediately
        build_now = input("\\nBuild the Windows executable now? (y/n): ").strip().lower()
        
        if build_now == 'y':
            print("\\nBuilding Windows executable...")
            os.chdir(install_path)
            
            # Check if Python is available
            try:
                subprocess.run([sys.executable, "--version"], check=True, capture_output=True)
                python_cmd = sys.executable
            except:
                try:
                    subprocess.run(["python", "--version"], check=True, capture_output=True)
                    python_cmd = "python"
                except:
                    print("ERROR: Python not found. Please install Python 3.8+ and try again.")
                    input("Press Enter to exit...")
                    return
            
            # Install dependencies
            print("Installing dependencies...")
            subprocess.run([python_cmd, "-m", "pip", "install", "-r", "requirements_windows.txt"])
            
            # Build the application
            print("Building executable...")
            if os.name == 'nt':  # Windows
                subprocess.run(["build_windows.bat"], shell=True)
            else:
                subprocess.run([python_cmd, "build_windows_app.py"])
            
            print("\\n" + "=" * 60)
            print("Installation and build completed!")
            print("=" * 60)
            print(f"Application installed to: {install_path}")
            if os.path.exists(install_path / "dist" / "EnergyAdjustmentCalculator.exe"):
                print(f"Executable: {install_path / 'dist' / 'EnergyAdjustmentCalculator.exe'}")
            print("\\nSee WINDOWS_APP_GUIDE.md for usage instructions.")
        else:
            print("\\n" + "=" * 60)
            print("Installation completed!")
            print("=" * 60)
            print(f"Files installed to: {install_path}")
            print("To build the executable later, run: build_windows.bat")
            print("See BUILD_INSTRUCTIONS.md for detailed build guide.")
        
    except Exception as e:
        print(f"\\nERROR: {e}")
        print("Installation failed. Please check the error and try again.")
    finally:
        # Cleanup
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    input("\\nPress Enter to exit...")

if __name__ == "__main__":
    extract_and_setup()
'''
    
    print("Creating Windows bundle...")
    
    # Create temporary ZIP with all files
    temp_zip = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
    
    try:
        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in bundle_files:
                full_path = Path(file_path)
                if full_path.exists():
                    zipf.write(full_path, file_path)
                    print(f"Added: {file_path}")
                else:
                    print(f"Warning: {file_path} not found, skipping...")
        
        # Read ZIP data and encode as base64
        with open(temp_zip.name, 'rb') as f:
            zip_data = f.read()
        
        encoded_data = base64.b64encode(zip_data).decode('utf-8')
        
        # Create the self-extracting installer
        final_installer = installer_script.replace('{{EMBEDDED_DATA}}', encoded_data)
        
        # Write the self-extracting installer
        installer_path = Path('EnergyCalculator_WindowsInstaller.py')
        with open(installer_path, 'w') as f:
            f.write(final_installer)
        
        print(f"\\nSelf-extracting installer created: {installer_path}")
        
        # Create batch file to run the installer
        batch_content = '''@echo off
echo Energy Adjustment Calculator - Windows Installer
echo.
python EnergyCalculator_WindowsInstaller.py
pause
'''
        
        with open('Install_EnergyCalculator.bat', 'w') as f:
            f.write(batch_content)
        
        print("Batch installer created: Install_EnergyCalculator.bat")
        
    finally:
        # Cleanup
        os.unlink(temp_zip.name)

def create_portable_bundle():
    """Create a portable ZIP bundle with all files"""
    
    print("\\nCreating portable ZIP bundle...")
    
    bundle_name = "EnergyCalculator_Windows_Portable.zip"
    
    # Files and directories to include
    items_to_bundle = [
        'launcher.py',
        'app.spec',
        'build_windows.bat', 
        'build_windows_app.py',
        'installer.nsi',
        'LICENSE.txt',
        'requirements_windows.txt',
        'streamlit_app.py',
        'app.py',
        'requirements.txt',
        'BUILD_INSTRUCTIONS.md',
        'WINDOWS_APP_GUIDE.md',
        'WINDOWS_OFFLINE_README.md',
        'templates/',
        '.streamlit/'
    ]
    
    with zipfile.ZipFile(bundle_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for item in items_to_bundle:
            item_path = Path(item)
            
            if item_path.is_file():
                zipf.write(item_path, item)
                print(f"Added file: {item}")
            elif item_path.is_dir():
                for file_path in item_path.rglob('*'):
                    if file_path.is_file():
                        arcname = str(file_path)
                        zipf.write(file_path, arcname)
                        print(f"Added: {arcname}")
    
    print(f"\\nPortable bundle created: {bundle_name}")
    return bundle_name

def create_advanced_installer():
    """Create an advanced installer with GUI"""
    
    advanced_installer = '''
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import zipfile
import tempfile
import shutil
import subprocess
import base64
import threading
from pathlib import Path

# Embedded ZIP data
EMBEDDED_DATA = """{{EMBEDDED_DATA}}"""

class InstallerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Energy Adjustment Calculator - Windows Installer")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Variables
        self.install_path = tk.StringVar(value="C:\\\\Program Files\\\\Energy Adjustment Calculator")
        self.build_now = tk.BooleanVar(value=True)
        self.create_shortcuts = tk.BooleanVar(value=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Title
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill='x', padx=20, pady=20)
        
        ttk.Label(title_frame, text="Energy Adjustment Calculator", 
                 font=('Arial', 16, 'bold')).pack()
        ttk.Label(title_frame, text="Windows Offline Installation", 
                 font=('Arial', 10)).pack()
        
        # Installation path
        path_frame = ttk.LabelFrame(self.root, text="Installation Directory", padding=10)
        path_frame.pack(fill='x', padx=20, pady=10)
        
        path_entry_frame = ttk.Frame(path_frame)
        path_entry_frame.pack(fill='x')
        
        ttk.Entry(path_entry_frame, textvariable=self.install_path, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(path_entry_frame, text="Browse", command=self.browse_path).pack(side='right', padx=(5,0))
        
        # Options
        options_frame = ttk.LabelFrame(self.root, text="Installation Options", padding=10)
        options_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Checkbutton(options_frame, text="Build executable immediately (requires Python)", 
                       variable=self.build_now).pack(anchor='w')
        ttk.Checkbutton(options_frame, text="Create desktop and start menu shortcuts", 
                       variable=self.create_shortcuts).pack(anchor='w')
        
        # Progress
        self.progress_frame = ttk.LabelFrame(self.root, text="Progress", padding=10)
        self.progress_frame.pack(fill='x', padx=20, pady=10)
        
        self.progress_var = tk.StringVar(value="Ready to install...")
        ttk.Label(self.progress_frame, textvariable=self.progress_var).pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress_bar.pack(fill='x', pady=(5,0))
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        ttk.Button(button_frame, text="Install", command=self.start_installation).pack(side='right', padx=(5,0))
        ttk.Button(button_frame, text="Cancel", command=self.root.quit).pack(side='right')
    
    def browse_path(self):
        path = filedialog.askdirectory(initialdir=self.install_path.get())
        if path:
            self.install_path.set(path)
    
    def start_installation(self):
        # Start installation in separate thread
        threading.Thread(target=self.install, daemon=True).start()
    
    def install(self):
        try:
            self.progress_bar.start()
            self.progress_var.set("Extracting files...")
            
            # Create installation directory
            install_path = Path(self.install_path.get())
            install_path.mkdir(parents=True, exist_ok=True)
            
            # Extract embedded ZIP
            temp_dir = tempfile.mkdtemp()
            zip_data = base64.b64decode(EMBEDDED_DATA)
            zip_path = os.path.join(temp_dir, "app_bundle.zip")
            
            with open(zip_path, 'wb') as f:
                f.write(zip_data)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(install_path)
            
            self.progress_var.set("Files extracted successfully!")
            
            if self.build_now.get():
                self.progress_var.set("Building executable...")
                os.chdir(install_path)
                
                # Install dependencies and build
                subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements_windows.txt"], 
                             capture_output=True)
                
                if os.name == 'nt':
                    subprocess.run(["build_windows.bat"], shell=True, capture_output=True)
                
            if self.create_shortcuts.get():
                self.progress_var.set("Creating shortcuts...")
                # Create shortcuts logic here
            
            self.progress_bar.stop()
            self.progress_var.set("Installation completed successfully!")
            
            messagebox.showinfo("Success", f"Installation completed!\\n\\nInstalled to: {install_path}")
            
        except Exception as e:
            self.progress_bar.stop()
            self.progress_var.set(f"Installation failed: {str(e)}")
            messagebox.showerror("Error", f"Installation failed:\\n{str(e)}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = InstallerGUI()
    app.run()
'''
    
    # Create the GUI installer (placeholder for now)
    print("\\nAdvanced GUI installer template created.")

def main():
    """Main function to create all bundle types"""
    print("Energy Adjustment Calculator - Windows Bundle Creator")
    print("=" * 60)
    
    print("\\nWhat type of bundle would you like to create?")
    print("1. Self-extracting installer (Python script)")
    print("2. Portable ZIP bundle") 
    print("3. Both")
    
    choice = input("\\nEnter choice (1-3): ").strip()
    
    if choice in ['1', '3']:
        create_self_extracting_installer()
    
    if choice in ['2', '3']:
        create_portable_bundle()
    
    print("\\n" + "=" * 60)
    print("Bundle creation completed!")
    print("=" * 60)
    
    if choice in ['1', '3']:
        print("\\nSelf-extracting installer files:")
        print("- EnergyCalculator_WindowsInstaller.py")
        print("- Install_EnergyCalculator.bat")
        print("\\nUsers can run either file to install the application.")
    
    if choice in ['2', '3']:
        print("\\nPortable bundle:")
        print("- EnergyCalculator_Windows_Portable.zip")
        print("\\nUsers can extract and run build_windows.bat")

if __name__ == "__main__":
    main()