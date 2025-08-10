#!/usr/bin/env python3
"""
Ultimate Windows Installer Creator
Creates a single executable file that contains everything needed
"""

import os
import sys
import zipfile
import base64
import tempfile
import subprocess
from pathlib import Path

def create_ultimate_installer():
    """Create the ultimate single-file installer"""
    
    # All files to embed
    files_to_embed = [
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
    
    # Create the ultimate installer script
    ultimate_installer = '''#!/usr/bin/env python3
"""
Energy Adjustment Calculator - Ultimate Windows Installer
Single file installer with GUI and all components embedded
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import zipfile
import tempfile
import shutil
import subprocess
import base64
import threading
from pathlib import Path
import webbrowser

# Embedded application data
EMBEDDED_DATA = """{{EMBEDDED_DATA}}"""

class UltimateInstaller:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Energy Adjustment Calculator - Ultimate Installer")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Variables
        self.install_path = tk.StringVar(value=str(Path.home() / "EnergyAdjustmentCalculator"))
        self.build_executable = tk.BooleanVar(value=True)
        self.create_shortcuts = tk.BooleanVar(value=True)
        self.open_after_install = tk.BooleanVar(value=True)
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.setup_ui()
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def setup_ui(self):
        """Setup the user interface"""
        
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=20, pady=20)
        
        # Logo/Title
        title_label = ttk.Label(header_frame, text="⚡ Energy Adjustment Calculator", 
                               font=('Arial', 18, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, text="Ultimate Windows Installer", 
                                  font=('Arial', 10))
        subtitle_label.pack()
        
        # Description
        desc_frame = ttk.LabelFrame(self.root, text="About", padding=10)
        desc_frame.pack(fill='x', padx=20, pady=(0,10))
        
        desc_text = """This installer will set up the Energy Adjustment Calculator on your Windows system.
The application works completely offline and includes all necessary components.

Features:
• Excel file processing for energy data
• T&D loss calculations and adjustments  
• PDF report generation with multiple output options
• TOD (Time of Day) categorization
• Complete offline functionality"""
        
        ttk.Label(desc_frame, text=desc_text, justify='left').pack(anchor='w')
        
        # Installation path
        path_frame = ttk.LabelFrame(self.root, text="Installation Directory", padding=10)
        path_frame.pack(fill='x', padx=20, pady=10)
        
        path_entry_frame = ttk.Frame(path_frame)
        path_entry_frame.pack(fill='x')
        
        self.path_entry = ttk.Entry(path_entry_frame, textvariable=self.install_path, font=('Consolas', 9))
        self.path_entry.pack(side='left', fill='x', expand=True)
        
        ttk.Button(path_entry_frame, text="Browse...", 
                  command=self.browse_path).pack(side='right', padx=(10,0))
        
        # Options
        options_frame = ttk.LabelFrame(self.root, text="Installation Options", padding=10)
        options_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Checkbutton(options_frame, 
                       text="Build Windows executable (requires Python 3.8+)", 
                       variable=self.build_executable).pack(anchor='w', pady=2)
        
        ttk.Checkbutton(options_frame, 
                       text="Create desktop shortcut and start menu entry", 
                       variable=self.create_shortcuts).pack(anchor='w', pady=2)
        
        ttk.Checkbutton(options_frame, 
                       text="Open application after installation", 
                       variable=self.open_after_install).pack(anchor='w', pady=2)
        
        # Progress section
        self.progress_frame = ttk.LabelFrame(self.root, text="Installation Progress", padding=10)
        self.progress_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Progress bar
        self.progress_var = tk.StringVar(value="Ready to install...")
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.progress_var)
        self.status_label.pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=(5,10))
        
        # Log area
        self.log_text = scrolledtext.ScrolledText(self.progress_frame, height=8, 
                                                 font=('Consolas', 8))
        self.log_text.pack(fill='both', expand=True)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        self.install_btn = ttk.Button(button_frame, text="Install", 
                                     command=self.start_installation)
        self.install_btn.pack(side='right', padx=(10,0))
        
        ttk.Button(button_frame, text="Exit", 
                  command=self.root.quit).pack(side='right')
        
        # Help button
        ttk.Button(button_frame, text="Help", 
                  command=self.show_help).pack(side='left')
    
    def browse_path(self):
        """Browse for installation directory"""
        path = filedialog.askdirectory(initialdir=self.install_path.get(),
                                      title="Select Installation Directory")
        if path:
            self.install_path.set(path)
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_progress(self, value, status):
        """Update progress bar and status"""
        self.progress_bar['value'] = value
        self.progress_var.set(status)
        self.root.update()
    
    def start_installation(self):
        """Start installation in separate thread"""
        self.install_btn.config(state='disabled')
        threading.Thread(target=self.install, daemon=True).start()
    
    def install(self):
        """Main installation process"""
        try:
            self.log("Starting installation...")
            self.update_progress(10, "Preparing installation...")
            
            # Create installation directory
            install_path = Path(self.install_path.get())
            install_path.mkdir(parents=True, exist_ok=True)
            self.log(f"Created directory: {install_path}")
            
            self.update_progress(20, "Extracting application files...")
            
            # Extract embedded files
            temp_dir = tempfile.mkdtemp()
            try:
                zip_data = base64.b64decode(EMBEDDED_DATA)
                zip_path = os.path.join(temp_dir, "app_bundle.zip")
                
                with open(zip_path, 'wb') as f:
                    f.write(zip_data)
                
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(install_path)
                
                self.log("Application files extracted successfully")
                self.update_progress(40, "Files extracted successfully")
                
                # Build executable if requested
                if self.build_executable.get():
                    self.update_progress(50, "Checking Python installation...")
                    
                    # Check Python
                    try:
                        result = subprocess.run([sys.executable, "--version"], 
                                              capture_output=True, text=True)
                        python_version = result.stdout.strip()
                        self.log(f"Found Python: {python_version}")
                        
                        self.update_progress(60, "Installing dependencies...")
                        os.chdir(install_path)
                        
                        # Install requirements
                        subprocess.run([sys.executable, "-m", "pip", "install", 
                                      "-r", "requirements_windows.txt"], 
                                     capture_output=True)
                        self.log("Dependencies installed")
                        
                        self.update_progress(80, "Building Windows executable...")
                        
                        # Build executable
                        if os.name == 'nt':
                            result = subprocess.run(["build_windows.bat"], 
                                                  shell=True, capture_output=True, text=True)
                        else:
                            result = subprocess.run([sys.executable, "build_windows_app.py"], 
                                                  capture_output=True, text=True)
                        
                        if result.returncode == 0:
                            self.log("Executable built successfully!")
                            exe_path = install_path / "dist" / "EnergyAdjustmentCalculator.exe"
                            if exe_path.exists():
                                self.log(f"Executable location: {exe_path}")
                        else:
                            self.log("Build failed - you can build manually later")
                            self.log(result.stderr)
                        
                    except Exception as e:
                        self.log(f"Python build failed: {e}")
                        self.log("You can build the executable manually later")
                
                # Create shortcuts if requested
                if self.create_shortcuts.get():
                    self.update_progress(90, "Creating shortcuts...")
                    self.create_shortcuts_func(install_path)
                
                self.update_progress(100, "Installation completed successfully!")
                self.log("\\n" + "="*50)
                self.log("INSTALLATION COMPLETED SUCCESSFULLY!")
                self.log("="*50)
                self.log(f"Installation directory: {install_path}")
                
                exe_path = install_path / "dist" / "EnergyAdjustmentCalculator.exe"
                if exe_path.exists():
                    self.log(f"Executable: {exe_path}")
                    self.log("\\nTo run: Double-click the executable")
                else:
                    self.log("\\nTo build executable: Run build_windows.bat")
                
                self.log("\\nSee WINDOWS_APP_GUIDE.md for detailed instructions")
                
                # Show completion dialog
                result = messagebox.showinfo("Installation Complete", 
                    f"Energy Adjustment Calculator installed successfully!\\n\\n"
                    f"Location: {install_path}\\n\\n"
                    f"{'Executable ready to use!' if exe_path.exists() else 'Run build_windows.bat to create executable'}")
                
                # Open application if requested
                if self.open_after_install.get() and exe_path.exists():
                    self.log("Opening application...")
                    subprocess.Popen([str(exe_path)])
                
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)
                
        except Exception as e:
            self.log(f"\\nERROR: {e}")
            self.update_progress(0, f"Installation failed: {e}")
            messagebox.showerror("Installation Failed", 
                               f"Installation failed with error:\\n{e}\\n\\n"
                               f"Check the log for details.")
        finally:
            self.install_btn.config(state='normal')
    
    def create_shortcuts_func(self, install_path):
        """Create desktop and start menu shortcuts"""
        try:
            exe_path = install_path / "dist" / "EnergyAdjustmentCalculator.exe"
            
            if os.name == 'nt' and exe_path.exists():
                # Create desktop shortcut
                desktop = Path.home() / "Desktop"
                shortcut_path = desktop / "Energy Adjustment Calculator.lnk"
                
                # Simple batch file as shortcut alternative
                batch_path = desktop / "Energy Adjustment Calculator.bat"
                with open(batch_path, 'w') as f:
                    f.write(f'@echo off\\ncd /d "{install_path}"\\nstart "" "{exe_path}"')
                
                self.log("Desktop shortcut created")
            
        except Exception as e:
            self.log(f"Shortcut creation failed: {e}")
    
    def show_help(self):
        """Show help dialog"""
        help_text = """Energy Adjustment Calculator - Installation Help

SYSTEM REQUIREMENTS:
• Windows 7, 8, 10, or 11
• Python 3.8+ (for building executable)
• 4GB RAM recommended
• 500MB free disk space

INSTALLATION OPTIONS:
• Build executable: Creates standalone .exe file
• Create shortcuts: Adds desktop and start menu shortcuts  
• Open after install: Launches app after installation

USAGE:
1. Choose installation directory
2. Select desired options
3. Click Install
4. Wait for completion
5. Run the application

SUPPORT:
• Check installation log for errors
• See documentation files after installation
• Visit GitHub repository for updates

For technical support, check the log output and documentation files."""
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Installation Help")
        help_window.geometry("500x400")
        
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill='both', expand=True)
        text_widget.insert('1.0', help_text)
        text_widget.config(state='disabled')
        
        ttk.Button(help_window, text="Close", 
                  command=help_window.destroy).pack(pady=10)
    
    def run(self):
        """Run the installer"""
        self.root.mainloop()

def main():
    """Main entry point"""
    try:
        app = UltimateInstaller()
        app.run()
    except Exception as e:
        # Fallback console mode
        print("GUI mode failed, falling back to console mode...")
        print(f"Error: {e}")
        
        # Simple console installation
        install_path = input("Enter installation directory: ").strip()
        if not install_path:
            install_path = str(Path.home() / "EnergyAdjustmentCalculator")
        
        print(f"Installing to: {install_path}")
        
        # Extract files
        install_dir = Path(install_path)
        install_dir.mkdir(parents=True, exist_ok=True)
        
        zip_data = base64.b64decode(EMBEDDED_DATA)
        temp_zip = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
        
        try:
            with open(temp_zip.name, 'wb') as f:
                f.write(zip_data)
            
            with zipfile.ZipFile(temp_zip.name, 'r') as zip_ref:
                zip_ref.extractall(install_dir)
            
            print("Installation completed!")
            print(f"Files installed to: {install_dir}")
            print("Run build_windows.bat to create the executable")
            
        finally:
            os.unlink(temp_zip.name)

if __name__ == "__main__":
    main()
'''
    
    print("Creating ultimate installer...")
    
    # Create ZIP with all files
    temp_zip = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
    
    try:
        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in files_to_embed:
                full_path = Path(file_path)
                if full_path.exists():
                    if full_path.is_file():
                        zipf.write(full_path, file_path)
                        print(f"Added: {file_path}")
                    elif full_path.is_dir():
                        for sub_file in full_path.rglob('*'):
                            if sub_file.is_file():
                                arcname = str(sub_file)
                                zipf.write(sub_file, arcname)
                                print(f"Added: {arcname}")
                else:
                    print(f"Warning: {file_path} not found")
        
        # Read and encode ZIP data
        with open(temp_zip.name, 'rb') as f:
            zip_data = f.read()
        
        encoded_data = base64.b64encode(zip_data).decode('utf-8')
        
        # Create final installer
        final_installer = ultimate_installer.replace('{{EMBEDDED_DATA}}', encoded_data)
        
        # Write the ultimate installer
        installer_path = Path('EnergyCalculator_UltimateInstaller.py')
        with open(installer_path, 'w', encoding='utf-8') as f:
            f.write(final_installer)
        
        print(f"\\nUltimate installer created: {installer_path}")
        print(f"File size: {installer_path.stat().st_size / 1024 / 1024:.1f} MB")
        
        # Create Windows executable version using PyInstaller
        try:
            print("\\nCreating Windows executable...")
            subprocess.run([
                sys.executable, '-m', 'PyInstaller',
                '--onefile',
                '--windowed', 
                '--name=EnergyCalculator_Setup',
                '--icon=app_icon.ico' if Path('app_icon.ico').exists() else '',
                str(installer_path)
            ], check=True, capture_output=True)
            
            exe_path = Path('dist/EnergyCalculator_Setup.exe')
            if exe_path.exists():
                print(f"Windows executable created: {exe_path}")
                print(f"Executable size: {exe_path.stat().st_size / 1024 / 1024:.1f} MB")
            
        except subprocess.CalledProcessError:
            print("PyInstaller not available - Python script version created")
        except Exception as e:
            print(f"Executable creation failed: {e}")
        
    finally:
        os.unlink(temp_zip.name)

if __name__ == "__main__":
    create_ultimate_installer()