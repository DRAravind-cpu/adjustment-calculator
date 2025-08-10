#!/usr/bin/env python3
"""
Auto-Updater System for Energy Adjustment Calculator
Handles automatic updates when internet is available
"""

import os
import sys
import json
import requests
import zipfile
import tempfile
import shutil
import subprocess
import threading
import time
from pathlib import Path
from datetime import datetime, timedelta
import hashlib
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser

class AutoUpdater:
    def __init__(self, current_version="1.0.0", app_name="Energy Adjustment Calculator"):
        self.current_version = current_version
        self.app_name = app_name
        self.app_dir = Path(__file__).parent
        self.config_file = self.app_dir / "update_config.json"
        self.github_repo = "DRAravind-cpu/adjustment-calculator"
        self.github_api_url = f"https://api.github.com/repos/{self.github_repo}"
        self.update_check_interval = 24  # hours
        
        # Load configuration
        self.config = self.load_config()
        
    def load_config(self):
        """Load update configuration"""
        default_config = {
            "auto_check": True,
            "last_check": None,
            "update_channel": "stable",  # stable, beta, dev
            "backup_enabled": True,
            "silent_updates": False,
            "check_interval_hours": 24
        }
        
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Merge with defaults
                    for key, value in default_config.items():
                        if key not in config:
                            config[key] = value
                    return config
            except Exception as e:
                print(f"Error loading config: {e}")
        
        return default_config
    
    def save_config(self):
        """Save update configuration"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def check_internet_connection(self):
        """Check if internet connection is available"""
        try:
            response = requests.get("https://api.github.com", timeout=5)
            return response.status_code == 200
        except:
            return False
    
    def get_latest_version_info(self):
        """Get latest version information from GitHub"""
        try:
            # Get latest release
            response = requests.get(f"{self.github_api_url}/releases/latest", timeout=10)
            if response.status_code == 200:
                release_data = response.json()
                return {
                    "version": release_data["tag_name"].lstrip('v'),
                    "name": release_data["name"],
                    "description": release_data["body"],
                    "download_url": release_data["zipball_url"],
                    "published_at": release_data["published_at"],
                    "assets": release_data.get("assets", [])
                }
            else:
                # Fallback: check commits for version info
                response = requests.get(f"{self.github_api_url}/commits", timeout=10)
                if response.status_code == 200:
                    commits = response.json()
                    latest_commit = commits[0]
                    return {
                        "version": f"dev-{latest_commit['sha'][:8]}",
                        "name": "Latest Development Version",
                        "description": latest_commit["commit"]["message"],
                        "download_url": f"https://github.com/{self.github_repo}/archive/main.zip",
                        "published_at": latest_commit["commit"]["committer"]["date"],
                        "assets": []
                    }
        except Exception as e:
            print(f"Error checking for updates: {e}")
            return None
    
    def compare_versions(self, version1, version2):
        """Compare two version strings"""
        def version_tuple(v):
            if v.startswith('dev-'):
                return (0, 0, 0, v)  # Dev versions are always newer
            try:
                return tuple(map(int, v.split('.')))
            except:
                return (0, 0, 0)
        
        v1_tuple = version_tuple(version1)
        v2_tuple = version_tuple(version2)
        
        if v1_tuple < v2_tuple:
            return -1
        elif v1_tuple > v2_tuple:
            return 1
        else:
            return 0
    
    def should_check_for_updates(self):
        """Determine if we should check for updates"""
        if not self.config["auto_check"]:
            return False
        
        last_check = self.config.get("last_check")
        if not last_check:
            return True
        
        try:
            last_check_time = datetime.fromisoformat(last_check)
            time_since_check = datetime.now() - last_check_time
            return time_since_check > timedelta(hours=self.config["check_interval_hours"])
        except:
            return True
    
    def check_for_updates(self, show_no_updates=False):
        """Check for available updates"""
        if not self.check_internet_connection():
            if show_no_updates:
                messagebox.showinfo("Update Check", "No internet connection available.")
            return None
        
        try:
            latest_info = self.get_latest_version_info()
            if not latest_info:
                if show_no_updates:
                    messagebox.showerror("Update Check", "Failed to check for updates.")
                return None
            
            # Update last check time
            self.config["last_check"] = datetime.now().isoformat()
            self.save_config()
            
            # Compare versions
            if self.compare_versions(self.current_version, latest_info["version"]) < 0:
                return latest_info
            else:
                if show_no_updates:
                    messagebox.showinfo("Update Check", 
                                      f"You have the latest version ({self.current_version})")
                return None
                
        except Exception as e:
            if show_no_updates:
                messagebox.showerror("Update Check", f"Error checking for updates: {e}")
            return None
    
    def download_update(self, update_info, progress_callback=None):
        """Download the update"""
        try:
            download_url = update_info["download_url"]
            
            # Create temporary file
            temp_file = tempfile.NamedTemporaryFile(suffix='.zip', delete=False)
            
            # Download with progress
            response = requests.get(download_url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            
            with open(temp_file.name, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress = (downloaded / total_size) * 100
                            progress_callback(progress)
            
            return temp_file.name
            
        except Exception as e:
            print(f"Error downloading update: {e}")
            return None
    
    def backup_current_version(self):
        """Create backup of current version"""
        if not self.config["backup_enabled"]:
            return True
        
        try:
            backup_dir = self.app_dir / "backups"
            backup_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"backup_{self.current_version}_{timestamp}.zip"
            backup_path = backup_dir / backup_name
            
            # Create backup ZIP
            with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in self.app_dir.rglob('*'):
                    if (file_path.is_file() and 
                        not file_path.is_relative_to(backup_dir) and
                        not file_path.name.startswith('.')):
                        arcname = file_path.relative_to(self.app_dir)
                        zipf.write(file_path, arcname)
            
            # Keep only last 5 backups
            backups = sorted(backup_dir.glob("backup_*.zip"))
            if len(backups) > 5:
                for old_backup in backups[:-5]:
                    old_backup.unlink()
            
            return True
            
        except Exception as e:
            print(f"Error creating backup: {e}")
            return False
    
    def apply_update(self, update_file, update_info):
        """Apply the downloaded update"""
        try:
            # Create backup
            if not self.backup_current_version():
                print("Warning: Backup creation failed")
            
            # Extract update
            temp_dir = tempfile.mkdtemp()
            
            with zipfile.ZipFile(update_file, 'r') as zipf:
                zipf.extractall(temp_dir)
            
            # Find the extracted folder (GitHub creates a folder with repo name)
            extracted_folders = [d for d in Path(temp_dir).iterdir() if d.is_dir()]
            if not extracted_folders:
                raise Exception("No extracted folder found")
            
            source_dir = extracted_folders[0]
            
            # Files to update (exclude user data and configs)
            update_files = [
                'streamlit_app.py',
                'app.py',
                'launcher.py',
                'requirements.txt',
                'requirements_windows.txt',
                'templates/',
                '.streamlit/',
                'auto_updater.py'
            ]
            
            # Copy updated files
            for file_pattern in update_files:
                source_path = source_dir / file_pattern
                dest_path = self.app_dir / file_pattern
                
                if source_path.exists():
                    if source_path.is_file():
                        shutil.copy2(source_path, dest_path)
                    elif source_path.is_dir():
                        if dest_path.exists():
                            shutil.rmtree(dest_path)
                        shutil.copytree(source_path, dest_path)
            
            # Update version info
            version_file = self.app_dir / "version.json"
            version_data = {
                "version": update_info["version"],
                "updated_at": datetime.now().isoformat(),
                "update_source": "auto_updater"
            }
            
            with open(version_file, 'w') as f:
                json.dump(version_data, f, indent=2)
            
            # Cleanup
            shutil.rmtree(temp_dir)
            os.unlink(update_file)
            
            return True
            
        except Exception as e:
            print(f"Error applying update: {e}")
            return False
    
    def show_update_dialog(self, update_info):
        """Show update available dialog"""
        dialog = UpdateDialog(self, update_info)
        return dialog.show()
    
    def auto_check_updates(self):
        """Automatically check for updates in background"""
        def check_thread():
            if self.should_check_for_updates():
                update_info = self.check_for_updates()
                if update_info and not self.config["silent_updates"]:
                    # Show update dialog in main thread
                    self.show_update_dialog(update_info)
        
        threading.Thread(target=check_thread, daemon=True).start()

class UpdateDialog:
    def __init__(self, updater, update_info):
        self.updater = updater
        self.update_info = update_info
        self.result = None
        
    def show(self):
        """Show the update dialog"""
        root = tk.Tk()
        root.title("Update Available")
        root.geometry("500x400")
        root.resizable(False, False)
        
        # Center window
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        # Header
        header_frame = ttk.Frame(root)
        header_frame.pack(fill='x', padx=20, pady=20)
        
        ttk.Label(header_frame, text="🔄 Update Available", 
                 font=('Arial', 16, 'bold')).pack()
        
        ttk.Label(header_frame, 
                 text=f"Version {self.update_info['version']} is available",
                 font=('Arial', 10)).pack()
        
        # Current vs New version
        version_frame = ttk.LabelFrame(root, text="Version Information", padding=10)
        version_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(version_frame, 
                 text=f"Current Version: {self.updater.current_version}").pack(anchor='w')
        ttk.Label(version_frame, 
                 text=f"New Version: {self.update_info['version']}").pack(anchor='w')
        
        # Release notes
        notes_frame = ttk.LabelFrame(root, text="What's New", padding=10)
        notes_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        notes_text = tk.Text(notes_frame, height=8, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(notes_frame, orient="vertical", command=notes_text.yview)
        notes_text.configure(yscrollcommand=scrollbar.set)
        
        notes_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Insert release notes
        description = self.update_info.get('description', 'No release notes available.')
        notes_text.insert('1.0', description)
        notes_text.config(state='disabled')
        
        # Progress bar (hidden initially)
        self.progress_frame = ttk.Frame(root)
        self.progress_var = tk.StringVar(value="")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_var)
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate')
        
        # Buttons
        button_frame = ttk.Frame(root)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        def update_now():
            self.result = 'update'
            self.start_update(root)
        
        def update_later():
            self.result = 'later'
            root.quit()
        
        def skip_version():
            self.result = 'skip'
            root.quit()
        
        ttk.Button(button_frame, text="Update Now", 
                  command=update_now).pack(side='right', padx=(5,0))
        ttk.Button(button_frame, text="Later", 
                  command=update_later).pack(side='right')
        ttk.Button(button_frame, text="Skip This Version", 
                  command=skip_version).pack(side='left')
        
        root.mainloop()
        root.destroy()
        
        return self.result
    
    def start_update(self, root):
        """Start the update process"""
        # Show progress
        self.progress_frame.pack(fill='x', padx=20, pady=(0,20))
        self.progress_label.pack()
        self.progress_bar.pack(fill='x', pady=(5,0))
        
        def update_progress(progress):
            self.progress_bar['value'] = progress
            self.progress_var.set(f"Downloading update... {progress:.1f}%")
            root.update()
        
        def update_thread():
            try:
                self.progress_var.set("Starting download...")
                root.update()
                
                # Download update
                update_file = self.updater.download_update(self.update_info, update_progress)
                if not update_file:
                    messagebox.showerror("Update Failed", "Failed to download update.")
                    return
                
                self.progress_var.set("Applying update...")
                self.progress_bar['mode'] = 'indeterminate'
                self.progress_bar.start()
                root.update()
                
                # Apply update
                if self.updater.apply_update(update_file, self.update_info):
                    self.progress_bar.stop()
                    messagebox.showinfo("Update Complete", 
                                      "Update applied successfully! Please restart the application.")
                    root.quit()
                else:
                    self.progress_bar.stop()
                    messagebox.showerror("Update Failed", "Failed to apply update.")
                    
            except Exception as e:
                self.progress_bar.stop()
                messagebox.showerror("Update Error", f"Update failed: {e}")
        
        threading.Thread(target=update_thread, daemon=True).start()

class UpdateSettings:
    def __init__(self, updater):
        self.updater = updater
        
    def show_settings(self):
        """Show update settings dialog"""
        root = tk.Tk()
        root.title("Update Settings")
        root.geometry("400x300")
        
        # Variables
        auto_check = tk.BooleanVar(value=self.updater.config["auto_check"])
        backup_enabled = tk.BooleanVar(value=self.updater.config["backup_enabled"])
        silent_updates = tk.BooleanVar(value=self.updater.config["silent_updates"])
        check_interval = tk.IntVar(value=self.updater.config["check_interval_hours"])
        
        # Settings
        settings_frame = ttk.LabelFrame(root, text="Update Settings", padding=10)
        settings_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        ttk.Checkbutton(settings_frame, text="Automatically check for updates", 
                       variable=auto_check).pack(anchor='w', pady=5)
        
        ttk.Checkbutton(settings_frame, text="Create backup before updating", 
                       variable=backup_enabled).pack(anchor='w', pady=5)
        
        ttk.Checkbutton(settings_frame, text="Silent updates (no notifications)", 
                       variable=silent_updates).pack(anchor='w', pady=5)
        
        # Check interval
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.pack(fill='x', pady=10)
        
        ttk.Label(interval_frame, text="Check interval (hours):").pack(side='left')
        ttk.Spinbox(interval_frame, from_=1, to=168, textvariable=check_interval, 
                   width=10).pack(side='right')
        
        # Buttons
        button_frame = ttk.Frame(root)
        button_frame.pack(fill='x', padx=20, pady=20)
        
        def save_settings():
            self.updater.config["auto_check"] = auto_check.get()
            self.updater.config["backup_enabled"] = backup_enabled.get()
            self.updater.config["silent_updates"] = silent_updates.get()
            self.updater.config["check_interval_hours"] = check_interval.get()
            self.updater.save_config()
            messagebox.showinfo("Settings", "Settings saved successfully!")
            root.destroy()
        
        ttk.Button(button_frame, text="Save", command=save_settings).pack(side='right', padx=(5,0))
        ttk.Button(button_frame, text="Cancel", command=root.destroy).pack(side='right')
        
        root.mainloop()

# Integration functions for the main application
def initialize_updater(current_version="1.0.0"):
    """Initialize the auto-updater system"""
    return AutoUpdater(current_version)

def check_for_updates_on_startup(updater):
    """Check for updates when application starts"""
    updater.auto_check_updates()

def show_update_settings(updater):
    """Show update settings dialog"""
    settings = UpdateSettings(updater)
    settings.show_settings()

def manual_update_check(updater):
    """Manually check for updates"""
    update_info = updater.check_for_updates(show_no_updates=True)
    if update_info:
        updater.show_update_dialog(update_info)

if __name__ == "__main__":
    # Test the updater
    updater = initialize_updater("1.0.0")
    manual_update_check(updater)