#!/usr/bin/env python3
"""
Update Notification System
Provides desktop notifications and system tray integration for updates
"""

import os
import sys
import threading
import time
from pathlib import Path
import json
from datetime import datetime, timedelta

# Try to import notification libraries
try:
    import tkinter as tk
    from tkinter import ttk, messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

try:
    # Windows notifications
    if os.name == 'nt':
        import win10toast
        TOAST_AVAILABLE = True
    else:
        TOAST_AVAILABLE = False
except ImportError:
    TOAST_AVAILABLE = False

try:
    from auto_updater import AutoUpdater
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False

class UpdateNotifier:
    """Handles update notifications and user interactions"""
    
    def __init__(self, updater=None):
        self.updater = updater
        self.notification_shown = False
        self.update_available = False
        self.latest_update_info = None
        
        # Initialize toast notifier for Windows
        if TOAST_AVAILABLE:
            self.toaster = win10toast.ToastNotifier()
        
    def show_toast_notification(self, title, message, duration=10):
        """Show Windows toast notification"""
        if not TOAST_AVAILABLE:
            return False
        
        try:
            self.toaster.show_toast(
                title,
                message,
                duration=duration,
                icon_path=None,  # Use default icon
                threaded=True
            )
            return True
        except Exception as e:
            print(f"Toast notification failed: {e}")
            return False
    
    def show_popup_notification(self, update_info):
        """Show popup notification dialog"""
        if not TKINTER_AVAILABLE:
            return False
        
        try:
            # Create notification window
            root = tk.Tk()
            root.withdraw()  # Hide main window
            
            # Show notification dialog
            result = messagebox.askyesno(
                "Update Available",
                f"Energy Adjustment Calculator v{update_info['version']} is available!\n\n"
                f"Current version: {self.updater.current_version}\n"
                f"New version: {update_info['version']}\n\n"
                f"What's new:\n{update_info['description'][:200]}...\n\n"
                f"Would you like to download and install the update now?",
                icon='question'
            )
            
            root.destroy()
            return result
            
        except Exception as e:
            print(f"Popup notification failed: {e}")
            return False
    
    def show_system_notification(self, update_info):
        """Show system notification using available method"""
        title = "Energy Adjustment Calculator - Update Available"
        message = f"Version {update_info['version']} is ready to install!"
        
        # Try toast notification first (Windows)
        if self.show_toast_notification(title, message):
            return True
        
        # Fallback to popup
        return self.show_popup_notification(update_info)
    
    def check_and_notify(self):
        """Check for updates and show notification if available"""
        if not self.updater:
            return False
        
        try:
            # Check for updates
            update_info = self.updater.check_for_updates(show_no_updates=False)
            
            if update_info and not self.notification_shown:
                self.latest_update_info = update_info
                self.update_available = True
                
                # Show notification
                if self.show_system_notification(update_info):
                    self.notification_shown = True
                    return True
                
        except Exception as e:
            print(f"Update check failed: {e}")
        
        return False
    
    def start_background_checking(self, interval_hours=24):
        """Start background update checking"""
        def check_loop():
            while True:
                try:
                    if self.updater and self.updater.should_check_for_updates():
                        self.check_and_notify()
                    
                    # Wait for next check
                    time.sleep(interval_hours * 3600)  # Convert hours to seconds
                    
                except Exception as e:
                    print(f"Background update check error: {e}")
                    time.sleep(3600)  # Wait 1 hour on error
        
        # Start background thread
        thread = threading.Thread(target=check_loop, daemon=True)
        thread.start()
        return thread

class UpdateStatusWidget:
    """Widget to show update status in applications"""
    
    def __init__(self, parent, updater):
        self.parent = parent
        self.updater = updater
        self.frame = None
        self.status_label = None
        self.update_button = None
        
    def create_widget(self):
        """Create the update status widget"""
        if not TKINTER_AVAILABLE:
            return None
        
        self.frame = ttk.Frame(self.parent)
        
        # Status label
        self.status_label = ttk.Label(self.frame, text="Checking for updates...")
        self.status_label.pack(side='left', padx=(0, 10))
        
        # Update button (initially hidden)
        self.update_button = ttk.Button(
            self.frame, 
            text="Update Available", 
            command=self.handle_update_click,
            state='disabled'
        )
        
        # Start update checking
        self.check_updates()
        
        return self.frame
    
    def check_updates(self):
        """Check for updates and update widget"""
        def check_thread():
            try:
                if self.updater.check_internet_connection():
                    update_info = self.updater.check_for_updates(show_no_updates=False)
                    
                    if update_info:
                        self.show_update_available(update_info)
                    else:
                        self.show_up_to_date()
                else:
                    self.show_offline()
                    
            except Exception as e:
                self.show_error(str(e))
        
        threading.Thread(target=check_thread, daemon=True).start()
    
    def show_update_available(self, update_info):
        """Show update available state"""
        if self.status_label:
            self.status_label.config(text=f"Update v{update_info['version']} available")
        if self.update_button:
            self.update_button.config(state='normal')
            self.update_button.pack(side='right')
        
        self.latest_update_info = update_info
    
    def show_up_to_date(self):
        """Show up to date state"""
        if self.status_label:
            self.status_label.config(text="✅ Up to date")
    
    def show_offline(self):
        """Show offline state"""
        if self.status_label:
            self.status_label.config(text="🔌 Offline")
    
    def show_error(self, error):
        """Show error state"""
        if self.status_label:
            self.status_label.config(text=f"❌ Error: {error}")
    
    def handle_update_click(self):
        """Handle update button click"""
        if hasattr(self, 'latest_update_info'):
            # Show update dialog
            notifier = UpdateNotifier(self.updater)
            if notifier.show_popup_notification(self.latest_update_info):
                self.start_update()
    
    def start_update(self):
        """Start the update process"""
        if not hasattr(self, 'latest_update_info'):
            return
        
        def update_thread():
            try:
                # Update status
                if self.status_label:
                    self.status_label.config(text="Downloading update...")
                
                # Download update
                update_file = self.updater.download_update(self.latest_update_info)
                
                if update_file:
                    if self.status_label:
                        self.status_label.config(text="Installing update...")
                    
                    # Apply update
                    if self.updater.apply_update(update_file, self.latest_update_info):
                        if self.status_label:
                            self.status_label.config(text="✅ Update installed - Restart required")
                        
                        messagebox.showinfo(
                            "Update Complete",
                            "Update installed successfully!\nPlease restart the application to use the new version."
                        )
                    else:
                        if self.status_label:
                            self.status_label.config(text="❌ Update failed")
                        messagebox.showerror("Update Failed", "Failed to install update.")
                else:
                    if self.status_label:
                        self.status_label.config(text="❌ Download failed")
                    messagebox.showerror("Update Failed", "Failed to download update.")
                    
            except Exception as e:
                if self.status_label:
                    self.status_label.config(text=f"❌ Error: {str(e)}")
                messagebox.showerror("Update Error", f"Update failed: {e}")
        
        threading.Thread(target=update_thread, daemon=True).start()

def create_update_notifier(updater):
    """Factory function to create update notifier"""
    return UpdateNotifier(updater)

def show_update_notification(update_info, updater):
    """Show update notification"""
    notifier = UpdateNotifier(updater)
    return notifier.show_system_notification(update_info)

def start_background_update_checking(updater, interval_hours=24):
    """Start background update checking"""
    notifier = UpdateNotifier(updater)
    return notifier.start_background_checking(interval_hours)

# Integration with Streamlit (for web interface)
def create_streamlit_update_component(updater):
    """Create Streamlit component for update notifications"""
    try:
        import streamlit as st
        
        # Check for updates
        if 'update_checked' not in st.session_state:
            st.session_state.update_checked = False
            st.session_state.update_available = False
            st.session_state.update_info = None
        
        if not st.session_state.update_checked:
            with st.spinner("Checking for updates..."):
                try:
                    if updater.check_internet_connection():
                        update_info = updater.check_for_updates(show_no_updates=False)
                        if update_info:
                            st.session_state.update_available = True
                            st.session_state.update_info = update_info
                        st.session_state.update_checked = True
                except Exception:
                    st.session_state.update_checked = True
        
        # Show update notification
        if st.session_state.update_available and st.session_state.update_info:
            update_info = st.session_state.update_info
            
            st.info(f"🔄 **Update Available: v{update_info['version']}**")
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**What's New:** {update_info['description'][:100]}...")
            
            with col2:
                if st.button("📥 Update Now", key="update_now"):
                    with st.spinner("Downloading and installing update..."):
                        try:
                            update_file = updater.download_update(update_info)
                            if update_file and updater.apply_update(update_file, update_info):
                                st.success("✅ Update installed! Please restart the application.")
                                st.balloons()
                            else:
                                st.error("❌ Update failed. Please try again.")
                        except Exception as e:
                            st.error(f"❌ Update error: {e}")
        
        elif st.session_state.update_checked and not st.session_state.update_available:
            st.success("✅ Application is up to date")
    
    except ImportError:
        pass  # Streamlit not available

if __name__ == "__main__":
    # Test the notification system
    print("Testing Update Notification System...")
    
    if UPDATER_AVAILABLE:
        from auto_updater import initialize_updater
        updater = initialize_updater("1.0.0")
        
        # Test notification
        notifier = UpdateNotifier(updater)
        
        # Simulate update available
        fake_update = {
            "version": "1.1.0",
            "name": "Test Update",
            "description": "This is a test update with new features and bug fixes."
        }
        
        print("Showing test notification...")
        notifier.show_system_notification(fake_update)
        
    else:
        print("❌ Auto-updater not available for testing")