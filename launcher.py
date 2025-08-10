
import os
import sys
import subprocess
import webbrowser
import time
import socket
from pathlib import Path

# Import auto-updater
try:
    from auto_updater import initialize_updater, check_for_updates_on_startup
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False

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
    
    # Initialize auto-updater
    updater = None
    if UPDATER_AVAILABLE:
        try:
            # Get current version
            version_file = app_dir / "version.json"
            current_version = "1.0.0"
            if version_file.exists():
                import json
                with open(version_file, 'r') as f:
                    version_data = json.load(f)
                    current_version = version_data.get("version", "1.0.0")
            
            updater = initialize_updater(current_version)
            print("Auto-updater initialized successfully")
        except Exception as e:
            print(f"Auto-updater initialization failed: {e}")
    
    # Find free port
    port = find_free_port()
    
    print("=" * 60)
    print("Energy Adjustment Calculator - Offline Version")
    print("=" * 60)
    print(f"Version: {current_version if 'current_version' in locals() else '1.0.0'}")
    print(f"Starting application on port {port}...")
    print("The application will open in your default web browser.")
    if updater:
        print("Auto-update system: Enabled")
    print("To stop the application, close this window.")
    print("=" * 60)
    
    # Check for updates in background
    if updater:
        check_for_updates_on_startup(updater)
    
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
        print("\nApplication stopped by user.")
    except Exception as e:
        print(f"Error starting application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
