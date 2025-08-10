
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
        print("\nApplication stopped by user.")
    except Exception as e:
        print(f"Error starting application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
