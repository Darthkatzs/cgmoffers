#!/usr/bin/env python3
"""
Quotation System Startup Script
Starts both the web interface server and the quotation processing server.
"""

import subprocess
import sys
import os
import time
import threading

def start_web_server():
    """Start the web interface server on port 8000."""
    print("🌐 Starting web interface server on port 8000...")
    try:
        subprocess.run([sys.executable, "server.py", "8000"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Web server failed: {e}")
    except KeyboardInterrupt:
        print("🛑 Web server stopped by user")

def start_quotation_server():
    """Start the quotation processing server on port 8001."""
    print("⚙️  Starting quotation processing server on port 8001...")
    try:
        subprocess.run([sys.executable, "final_quotation_server.py", "8001"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Quotation server failed: {e}")
    except KeyboardInterrupt:
        print("🛑 Quotation server stopped by user")

def main():
    """Start both servers."""
    
    print("🚀 Starting Complete Quotation System")
    print("=" * 60)
    
    # Check required files
    required_files = [
        "server.py",
        "final_quotation_server.py",
        "index.html",
        "script.js",
        "style.css",
        "standaardofferte Compufit NL.docx"
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("❌ Missing required files:")
        for file in missing_files:
            print(f"   - {file}")
        print("\nPlease ensure all files are in the current directory.")
        return 1
    
    # Check python-docx
    try:
        import docx
        print("✅ python-docx library found")
    except ImportError:
        print("❌ python-docx library not found!")
        print("   Install it with: pip install python-docx")
        return 1
    
    print("\n🎯 System will start:")
    print("   📱 Web interface: http://localhost:8000")
    print("   ⚙️  Quotation API: http://localhost:8001")
    print("\n🛑 Press Ctrl+C to stop both servers")
    print("=" * 60)
    
    # Start both servers in separate threads
    web_thread = threading.Thread(target=start_web_server, daemon=True)
    quotation_thread = threading.Thread(target=start_quotation_server, daemon=True)
    
    try:
        web_thread.start()
        time.sleep(2)  # Give web server time to start
        quotation_thread.start()
        
        # Keep main thread alive
        while True:
            time.sleep(1)
            
    except KeyboardInterrupt:
        print("\n👋 Shutting down quotation system...")
        return 0

if __name__ == "__main__":
    exit(main())