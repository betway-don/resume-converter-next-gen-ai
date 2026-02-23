
import os
import sys
import time
import subprocess
import threading

import socket

def get_local_ip():
    try:
        # Connect to an external server (doesn't actually send data) to get the interface IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

def start_share():
    local_ip = get_local_ip()
    port = 8000
    
    print("\n" + "="*60)
    print("🚀 App is running locally!")
    print(f"🏠 Local Access:   http://localhost:{port}")
    if local_ip != "127.0.0.1":
        print(f"🔗 Network Access: http://{local_ip}:{port}")
        print("   (Share this 'Network Access' link with colleagues on the same Wi-Fi/VPN)")
    print("="*60 + "\n")

    # Start Uvicorn bound to 0.0.0.0 to allow external access
    # We use subprocess.run so it takes over the console and users can see logs directly
    subprocess.run([sys.executable, "-m", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", str(port)])

if __name__ == "__main__":
    start_share()
