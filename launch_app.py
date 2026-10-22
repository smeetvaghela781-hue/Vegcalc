import os
import socket
import webbrowser
import http.server
import socketserver

PORT = 5174
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = '127.0.0.1'
    finally:
        s.close()
    return IP

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

if __name__ == '__main__':
    local_ip = get_ip()
    phone_url = f"http://{local_ip}:{PORT}/"
    pc_url = f"http://localhost:{PORT}/"
    
    print("=" * 60)
    print("  VEGCALC PRO — VEGETABLE INTAKE & PROFIT TRACKER")
    print("=" * 60)
    print(f"  PC Browser URL : {pc_url}")
    print(f"  Phone App URL  : {phone_url}")
    print("=" * 60)
    print("  ON YOUR ANDROID PHONE:")
    print(f"  1. Connect to the same Wi-Fi / Hotspot.")
    print(f"  2. Open Chrome and visit: {phone_url}")
    print("  3. Tap (⋮) in Chrome -> 'Install app' or 'Add to Home screen'.")
    print("  4. The app will install with its own icon on your mobile!")
    print("=" * 60)
    
    webbrowser.open(pc_url)
    
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")
