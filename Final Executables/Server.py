import socket
import threading
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import os
import time
import sys

CONFIG_FILE = 'C:\\NorthCoastTech\\Bosch\\Bosch_Two_Camera_Config.txt'
CONFIG_DIR = os.path.dirname(CONFIG_FILE)

# Controls turning on and off the server.
CONTROL_PORT = 9999

# Controls sending information to the front end.
GUI_PORT = 9998

def load_config():
    print('config found...111')
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)
    if os.path.exists(CONFIG_FILE):
        print('config found...')
        with open(CONFIG_FILE, 'r') as f:
            ip_address = f.readline().strip()
            port = f.readline().strip()
            log_file = f.readline().strip() if f.readline().strip() else get_desktop_path('message_log.xlsx')
            return ip_address, int(port), log_file if log_file else get_desktop_path('message_log.xlsx')
    else:
        print('config not found')
        with open(CONFIG_FILE, 'w') as f:
            f.write(f"{'192.168.0.116'}\n")
            f.write(f"{'25251'}\n")
            f.write(f"{get_desktop_path('message_log.xlsx')}\n")
        return '192.168.0.116', 25251, get_desktop_path('message_log.xlsx')

def get_desktop_path(filename):
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
    if os.path.exists(desktop):
        return os.path.join(desktop, filename)
    else:
        print(desktop)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive\\Desktop')
        return os.path.join(desktop, filename)

class Server:
    def __init__(self, ip_address, port, log_file):
        self.running = False
        self.ip_address = ip_address
        self.port = port
        self.first_barcode = None
        self.last_barcode = None
        self.server_socket = None
        self.server_thread = None
        self.log_file = log_file
        self.setup_excel()
    
    def save_config(self, ip_address, port, log_file):
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, 'w') as f:
            f.write(f"{ip_address}\n")
            f.write(f"{port}\n")
            f.write(f"{log_file}\n")

    def setup_excel(self):
        try:
            self.workbook = openpyxl.load_workbook(self.log_file)
            self.sheet = self.workbook.active
        except FileNotFoundError:
            self.workbook = Workbook()
            self.sheet = self.workbook.active
            self.sheet.append(["IP Address", "Message", "Time", "Success/Failure"])
            self.workbook.save(self.log_file)
    
    def log_message(self, ip_address, message, status):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.sheet.append([ip_address, message, now, status])
        self.workbook.save(self.log_file)
        self.send_to_gui(f"{status},{message}")
        
    def start_server(self):
        if self.running:
            self.log_to_gui("Server is already running.")
            return
        
        self.running = True
        self.setup_server_socket()
        
        self.log_to_gui('Starting server...')
        self.server_thread = threading.Thread(target=self.accept_connections)
        self.server_thread.daemon = True
        self.server_thread.start()
        self.send_to_gui("RUNNING")

    def stop_server(self):
        if not self.running:
            self.log_to_gui("Server is not running.")
            return
        
        self.running = False
        self.log_to_gui('Closing server socket...')
        if self.server_socket:
            self.server_socket.close()
        self.log_to_gui("Server stopped.")
        self.send_to_gui("STOPPED")
        
    
    def setup_server_socket(self):
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_socket.bind((self.ip_address, self.port))
            self.server_socket.listen(5)
            self.log_to_gui(f"Listening on {self.ip_address}:{self.port}")
        except socket.error as e:
            self.log_to_gui(f"Failed to set up server socket: {e}")
            self.stop_server()

    def handle_client(self, client_socket, client_address):
        with client_socket:
            message = client_socket.recv(1024).decode('utf-8')
            self.log_to_gui(f"Received message: {message} from {client_address[0]}")
            code = message.split(',')[0]
            scanned_data = message.split(',')[1]
            if code.find("QRCODE") == -1:
                barcode = message.split(',')[1]
            if code.find("CODE39") == -1:
                barcode = message.split(',')[1][2:16]
            self.log_to_gui(f"Barcode value: {barcode}")
            if len(barcode) < 13:
                self.log_to_gui(f"Error: Badscan: {barcode}")
                self.log_message(client_address[0], scanned_data, "Bad Scan")
            else:
                if self.first_barcode is None and code.find("QRCODE") == -1:
                    self.log_to_gui("CODE39 scanned, waiting for QR code.")
                    self.first_barcode = barcode
                    self.log_to_gui(f"first barcode: {self.first_barcode}")
                    self.log_message(client_address[0], scanned_data, "Barcode Scanned")
                if self.first_barcode is None and code.find("CODE39") == -1:
                    self.log_to_gui("Error: Barcode not scanned and documented.")
                    self.log_message(client_address[0], scanned_data, "Failure: Barcode not scanned and documented. QR code scanned.")
                    
                if self.first_barcode is not None and code.find("CODE39") != -1 and self.first_barcode != barcode:
                    self.log_to_gui("Error: QR code not attached.")
                    self.log_message(client_address[0], scanned_data, "Failure: QR code not attached, process step skipped.")
                    self.first_barcode = barcode
                    self.log_to_gui(f"first barcode: {self.first_barcode}")
                if self.first_barcode is not None and code.find("CODE39") == -1:
                    if self.first_barcode == barcode:
                        self.log_to_gui("Success: Codes match.")
                        self.log_message(client_address[0], scanned_data, "Success")
                    else:
                        self.log_to_gui(f"{self.first_barcode}{barcode}")
                        self.log_to_gui("Error: Codes do not match")
                        self.log_message(client_address[0], scanned_data, f"Failure: QR code and barcode do not match. QR Code: {barcode} ")
                    self.last_barcode = self.first_barcode
                    self.first_barcode = None
    
    def accept_connections(self):
        self.log_to_gui('Accepting connections...')
        while self.running:
            try:
                client_socket, client_address = self.server_socket.accept()
                self.log_to_gui(f"Connection from {client_address}")
                threading.Thread(target=self.handle_client, args=(client_socket, client_address)).start()
            except socket.error as e:
                if self.running:
                    self.log_to_gui(f"Socket error: {e}")
    
    def send_to_gui(self, message):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.connect(('localhost', GUI_PORT))
                s.sendall(message.encode('utf-8'))
        except ConnectionRefusedError:
            self.log_to_gui("GUI is not running")

    def log_to_gui(self, message):
        print(message)  # Ensure the message is printed to the terminal
        self.send_to_gui(f"LOG,{message}")

def listen_for_commands(server):
    print('Listening for control commands...')
    control_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    control_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    control_socket.bind(('localhost', CONTROL_PORT))
    control_socket.listen(1)

    while True:
        try:
            conn, addr = control_socket.accept()
            with conn:
                command = conn.recv(1024).decode('utf-8')
                if command == 'STOP':
                    server.stop_server()
                elif command.startswith('START'):
                    print('Starting server...')
                    _, ip_address, port, log_file = command.split(',')
                    server.ip_address = ip_address
                    server.port = int(port)
                    server.log_file = log_file
                    server.save_config(ip_address, port, log_file)
                    server.start_server()
                elif command.startswith('STATUS'):
                    conn.sendall(b'RUNNING' if server.running else b'STOPPED')
        except socket.error as e:
            print(f"Control socket error: {e}")

if __name__ == "__main__":
    ip, port, log_file = load_config()
    server = Server(ip, port, log_file)
    
    control_thread = threading.Thread(target=listen_for_commands, args=(server,))
    control_thread.daemon = True
    control_thread.start()
    server.start_server()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        server.stop_server()
        print("Server shut down.")
