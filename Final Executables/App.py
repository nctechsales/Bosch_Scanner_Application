import tkinter as tk
from tkinter import messagebox, filedialog
import socket
import os
import time
import threading

CONFIG_FILE = 'C:\\NorthCoastTech\\Bosch\\Bosch_Two_Camera_Config.txt'
CONFIG_DIR = os.path.dirname(CONFIG_FILE)

# Controls turning on and off the server
CONTROL_PORT = 9999

# Controls the messages between the GUI and the server
GUI_PORT = 9998


class App:
    def __init__(self, master):
        self.master = master
        self.master.title("Bosch TCP Server GUI")

        self.ip_var = tk.StringVar()
        self.port_var = tk.StringVar()
        self.log_file_var = tk.StringVar()

        self.load_config()
        self.create_widgets()
        self.check_server_status()

        self.gui_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.gui_socket.bind(('localhost', GUI_PORT))
        self.gui_socket.listen(1)
        self.server_messages_thread = threading.Thread(target=self.receive_server_messages)
        self.server_messages_thread.start()

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.master.rowconfigure(0, weight=1)
        self.master.rowconfigure(1, weight=1)
        self.master.rowconfigure(2, weight=1)
        self.master.rowconfigure(3, weight=1)
        self.master.rowconfigure(4, weight=1)
        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)

        tk.Label(self.master, text="IP Address").grid(row=0, column=0, padx=10, pady=5, sticky="nsew")
        self.ip_entry = tk.Entry(self.master, textvariable=self.ip_var)
        self.ip_entry.grid(row=0, column=1, padx=10, pady=5, sticky="nsew")

        tk.Label(self.master, text="Port").grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.port_entry = tk.Entry(self.master, textvariable=self.port_var)
        self.port_entry.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")

        tk.Label(self.master, text="Log File Location").grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
        self.log_file_entry = tk.Entry(self.master, textvariable=self.log_file_var)
        self.log_file_entry.grid(row=2, column=1, padx=10, pady=5, sticky="nsew")
        self.log_file_button = tk.Button(self.master, text="Browse", command=self.browse_log_file)
        self.log_file_button.grid(row=2, column=2, padx=10, pady=5, sticky="nsew")

        self.SERVER_status_box = tk.Label(self.master, text="Server Status", bg="grey", width=20, height=5)
        self.SERVER_status_box.grid(row=3, column=0, columnspan=1, pady=10, padx=10, sticky="nsew")

        self.STATE_status_box = tk.Canvas(self.master, bg="grey", width=20, height=5)
        self.STATE_status_box.grid(row=3, column=1, columnspan=1, pady=10, padx=10, sticky="nsew")
        self.STATE_status_box.create_rectangle(0, 0, 0, 0, fill="grey", tags="half_rect")
        self.STATE_status_box.bind("<Configure>", self.resize_half_rectangle)

        self.start_button = tk.Button(self.master, text="Start Server", command=self.start_server)
        self.start_button.grid(row=4, column=0, pady=10, sticky="nsew")
        self.stop_button = tk.Button(self.master, text="Stop Server", command=self.stop_server)
        self.stop_button.grid(row=4, column=1, pady=10, sticky="nsew")

    def browse_log_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.log_file_var.set(file_path)

    def resize_half_rectangle(self, event):
        self.STATE_status_box.coords("half_rect", 0, 0, self.STATE_status_box.winfo_width() / 2, self.STATE_status_box.winfo_height())

    def disable_inputs(self):
        self.ip_entry.config(state='disabled')
        self.port_entry.config(state='disabled')
        self.log_file_entry.config(state='disabled')
        self.log_file_button.config(state='disabled')
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')

    def enable_inputs(self):
        self.ip_entry.config(state='normal')
        self.port_entry.config(state='normal')
        self.log_file_entry.config(state='normal')
        self.log_file_button.config(state='normal')
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')
    
    def update_server_status(self, status):
        colors = {
            "Running": "blue",
            "Stopped": "grey"
        }
        self.SERVER_status_box.config(bg=colors.get(status, "grey"))
        self.SERVER_status_box.config(text=status)

    def update_state_status(self, status):
        colors = {
            "Success": "#0bfb00",
            "Failure": "red",
            "Bad Scan": "yellow",
            "Barcode Scanned": "#0bfb00"
        }
        self.STATE_status_box.delete("half_rect")
        print(f"status : {status}")
        if status == "Barcode Scanned":
            self.STATE_status_box.config(bg="grey")
            self.STATE_status_box.create_rectangle(0, 0, self.STATE_status_box.winfo_width() / 2, self.STATE_status_box.winfo_height(), fill=colors.get(status, "grey"), tags="half_rect")
        else:
            self.STATE_status_box.config(bg=colors.get(status, "grey"))
            self.STATE_status_box.create_rectangle(0, 0, self.STATE_status_box.winfo_width(), self.STATE_status_box.winfo_height(), fill=colors.get(status, "grey"), tags="half_rect")
        self.STATE_status_box.update_idletasks()
    
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                ip_address = f.readline().strip()
                port = f.readline().strip()
                log_file = f.readline().strip()
                self.ip_var.set(ip_address)
                self.port_var.set(port)
                self.log_file_var.set(log_file)
        else:
            self.ip_var.set('192.168.0.116')
            self.port_var.set('25251')
            self.log_file_var.set('C:\\NorthCoastTech\\Bosch\\messages_log.xlsx')
            
    def save_config(self, ip_address, port, log_file):
        with open(CONFIG_FILE, 'w') as f:
            f.write(f"{ip_address}\n")
            f.write(f"{port}\n")
            f.write(f"{log_file}\n")
    
    def start_server(self):
        ip_address = self.ip_var.get()
        port = self.port_var.get()
        log_file = self.log_file_var.get()

        if not ip_address or not port or not log_file:
            messagebox.showwarning("Input Error", "IP Address, Port, and Log File Location are required")
            return

        try:
            port = int(port)
        except ValueError:
            messagebox.showwarning("Input Error", "Port must be an integer")
            return

        self.save_config(ip_address, port, log_file)

        self.send_command(f'START,{ip_address},{port},{log_file}')
        self.update_server_status("Running")
        self.disable_inputs()
    
    def stop_server(self):
        self.send_command('STOP')
        self.update_server_status("Stopped")
        self.enable_inputs()
        time.sleep(1)
    
    def check_server_status(self):
        try:
            print('Checking server status...')
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.connect(('localhost', CONTROL_PORT))
                s.sendall(b'STATUS')
                print('Sent status message... waiting...')
                response = s.recv(1024).decode('utf-8')
                print(f'Response: {response}')
                if response == 'RUNNING':
                    self.update_server_status("Running")
                    self.disable_inputs()
                else:
                    self.update_server_status("Stopped")
                    self.enable_inputs()
        except ConnectionRefusedError:
            self.update_server_status("Stopped")
            self.enable_inputs()
    
    def send_command(self, command):
        for _ in range(5):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.connect(('localhost', CONTROL_PORT))
                    s.sendall(command.encode('utf-8'))
                    response = s.recv(1024)
                    print(response.decode('utf-8'))
                return
            except ConnectionRefusedError:
                print("Waiting for server to be ready...")
                time.sleep(1)
        messagebox.showerror("Connection Error", "Failed to connect to the server.")
        
    def listen_for_server_messages(self):
        self.server_messages_thread = threading.Thread(target=self.receive_server_messages)
        self.server_messages_thread.daemon = True
        self.server_messages_thread.start()

    def receive_server_messages(self):
        while True:
            try:
                conn, _ = self.gui_socket.accept()
                with conn:
                    message = conn.recv(1024).decode('utf-8')
                    if message.startswith('Success'):
                        self.update_state_status('Success')
                    elif message.startswith('Failure'):
                        self.update_state_status('Failure')
                    elif message.startswith('Bad Scan'):
                        self.update_state_status('Bad Scan')
                    elif message.startswith('Barcode Scanned'):
                        self.update_state_status('Barcode Scanned')
            except Exception as e:
                print(f"Error receiving server message: {e}")
                break

    def on_closing(self):
        # Perform cleanup
        try:
            self.gui_socket.close()
        except Exception as e:
            print(f"Error closing GUI socket: {e}")
        
        self.master.destroy()
        os._exit(0)  # Ensure all threads and sockets are closed

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
