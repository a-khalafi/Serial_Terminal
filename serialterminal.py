import serial
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import re
import logging
from logging.handlers import RotatingFileHandler
import json
import serial.tools.list_ports
from queue import Queue
import platform
import winreg
import tkinter.font as tkfont
import subprocess
import shlex
import os
import sys
from pathlib import Path
import time
import win32event
import win32api
import winerror

# === prevent dublicate Startup ===
mutex = win32event.CreateMutex(None, False, "Global\\MyUniqueBarcodeAppMutex")
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    print("Another instance is already running.")
    sys.exit(0)
    
# === CONFIGURATION ===
DEFAULT_PORT = "COM1"
DEFAULT_BAUDRATE = 9600
TIMEOUT = 2.0
EXPECTED_PORTS = ["COM1", "COM4"]
SAVE_FILE = "saved_commands.json"
# Get the user's default Documents folder
documents_path = Path.home() / "Documents"
# Create your app's folder inside Documents
app_folder = documents_path / "Serial Terminal"
app_folder.mkdir(parents=True, exist_ok=True)

# Define paths to your files inside that folder
LOG_FILE = app_folder / "serial_terminal.log"
COMMANDS_FILE = app_folder / "saved_commands.json"

# === Logging Setup ===
# Set up rotating log file handler (e.g., 1 files of 50MB each)

handler = RotatingFileHandler(LOG_FILE, maxBytes=50_000_000, backupCount=1)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

handler.setFormatter(formatter)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(handler)

# === Helper: Get current time string ===
def timestamp():
    return datetime.now().strftime("%H:%M:%S")

# === Helper: Get all COM ports from Windows registry (Windows only) ===
def get_registry_com_ports():
    com_ports = []
    if platform.system() != "Windows":
        return com_ports
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DEVICEMAP\SERIALCOMM")
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
                com_ports.append(value)
                logging.debug(f"Registry SERIALCOMM: {name} = {value}")
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\com0com\Parameters")
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, i)
                    if name.startswith("PortName"):
                        com_ports.append(value)
                        logging.debug(f"com0com registry: {name} = {value}")
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except Exception as e:
            logging.warning(f"com0com registry scan failed: {e}")
    except Exception as e:
        logging.warning(f"Failed to access registry for COM ports: {e}")
    return sorted(set(com_ports))

ansi_escape = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')


def clean_line(text):
    logging.debug(f"Raw input: {repr(text)}")
    text = ansi_escape.sub('', text)
    text = text.replace('\r', '').replace('\x00', '').strip()
    logging.debug(f"Cleaned: {repr(text)}")
    return text


class SerialTerminal:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Terminal")
        self.ser = None
        self.reader_thread = None
        self.response_queue = Queue()
        self.command_queue = Queue()
        self.running = False
        self.setup_gui()
        self.update_ports()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.loaded_commands = False
        self.load_saved_commands()
        self.loaded_commands = True
        self.hub4com_proc = None
        self.com2tcp_proc = None  

##        self.auto_save_commands()

    def on_closing(self):
        self.running = False  # Stop reader thread
        if self.reader_thread and self.reader_thread.is_alive():
            self.reader_thread.join(timeout=2)  # Wait up to 2 sec to stop
        self.terminate_routing()  # Stop hub4com if running
        self.save_saved_commands()
        self.root.destroy()


    def setup_gui(self):
        # Use a frame to hold left and right side
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left frame: config + output + input
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Config Frame
        config_frame = ttk.Frame(left_frame)
        config_frame.pack(padx=5, pady=5, fill=tk.X)

        ttk.Label(config_frame, text="Port:").pack(side=tk.LEFT)
        self.port_var = tk.StringVar(value=DEFAULT_PORT)
        self.port_menu = ttk.Combobox(config_frame, textvariable=self.port_var, state="normal",width=8)
        self.port_menu.pack(side=tk.LEFT, padx=5)
        self.port_menu.bind("<Double-Button-1>", self.scan_ports)

        ttk.Button(config_frame, text="Scan Ports", command=self.scan_ports).pack(side=tk.LEFT, padx=5)

        ttk.Label(config_frame, text="Baud Rate:").pack(side=tk.LEFT,padx=3)
        self.baud_var = tk.StringVar(value=str(DEFAULT_BAUDRATE))
        self.baud_menu = ttk.Combobox(
            config_frame, textvariable=self.baud_var, state="readonly",
            values=["300", "1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"],
            width=8 
        )
        self.baud_menu.pack(side=tk.LEFT, padx=5)

        ttk.Label(config_frame, text="Parity:").pack(side=tk.LEFT)
        self.parity_var = tk.StringVar(value="None")
        self.parity_menu = ttk.Combobox(
            config_frame, textvariable=self.parity_var, state="readonly",
            values=["None", "Even", "Odd", "Mark", "Space"],
            width=8
        )
        self.parity_menu.pack(side=tk.LEFT, padx=5)

        self.connect_button = ttk.Button(config_frame, text="Connect", command=self.toggle_connect)
        self.connect_button.pack(side=tk.LEFT, padx=5)

        # Set smaller font about 85% of default for output text and saved commands
        default_font = tkfont.nametofont("TkTextFont")
        smaller_font = (default_font.actual("family"), max(10, int(default_font.actual("size") * 0.85)))

        self.output_text = scrolledtext.ScrolledText(left_frame, height=20, width=20, state='disabled', font=smaller_font)
        self.output_text.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        input_frame = ttk.Frame(left_frame)
        input_frame.pack(padx=5, pady=5, fill=tk.X)
        self.command_var = tk.StringVar()
        self.command_entry = ttk.Entry(input_frame, width=20, textvariable=self.command_var)
        self.command_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.command_entry.bind("<Return>", self.send_command)
        ttk.Button(input_frame, text="Send", command=self.send_command).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Clear Output", command=self.clear_output).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Save Log", command=self.save_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Send File", command=self.send_commands_from_file).pack(side=tk.LEFT, padx=5)

        # Right frame: saved commands buttons and entries
        self.saved_cmd_frame = ttk.Frame(main_frame)
        self.saved_cmd_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

        self.saved_commands = [tk.StringVar() for _ in range(10)]

        smaller_font_entry = (default_font.actual("family"), max(9, int(default_font.actual("size") * 0.85)))

        ttk.Label(self.saved_cmd_frame, text="Saved Commands:").grid(row=0, column=0, pady=(0,5))

        self.saved_buttons = []
        
        # Frame for COM port routing (hub4com)
        self.com2tcp_port = tk.StringVar()
        ttk.Label(self.saved_cmd_frame, text="Routing (hub4com-com2tcp):").grid(row=12, column=0, columnspan=1, pady=(1, 2))

        self.com_dest1 = tk.StringVar()
        self.com_dest2 = tk.StringVar()
        self.com_dest3 = tk.StringVar()


        ports = get_registry_com_ports()
        port_choices = [""] + ports

        ttk.Label(self.saved_cmd_frame, text="Source:").grid(row=13, column=0, sticky=tk.W)
        self.com_source = tk.StringVar()
        self.source_menu = ttk.Combobox(self.saved_cmd_frame, textvariable=self.com_source, values=port_choices, width=10, state="readonly")
        self.source_menu.grid(row=13, column=1)
        if "COM1" in ports:
            self.com_source.set("COM1")


        ttk.Label(self.saved_cmd_frame, text="Dest 1:").grid(row=14, column=0, sticky=tk.W)
        self.dest1_menu = ttk.Combobox(self.saved_cmd_frame, textvariable=self.com_dest1, values=port_choices, width=10, state="readonly")
        self.dest1_menu.grid(row=14, column=1)

        ttk.Label(self.saved_cmd_frame, text="Dest 2:").grid(row=15, column=0, sticky=tk.W)
        self.dest2_menu = ttk.Combobox(self.saved_cmd_frame, textvariable=self.com_dest2, values=port_choices, width=10, state="readonly")
        self.dest2_menu.grid(row=15, column=1)

        ttk.Label(self.saved_cmd_frame, text="Dest 3 (Telnet)").grid(row=16, column=0, sticky=tk.W)
        self.dest3_menu = ttk.Combobox(self.saved_cmd_frame, textvariable=self.com_dest3, values=port_choices, width=10, state="readonly")
        self.dest3_menu.grid(row=16, column=1)

        self.com2tcp_menu = ttk.Combobox(self.saved_cmd_frame, textvariable=self.com2tcp_port, values=port_choices, width=10, state="readonly")
        self.com2tcp_menu.grid(row=16, column=0, padx=(1, 0),sticky=tk.E)


        ttk.Button(self.saved_cmd_frame, text="Start Routing", command=self.run_hub4com).grid(row=17, column=0, pady=(5, 0), sticky="w")
        ttk.Button(self.saved_cmd_frame, text="Stop Routing", command=self.terminate_routing).grid(row=17, column=1, pady=(2, 0), sticky="e")
        for i in range(10):
            entry = ttk.Entry(self.saved_cmd_frame, textvariable=self.saved_commands[i], width=25, font=smaller_font_entry)
            entry.grid(row=i+1, column=0, padx=5, pady=2)
            btn = ttk.Button(self.saved_cmd_frame, text=f"Send {i+1}", width=8, command=lambda i=i: self.send_saved_command(i))
            btn.grid(row=i+1, column=1, padx=5, pady=2)
            self.saved_buttons.append(btn)

        self.status_var = tk.StringVar(value="Disconnected")
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, padx=5, pady=2)
        # led indicating routing is running
        self.led_label = tk.Label(self.saved_cmd_frame, text="●", font=("Arial", 18))
        self.led_label.grid(row=12, column=1, padx=10, sticky="ew")  # adjust position
        self.update_routing_led(False)  # Start as red


    def save_saved_commands(self):
        try:
            data = [cmd.get() for cmd in self.saved_commands]
            with open(COMMANDS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.log_output(f"[Error] Failed to save commands: {e}")

    def load_saved_commands(self):
        if COMMANDS_FILE.exists():
            try:
                with open(COMMANDS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for i, cmd in enumerate(data):
                    if i < len(self.saved_commands):
                        self.saved_commands[i].set(cmd)
            except Exception as e:
                self.log_output(f"[Error] Failed to load saved commands: {e}")




    def update_ports(self):
        try:
            ports_info = list(serial.tools.list_ports.comports())
            ports = [port.device for port in ports_info]

            # Add registry ports if not already in the list
            registry_ports = get_registry_com_ports()
            for port in registry_ports:
                if port not in ports:
                    ports.append(port)

            # Sort ports numerically by COM number (e.g., COM1 < COM10)
            def port_sort_key(p):
                return int(p.replace("COM", "")) if p.startswith("COM") and p[3:].isdigit() else float('inf')

            display_ports = sorted(ports, key=port_sort_key)

            # Update all drop-down menus except dest3
            dropdowns = [
                (self.port_menu, self.port_var),
                (self.source_menu, self.com_source),
                (self.dest1_menu, self.com_dest1),
                (self.dest2_menu, self.com_dest2),
            ]

            for menu, var in dropdowns:
                menu['values'] = display_ports if display_ports else ["No ports found"]
                if display_ports:
                    selected_port = display_ports[0]
                    var.set(selected_port)
                else:
                    var.set("No ports found")

            # For dest3, add blank option
            self.dest3_menu['values'] = [""] + (display_ports if display_ports else ["No ports found"])
            self.com_dest3.set("")  # Default to blank
            self.status_var.set("Ports updated")
            # For com2tcp, add blank option
            self.com2tcp_menu['values'] = [""] + (display_ports if display_ports else ["No ports found"])
            self.com2tcp_menu.set("")  # Default to blank
            self.status_var.set("Ports updated")

        except Exception as e:
            self.log_output(f"[Error] Port scan failed: {e}")
            self.status_var.set(f"Error: {e}")
     

    def scan_ports(self, event=None):
        self.update_ports()
        self.status_var.set("Ports scanned")


    def toggle_connect(self):
        if self.ser and self.ser.is_open:
            self.disconnect()
        else:
            self.connect()

    def connect(self):
        try:
            port = self.port_var.get().split(" - ")[0] if " - " in self.port_var.get() else self.port_var.get()
            baudrate = int(self.baud_var.get())
            parity = {
                "None": serial.PARITY_NONE,
                "Even": serial.PARITY_EVEN,
                "Odd": serial.PARITY_ODD,
                "Mark": serial.PARITY_MARK,
                "Space": serial.PARITY_SPACE
            }[self.parity_var.get()]
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                parity=parity,
                bytesize=serial.EIGHTBITS,
                stopbits=serial.STOPBITS_ONE,
                timeout=TIMEOUT
            )
            self.running = True
            self.reader_thread = threading.Thread(target=self.reader, daemon=True)
            self.reader_thread.start()
            self.root.after(100, self.process_queue)
            self.connect_button.configure(text="Disconnect")
            self.log_output(f"[{timestamp()}] Connected to {port} at {baudrate} baud")
            self.status_var.set(f"Connected to {port}")
            self.command_entry.focus()
        except Exception as e:
            self.log_output(f"[Error] Failed to connect: {e}")
            self.status_var.set(f"Error: {e}")

    def disconnect(self):
        if self.ser and self.ser.is_open:
            self.running = False
            self.ser.close()
            self.ser = None
            self.connect_button.configure(text="Connect")
            self.log_output(f"[{timestamp()}] Port closed")
            self.status_var.set("Disconnected")

    def reader(self):
        buffer = ""
        while self.running and self.ser and self.ser.is_open:
            try:
                chunk = self.ser.read(self.ser.in_waiting or 1)
                if not chunk:
                    continue
                decoded = chunk.decode(errors='ignore')
                if not decoded:
                    continue
                buffer += decoded
                if '\r' in buffer or '\n' in buffer:
                    lines = re.split(r'[\r\n]+', buffer)
                    for line in lines[:-1]:
                        cleaned = clean_line(line)
                        if not cleaned:
                            continue
                        try:
                            cmd = self.command_queue.get_nowait() if not self.command_queue.empty() else "Unknown"
                        except Queue.Empty:
                            cmd = "Unknown"
                        self.response_queue.put((cmd, cleaned))
                    buffer = lines[-1]
            except Exception as e:
                self.log_output(f"[Error] Reader error: {e}")
                self.status_var.set(f"Error: {e}")
                break

    def send_command(self, event=None):
        cmd = self.command_var.get().strip()
        if not cmd:
            return
        if cmd.lower() in ['exit', 'quit']:
            self.disconnect()
            self.root.quit()
            return
        if not self.ser or not self.ser.is_open:
            self.log_output("[Error] Not connected")
            self.status_var.set("Error: Not connected")
            return
        try:
            self.command_queue.put(cmd)
            self.ser.write((cmd + '\r\n').encode())
            self.ser.flush()
            self.log_output(f"[{timestamp()}] >> {cmd}")
            self.command_var.set("")
        except Exception as e:
            self.log_output(f"[Error] Failed to send: {e}")
            self.status_var.set(f"Error: {e}")

    def send_saved_command(self, index):
        command = self.saved_commands[index].get()
        if command:
            if not self.ser or not self.ser.is_open:
                self.log_output("[Error] Not connected")
                self.status_var.set("Error: Not connected")
                return
            try:
                self.command_queue.put(command)
                self.ser.write((command + '\r\n').encode())
                self.ser.flush()
                self.log_output(f"[{timestamp()}] >> {command}")
            except Exception as e:
                self.log_output(f"[Error] Failed to send: {e}")
                self.status_var.set(f"Error: {e}")

    def process_queue(self):
        try:
            last_cmd = None
            last_response = None
            while not self.response_queue.empty():
                cmd, response = self.response_queue.get_nowait()
                self.log_output(f"[{timestamp()}] << {response}")
                if cmd != "Unknown":
                    last_cmd = cmd
                    last_response = response
            if last_response:
                self.status_var.set(f"Response: {last_response}")
        except Queue.Empty:
            pass
        if self.running:
            self.root.after(100, self.process_queue)

    def save_log(self):
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title="Save Log As"
            )
            if file_path:
                self.output_text.configure(state='normal')
                content = self.output_text.get(1.0, tk.END).strip()
                self.output_text.configure(state='disabled')
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.log_output(f"[{timestamp()}] Log saved to {file_path}")
                self.status_var.set(f"Log saved to {file_path}")
        except Exception as e:
            self.log_output(f"[Error] Failed to save log: {e}")
            self.status_var.set(f"Error: Failed to save log - {e}")

    def log_output(self, text):
        # If text already has timestamp pattern at start, don't add again
        if re.match(r"^\[\d{2}:\d{2}:\d{2}\]", text):
            line = text
        else:
            timestamp = datetime.now().strftime("%H:%M:%S")
            line = f"[{timestamp}] {text}"
        
        if hasattr(self, 'output_text'):
            self.output_text.configure(state='normal')
            self.output_text.insert(tk.END, line + "\n")
            self.output_text.see(tk.END)
            self.output_text.configure(state='disabled')
        
        try:
            logger.info(line)
        except Exception as e:
            print(f"[Error writing log] {e}")

    def clear_output(self):
        self.output_text.configure(state='normal')
        self.output_text.delete(1.0, tk.END)
        self.output_text.configure(state='disabled')
    def terminate_routing(self):
        self.update_routing_led(False)
        if self.hub4com_proc:
            try:
                if self.hub4com_proc.poll() is None:
                    self.hub4com_proc.terminate()
                    self.hub4com_proc.wait(timeout=2)
                    self.log_output("[Info] hub4com routing terminated.")
                    
                else:
                    self.log_output("[Info] hub4com already stopped.")
            except Exception as e:
                self.log_output(f"[Error] Failed to terminate hub4com: {e}")
            finally:
                self.hub4com_proc = None

    def verify_hub4com_started(self):
        if self.hub4com_proc.poll() is not None:
            output = self.hub4com_proc.stdout.read().decode(errors="ignore")
            self.log_output(f"[hub4com Error] {output.strip()}")
            self.status_var.set("hub4com failed to start.")
            self.update_routing_led(False)
            self.hub4com_proc = None
        else:
            self.update_routing_led(True)
            self.status_var.set("hub4com routing started.")
    def verify_com2tcp_started(self):

        if self.com2tcp_proc and self.com2tcp_proc.poll() is not None:
            output = self.com2tcp_proc.stdout.read().decode(errors="ignore")
            self.log_output(f"[com2tcp Error] {output.strip()}")
            self.status_var.set("com2tcp failed to start.")
            self.com2tcp_proc = None
        else:
            self.status_var.set("com2tcp started (telnet on port 23).")
            self.log_output("[Info] com2tcp started.")


    def run_hub4com(self):
        try:
            # Prevent launching if already running
            if self.hub4com_proc and self.hub4com_proc.poll() is None:
                self.log_output("[Info] hub4com is already running.")
                return

            source_port = self.com_source.get().strip()
            if not source_port:
                self.log_output("[Error] Source COM port not selected.")
                return

            dest_ports = []
            for var in [self.com_dest1, self.com_dest2, self.com_dest3]:
                val = var.get().strip()
                if val:
                    dest_ports.append(f"\\\\.\\{val}")

            if not dest_ports:
                self.log_output("[Error] At least one destination COM port is required.")
                return

            full_command = [
                r"C:\Program Files (x86)\com0com\hub4com.exe",
                "--baud=9600", "--octs=off", "--route=All:All",
                f"\\\\.\\{source_port}"
            ] + dest_ports

            # Hide console window on Windows
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            self.log_output(f"[Launching] {' '.join(full_command)}")

            self.hub4com_proc = subprocess.Popen(
                full_command,
                startupinfo=startupinfo,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT
            )

            self.root.after(500, self.verify_hub4com_started)

        except Exception as e:
            self.log_output(f"[Error] Failed to launch hub4com: {e}")
        # If Dest 3 is used, start com2tcp on that port
        dest3_port = self.com_dest3.get().strip()
        tcp_port = self.com2tcp_port.get().strip()
        if dest3_port:
            try:
                com2tcp_command = [
                    r"C:\Program Files (x86)\com0com\com2tcp.bat",
                    "--telnet",
                    f"\\\\.\\{tcp_port}",
                    "23"
                ]

                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                self.log_output(f"[Launching] {' '.join(com2tcp_command)}")

                self.com2tcp_proc = subprocess.Popen(
                    com2tcp_command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    startupinfo=startupinfo
                )

                self.root.after(500, self.verify_com2tcp_started)

            except Exception as e:
                self.log_output(f"[Error] Failed to launch com2tcp: {e}")

    def update_routing_led(self, is_running: bool):
        self.led_label.config(fg="green" if is_running else "red")
    def send_commands_from_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt")],
            title="Select Command File"
        )
        if file_path:
            threading.Thread(
                target=self._send_file_commands_thread,
                args=(file_path,),
                daemon=True
            ).start()

    def _send_file_commands_thread(self, file_path, delay_ms=500):
        if not self.ser or not self.ser.is_open:
            self.log_output("[Error] Not connected")
            self.status_var.set("Error: Not connected")
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()

            for line in lines:
                cmd = line.strip()
                if not cmd or cmd.startswith("#"):  # Allow comments
                    continue
                self.command_queue.put(cmd)
                self.ser.write((cmd + '\r\n').encode())
                self.ser.flush()
                self.log_output(f"[{timestamp()}] >> {cmd}")
                time.sleep(delay_ms / 1000.0)

            self.status_var.set("Finished sending file commands.")
        except Exception as e:
            self.log_output(f"[Error] Failed to send file: {e}")
            self.status_var.set(f"Error: {e}")
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    icon_path = app_folder / "app_icon.ico"
    try:
        root.iconbitmap(icon_path)  # Replace with icon path
    except Exception as e:
        print(f"Failed to set icon: {e}")
    app = SerialTerminal(root)
    app.run()

