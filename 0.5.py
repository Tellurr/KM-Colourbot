# Version: 0.5
# [+] Support more than one mouse type and don't hardcode mouse connection information (line 21)
# [-] PROCESS_ZONE_SIZE is still hardcoded (line 41) and should be a configuration option (line 101)

# [!] Can confirm this is working with net

# [!] If you can connect to the B+ but will not move check line 155

# [%] Python is shit. Changing to c++ to improve speed and accuracy.This will no longer get updates.

from threading import Thread
import dxcam
import kmNet
import win32api
import numpy as np
import cv2
import time
import math
import serial
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import json
import os
import subprocess
import logging

# Setup logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

camera = dxcam.create()

SCREEN_WIDTH, SCREEN_HEIGHT = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
middle_x, middle_y = SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2

trigger_bot = 0x06
if type(trigger_bot) == str:
    trigger_bot = int(trigger_bot, 16)

lower_bound = np.array([190, 40, 190])
upper_bound = np.array([310, 160, 310])

speed = 7.42
offset_mode = "head"  # Add other offsets like chest, etc. (not needed for now)

thread_running = False

PROCESS_ZONE_SIZE = 40
GRAB_ZONE = (
    int(SCREEN_WIDTH / 2 - PROCESS_ZONE_SIZE / 2),
    int(SCREEN_HEIGHT / 2 - PROCESS_ZONE_SIZE / 2),
    int(SCREEN_WIDTH / 2 + PROCESS_ZONE_SIZE / 2),
    int(SCREEN_HEIGHT / 2 + PROCESS_ZONE_SIZE / 2),
)

class MouseNet:
    def __init__(self, ip, port, uid):
        self.ip = ip
        self.port = port
        self.uid = uid

    def connect(self):
        try:
            logging.info("Now trying to initialize")
            logging.warning("If infomation is incorrect this will crash the program")
            kmNet.init(self.ip, self.port, self.uid)
            logging.info("Initialized successfully.")
        except Exception as e:
            logging.error(f"Failed to initialize MouseNet: {e}")
            raise

#    Old method, not needed anymore
#
#    def move(self, x, y):
#        if self.ip:
#            kmNet.move(x, y)


    def click(self):
        if self.ip:
            kmNet.click(0)

    def ping(self):
        try:
            response = subprocess.run(
                ["ping", "-n", "4", self.ip], capture_output=True, text=True
            )
            if response.returncode == 0:
                logging.info("Ping successful.")
                return True
            logging.warning("Ping failed.")
            return False
        except Exception as e:
            logging.error(f"Ping error: {e}")
            return False

class MouseB:
    def __init__(self, com_port, bitrate):
        self.com_port = com_port
        self.bitrate = bitrate
        self.serial_connection = None

    def connect(self):
        try:
            self.serial_connection = serial.Serial(self.com_port, self.bitrate)
            logging.info("MouseB connected successfully.")
        except Exception as e:
            logging.error(f"Could not connect to MouseB: {e}")
            raise

    def move(self, x, y, steps):
        if self.serial_connection:
            try:
                command = f"km.move({x},{y},{steps})"
                self.serial_connection.write(command.encode())
            except Exception as e:
                logging.error(f"Failed to send move command: {e}")

    def click(self):
        if self.serial_connection:
            try:
                self.serial_connection.write("km.click(0)".encode())
            except Exception as e:
                logging.error(f"Failed to send click command: {e}")

    def close(self):
        if self.serial_connection:
            try:
                self.serial_connection.close()
            except Exception as e:
                logging.error(f"Failed to close MouseB connection: {e}")

mouse = None  # Placeholder for the selected mouse type

def smoothmove(x, y, speed):
    start_x, start_y = win32api.GetCursorPos()
    distance = math.sqrt(x**2 + y**2)
    num_steps = max(int(distance / (speed * 0.2)), 1)
    step_x = x / num_steps
    step_y = y / num_steps

    target_x = start_x + x
    target_y = start_y

    for i in range(num_steps):
        if win32api.GetKeyState(trigger_bot) < 0:
            try:
                new_x = int(start_x + (i + 1) * step_x)
                new_y = int(start_y + (i + 1) * step_y)
                current_x, current_y = win32api.GetCursorPos()
                move_x = new_x - current_x
                move_y = new_y - current_y
                if abs(new_x - target_x) <= 3 and abs(new_y - target_y) <= 3:
                    break
                if isinstance(mouse, MouseNet):
                    kmNet.move(move_x, move_y)
                elif isinstance(mouse, MouseB): # if this shit dosent work then go to line 110 to 111 and paste that shit in
                    mouse.move(move_x, move_y, steps=5)
                time.sleep(0.001)
            except Exception as e:
                logging.error(f"Smooth move failed: {e}")


def process_zone_detection(frame):
    try:
        mask = cv2.inRange(frame, lower_bound, upper_bound)
        color_present = cv2.countNonZero(mask) > 0
        coordinates = cv2.findNonZero(mask)

        if color_present and coordinates is not None:
            box_size = mask.shape[1]  # Estimate box size based on mask height
            scaling_factor = 1 if box_size > 27 else 0.3

            x_pos, y_pos = coordinates[0][0]
            x_pos = x_pos + GRAB_ZONE[0]
            y_pos = y_pos + GRAB_ZONE[1]

            if offset_mode == "head":
                y_pos -= int(box_size / 200 * scaling_factor)

            return True, x_pos, y_pos

        return False, None, None
    except Exception as e:
        logging.error(f"Zone detection failed: {e}")
        return False, None, None

def main(frame):
    color_present, x, y = process_zone_detection(frame)

    if color_present:
        try:
            move_x = x - middle_x
            move_y = y - middle_y
            smoothmove(move_x, move_y, speed)
        except Exception as e:
            logging.error(f"Error in main movement logic: {e}")

def threaded_capture():
    global thread_running
    while thread_running:
        if win32api.GetKeyState(trigger_bot) < 0:
            try:
                screenshot = camera.grab(region=GRAB_ZONE)
                if screenshot is None:
                    continue
                main(screenshot)
            except Exception as e:
                logging.error(f'Threaded capture error: {e}')

class ConfigForm:
    def __init__(self, root):
        self.root = root
        self.root.title("Setting")
        self.root.geometry("320x600")

        self.config_file = "config.json"
        self.config = self.load_config()

        self.mouse_type = ttk.StringVar(value=self.config.get("mouseType", "Net"))
        self.keycode = ttk.StringVar(value=self.config.get("keycode", ""))
        self.zone = ttk.StringVar(value=self.config.get("zone", "50"))
        self.speed = ttk.StringVar(value=self.config.get("speed", ""))
        self.ip = ttk.StringVar(value=self.config.get("ip", ""))
        self.port = ttk.StringVar(value=self.config.get("port", ""))
        self.uid = ttk.StringVar(value=self.config.get("uid", ""))
        self.com_port = ttk.StringVar(value=self.config.get("comPort", ""))
        self.bitrate = ttk.StringVar(value=self.config.get("bitrate", ""))

        self.thread_running = False
        self.create_widgets()

    def create_widgets(self):
        # Mouse Type 
        mouse_type_label = ttk.Label(self.root, text="Mouse Type:")
        mouse_type_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        mouse_type_combo = ttk.Combobox(self.root, textvariable=self.mouse_type, values=["Net", "B"])
        mouse_type_combo.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        mouse_type_combo.bind("<<ComboboxSelected>>", self.update_fields)

        # Keycode
        keycode_label = ttk.Label(self.root, text="Keycode:")
        keycode_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        keycode_entry = ttk.Entry(self.root, textvariable=self.keycode)
        keycode_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Speed
        speed_label = ttk.Label(self.root, text="Speed:")
        speed_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.speed_entry = ttk.Entry(self.root, textvariable=self.speed)
        self.speed_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")


        # Net Fields
        self.ip_label = ttk.Label(self.root, text="IP:")
        self.ip_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.ip_entry = ttk.Entry(self.root, textvariable=self.ip)
        self.ip_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        self.port_label = ttk.Label(self.root, text="Port:")
        self.port_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.port_entry = ttk.Entry(self.root, textvariable=self.port)
        self.port_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        self.uid_label = ttk.Label(self.root, text="UID:")
        self.uid_label.grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.uid_entry = ttk.Entry(self.root, textvariable=self.uid)
        self.uid_entry.grid(row=6, column=1, padx=10, pady=5, sticky="ew")

        # B Fields
        self.com_port_label = ttk.Label(self.root, text="Com Port:")
        self.com_port_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.com_port_entry = ttk.Entry(self.root, textvariable=self.com_port)
        self.com_port_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        self.bitrate_label = ttk.Label(self.root, text="Bitrate:")
        self.bitrate_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.bitrate_entry = ttk.Entry(self.root, textvariable=self.bitrate)
        self.bitrate_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        # Zone
        zone_label = ttk.Label(self.root, text="Zone:")
        zone_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.zone_entry = ttk.Entry(self.root, textvariable=self.zone)
        self.zone_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Save Button
        save_button = ttk.Button(self.root, text="Save Configuration", command=self.save_config)
        save_button.grid(row=7, column=0, columnspan=2, padx=10, pady=20)

        # Add Toggle Switch
        switch_frame = ttk.LabelFrame(self.root, text="Control", padding=10)
        switch_frame.grid(row=10, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        self.toggle_switch = ttk.Checkbutton(
            switch_frame,
            text="Enable/Disable",
            style='Switch.TCheckbutton',
            command=self.toggle_thread
        )
        self.toggle_switch.pack(padx=5, pady=5)

        # Refresh Button
        refresh_button = ttk.Button(self.root, text="Refresh Configuration", command=self.refresh_config)
        refresh_button.grid(row=8, column=0, columnspan=2, padx=10, pady=10)

        # Mouse Connection Button
        connect_button = ttk.Button(self.root, text="Connect Mouse", command=self.connect_mouse)
        connect_button.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

        # Initial visibility update
        self.update_fields()

    def update_fields(self, event=None):
        if self.mouse_type.get() == "Net":
            self.ip_label.grid()
            self.ip_entry.grid()
            self.port_label.grid()
            self.port_entry.grid()
            self.uid_label.grid()
            self.uid_entry.grid()
            self.com_port_label.grid_remove()
            self.com_port_entry.grid_remove()
            self.bitrate_label.grid_remove()
            self.bitrate_entry.grid_remove()
        else:
            self.ip_label.grid_remove()
            self.ip_entry.grid_remove()
            self.port_label.grid_remove()
            self.port_entry.grid_remove()
            self.uid_label.grid_remove()
            self.uid_entry.grid_remove()
            self.com_port_label.grid()
            self.com_port_entry.grid()
            self.bitrate_label.grid()
            self.bitrate_entry.grid()

    def toggle_thread(self):
        global thread_running
        thread_running = not thread_running
        if thread_running:
            Thread(target=threaded_capture).start()

    def connect_mouse(self):
        global mouse
        try:
            if self.mouse_type.get() == "Net":
                mouse = MouseNet(self.ip.get(), self.port.get(), self.uid.get())
                if not mouse.ping():
                    Messagebox.show_error("Could not connect to Net Mouse: Ping failed.")
                    return
                mouse.connect()
                Messagebox.show_info("Connected successfully.")

            elif self.mouse_type.get() == "B":
                mouse = MouseB(self.com_port.get(), self.bitrate.get())
                mouse.connect()
                Messagebox.show_info("Connected successfully.")
            else:
                Messagebox.show_error("Invalid mouse type selected.")
        except Exception as e:
            logging.error(f"Error while connecting mouse: {e}")
            Messagebox.show_error(f"Failed to connect mouse: {e}")

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as file:
                    return json.load(file)
            else:
                return {}
        except Exception as e:
            logging.error(f"Failed to load configuration: {e}")
            return {}

    def save_config(self):
        config = {
            "mouseType": self.mouse_type.get(),
            "keycode": self.keycode.get(),
            "speed": self.speed.get(),
            "zone": self.zone.get(),
        }
        if self.mouse_type.get() == "Net":
            config.update({
                "ip": self.ip.get(),
                "port": self.port.get(),
                "uid": self.uid.get(),
            })
        else:
            config.update({
                "comPort": self.com_port.get(),
                "bitrate": self.bitrate.get(),
            })

        try:
            with open(self.config_file, 'w') as file:
                json.dump(config, file, indent=4)
            Messagebox.show_info("Configuration saved successfully!")
        except Exception as e:
            logging.error(f"Failed to save configuration: {e}")
            Messagebox.show_error(f"Failed to save configuration: {e}")

    def refresh_config(self):
        try:
            self.config = self.load_config()
            self.mouse_type.set(self.config.get("mouseType", "Net"))
            self.keycode.set(self.config.get("keycode", ""))
            self.speed.set(self.config.get("speed", ""))
            self.ip.set(self.config.get("ip", ""))
            self.port.set(self.config.get("port", ""))
            self.uid.set(self.config.get("uid", ""))
            self.com_port.set(self.config.get("comPort", ""))
            self.bitrate.set(self.config.get("bitrate", ""))
            self.zone.set(self.config.get("zone", "50"))
            Messagebox.show_info("Configuration reloaded successfully!")
        except Exception as e:
            logging.error(f"Failed to refresh configuration: {e}")

# Main Application
if __name__ == "__main__":
    root = ttk.Window(themename="flatly")
    app = ConfigForm(root)
    root.mainloop()
