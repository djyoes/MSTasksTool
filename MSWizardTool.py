import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import winreg
import os
from PIL import Image, ImageTk

# Define a function to execute system commands
def run_command(command, description):
    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True, shell=True)
        status = f"Success: {description} completed." if result.returncode == 0 else f"Error: {description} failed."
        log = f"Command: {command}\nOutput:\n{result.stdout.strip()}\nError:\n{result.stderr.strip()}\n"
    except subprocess.CalledProcessError as e:
        status = f"Error: {description} failed."
        log = f"Command: {command}\nError:\n{e.stderr.strip()}\n"
    return status, log

# Task functions
def add_work_school_account():
    try:
        command = "start ms-settings:workplace"
        subprocess.run(command, check=True, shell=True)
        status = "Success: Opened 'Access Work or School' settings. Please add your account manually."
        log = f"Command: {command}\nAction: Opened 'Access Work or School' settings."
    except Exception as e:
        status = f"Error: Unable to open 'Access Work or School' settings. {e}"
        log = f"Command: {command}\nError: {e}"
    update_info(status)
    logs.append(log)
    update_status("1", "Done")

def clear_outlook_profiles():
    try:
        reg_path = r"Software\\Microsoft\\Office\\16.0\\Outlook\\Profiles"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
            winreg.DeleteKey(key, "")
        status = "Success: Outlook profiles cleared."
    except FileNotFoundError:
        status = "Info: No Outlook profiles found."
    except Exception as e:
        status = f"Error: Unable to clear Outlook profiles. {e}"
    log = f"Operation: Clear Outlook Profiles\nStatus: {status}"
    update_info(status)
    logs.append(log)
    update_status("2", "Done")

def clear_edge():
    try:
        reg_path = r"Software\\Microsoft\\Edge"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
            winreg.DeleteKey(key, "")
        status = "Success: Edge data cleared."
    except FileNotFoundError:
        status = "Info: No Edge data found."
    except Exception as e:
        status = f"Error: Unable to clear Edge data. {e}"
    log = f"Operation: Clear Edge Data\nStatus: {status}"
    update_info(status)
    logs.append(log)
    update_status("3", "Done")

def set_edge_default():
    try:
        # Open the default apps settings
        os.system("start ms-settings:defaultapps")
        
        status = "Info: Default Apps settings opened. Please follow these steps to set Edge as the default browser:\n\n" \
                 "1. Scroll down and click on 'Web browser' under 'Web browser'\n" \
                 "2. Select 'Microsoft Edge' from the list of available browsers\n" \
                 "3. Close the Default Apps settings window"
        
        # Load and resize the screenshot image
        screenshot_path = "path/to/your/screenshot.png"  # Replace with the actual path to your screenshot image
        screenshot = Image.open(screenshot_path)
        screenshot_width = 400  # Adjust the width as needed
        screenshot_height = int(screenshot_width * screenshot.height / screenshot.width)
        screenshot = screenshot.resize((screenshot_width, screenshot_height), Image.ANTIALIAS)
        screenshot = ImageTk.PhotoImage(screenshot)
        
        # Display the screenshot below the instructions
        info_text.configure(state="normal")
        info_text.delete(1.0, tk.END)
        info_text.insert(tk.END, status)
        info_text.image_create(tk.END, image=screenshot)
        info_text.configure(state="disabled")
        
        log = "Operation: Set Edge as Default Browser\nStatus: Default Apps settings opened for manual configuration."
    except Exception as e:
        status = f"Error: Unable to open Default Apps settings. {e}"
        log = f"Operation: Set Edge as Default Browser\nError: {e}"
        
        info_text.configure(state="normal")
        info_text.delete(1.0, tk.END)
        info_text.insert(tk.END, status)
        info_text.configure(state="disabled")
    
    logs.append(log)
    update_status("4", "Done")


def setup_onedrive():
    command = "%LOCALAPPDATA%\\Microsoft\\OneDrive\\OneDrive.exe"
    status, log = run_command(command, "Setup OneDrive")
    update_info(status)
    logs.append(log)
    update_status("5", "Done")

def exit_app():
    root.destroy()

def update_info(info):
    info_text.configure(state="normal")
    info_text.delete(1.0, tk.END)
    info_text.insert(tk.END, info)
    info_text.configure(state="disabled")

def update_status(button_num, result):
    status_labels[int(button_num) - 1].configure(text=result)

def show_log_window():
    log_window = tk.Toplevel(root)
    log_window.title("Technical Log")
    log_window.geometry("900x700")

    log_text_window = tk.Text(log_window, wrap=tk.WORD, font=("Courier New", 10))
    log_text_window.pack(fill=tk.BOTH, expand=True)

    log_text_window.insert(tk.END, "\n\n".join(logs))
    log_text_window.configure(state="disabled")

# Create the GUI
root = tk.Tk()
root.title("Modernized Setup Wizard")

# Configure styles
style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Helvetica", 12), padding=10)
style.configure("TLabel", font=("Helvetica", 12))
style.configure("Header.TLabel", font=("Helvetica", 16, "bold"), foreground="white", background="#6a5acd")
style.configure("Info.TLabel", font=("Helvetica", 12, "bold"))
style.configure("TFrame", background="#e6e6fa")

# Set the window size
window_width = 1300
window_height = 800

# Center the window
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_cordinate = int((screen_width / 2) - (window_width / 2))
y_cordinate = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

# Add a header frame
header_frame = tk.Frame(root, bg="#e6e6fa", height=100)
header_frame.pack(fill=tk.X)

# Add a logo
try:
    logo = tk.PhotoImage(file="logo.png")  # Replace "logo.png" with your logo file
    logo_label = tk.Label(header_frame, image=logo, bg="#e6e6fa")
    logo_label.pack(side=tk.LEFT, padx=10, pady=10)
except tk.TclError:
    header_label = ttk.Label(header_frame, text="Modernized Setup Wizard", style="Header.TLabel")
    header_label.pack(fill=tk.X, padx=10, pady=10)

# Create a frame for buttons and status
content_frame = ttk.Frame(root, padding=20, style="TFrame")
content_frame.pack(fill=tk.BOTH, expand=True)

# Create a grid layout for buttons and statuses
button_status_frame = ttk.Frame(content_frame, style="TFrame")
button_status_frame.pack(side=tk.LEFT, padx=(20, 20))

# Add numbered buttons with a modern font
buttons = [
    ("1. Add Work Account", add_work_school_account),
    ("2. Clear Outlook Profile", clear_outlook_profiles),
    ("3. Reset Edge", clear_edge),
    ("4. Set Edge Default", set_edge_default),
    ("5. Setup OneDrive", setup_onedrive),
    ("6. Exit", exit_app),
]

status_labels = []
for i, (text, command) in enumerate(buttons):
    ttk.Button(button_status_frame, text=text, command=command, width=25).grid(row=i, column=0, pady=10, padx=5, sticky="w")
    status_label = ttk.Label(button_status_frame, text="", font=("Helvetica", 12), width=20, anchor="w")
    status_label.grid(row=i, column=1, pady=10, padx=5, sticky="w")
    status_labels.append(status_label)

# Add the logs button
logs_button = ttk.Button(button_status_frame, text="Logs", command=show_log_window, width=25)
logs_button.grid(row=len(buttons), column=0, pady=10, columnspan=2)

# Create a frame for info
info_frame = ttk.Frame(content_frame, style="TFrame")
info_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Add a label for info
info_label = ttk.Label(info_frame, text="Info:", style="Info.TLabel")
info_label.pack(anchor=tk.NW, pady=(0, 10))

# Add a styled text widget for displaying info
info_text = tk.Text(info_frame, wrap=tk.WORD, font=("Helvetica", 14), width=50, height=15, state="disabled", bg="#f8f8ff", relief=tk.GROOVE, bd=2)
info_text.pack(fill=tk.BOTH, expand=True)

# Add a footer
footer_frame = ttk.Frame(root, padding=10, style="TFrame")
footer_frame.pack(side=tk.BOTTOM, fill=tk.X)

footer_label = ttk.Label(footer_frame, text="Powered by Your Company", font=("Helvetica", 10), foreground="#6a5acd")
footer_label.pack()

# Initialize logs
logs = []

# Start the GUI event loop
root.mainloop()
