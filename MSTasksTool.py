import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import subprocess
import winreg as reg
import os
import threading
import logging
from PIL import Image, ImageTk
import time
import webbrowser
import re
import sys

# Constants
LOG_FILE = r"logs\MSTasksTool.log"
LOGO_PATH = "logo.png"
ADD_WORK_SCREENSHOT_PATH = "setwork.jpg"
ADD_WORK_SET_SCREENSHOT_PATH = "workset.jpg"
SET_EDGE_SCREENSHOT_PATH = "setedge.jpg"
EDGE_SET_SCREENSHOT_PATH = "edgeset.jpg"
ONEDRIVE_SET_SCREENSHOT_PATH = "onedriveset.jpg"
EDGE_CON_SCREENSHOT_PATH = "edgecon.jpg"
WINDOW_WIDTH = 1000
WINDOW_HEIGHT = 800
HEADER_BG_COLOR = "#6a5acd"
BUTTON_WIDTH = 25
INFO_BOX_BG_COLOR = "#ffffff"
FONT_FAMILY = "Consolas"
FONT_SIZE = 14
LARGE_FONT_SIZE = 18
BUTTON_PADY = 5
BUTTON_PADX = 5
FRAME_PADY = 20
FRAME_PADX = 20
TEXT_PADY = 10
TEXT_PADX = 10
STATUS_PADY = 2
EDGE_DEFAULT_BROWSER_REGISTRY = r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"
HYPERLINK_COLOR = "blue"
EDGE_SETTINGS_URL = "edge://settings/defaultBrowser"
WORK_ACCOUNT_REGISTRY = r"Software\Microsoft\Windows\CurrentVersion\WorkplaceJoin"
ONEDRIVE_REGISTRY = r"Software\Microsoft\OneDrive"
ONEDRIVE_ACCOUNT_REGISTRY = r"Software\Microsoft\OneDrive\Accounts\Business1"

# Logging setup
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

# Utility Functions
def execute_command(command, description):
    """Executes CMD commands and logs the result."""
    logging.debug(f"Executing command: {command} for {description}")
    try:
        result = subprocess.run(command, capture_output=True, text=True, shell=True, check=False)
        status = f"{description}"
        log = f"Command: {command}\nOutput:\n{result.stdout.strip()}\nError:\n{result.stderr.strip()}\n"
        logging.debug(f"Command successful: {status}")
        logs.append(log)
        return result
    except subprocess.CalledProcessError as e:
        status = f"Error: {description} failed."
        log = f"Command: {command}\nError: {e.stderr.strip()}\n"
        logging.error(f"Command failed: {status}")
        logs.append(log)
        return None

def threaded_task(task_func, status_label=None):
    """Run tasks in a thread to avoid UI blocking."""
    if status_label:
        status_label.config(text="Running...", fg="blue")
    progress_bar.start()
    if task_func in (show_logs, save_logs):
        threading.Thread(target=lambda: run_task(task_func), daemon=True).start()
    else:
        threading.Thread(target=lambda: run_task(task_func, status_label), daemon=True).start()


def run_task(task_func, status_label=None):
    """Executes the task and updates the status."""
    try:
      if status_label:
          task_func(status_label)
      else:
          task_func()
      logging.debug(f"Task {task_func.__name__} completed")
    except Exception as e:
      logging.error(f"Task {task_func.__name__} failed: {e}")
      if status_label:
          status_label.config(text="Error", fg="red")
    finally:
      progress_bar.stop()

def update_info(message, image_path=None, link_url=None):
    """Update the info display with text and optionally display a resized image and a hyperlink."""
    info_text.configure(state="normal", font=(FONT_FAMILY, FONT_SIZE))
    info_text.delete(1.0, tk.END)
    
    if link_url:
        message += f"\n\nCopy and Paste this URL into Edge: {EDGE_SETTINGS_URL}"
        info_text.insert(tk.END, message)
        info_text.insert(tk.END, "\n")  # New line
    else:
      info_text.insert(tk.END, message)
    
    if image_path and os.path.exists(image_path):
        try:
            image = Image.open(image_path)
            image.thumbnail((info_text.winfo_width() - 20, 400), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            info_text.image_create(tk.END, image=photo)
            info_text.image = photo  # Prevent garbage collection
        except Exception as e:
            logging.error(f"Error loading image: {e}")
            info_text.insert(tk.END, "\n[Image failed to load]")
    info_text.configure(state="disabled")

def show_logs():
    """Displays logs in a new window."""
    log_window = tk.Toplevel(root)
    log_window.title("Logs")
    log_window.geometry("700x500")
    log_text = scrolledtext.ScrolledText(log_window, wrap="word", font=("Courier New", 10))
    log_text.pack(fill="both", expand=True)
    log_text.insert(tk.END, "\n".join(logs))
    log_text.configure(state="disabled")
    log_window.protocol("WM_DELETE_WINDOW", lambda: on_log_window_close(log_window))

def on_log_window_close(log_window):
    """Handles the closing of the log window and resets the status label."""
    log_window.destroy()
    for label in status_labels:
        if label.cget("text") == "Running...":
            label.config(text="")

def save_logs():
    """Saves logs to a file."""
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
    if file_path:
        with open(file_path, "w") as f:
            f.write("\n".join(logs))
        messagebox.showinfo("Success", "Logs saved successfully.")
    logs.clear()

def exit_app(exit_code=None):
    """Exit the application."""
    if exit_code is not None:
        root.destroy()
        sys.exit(exit_code)
    else:
        root.destroy()

def is_work_account_set():
    """Checks if a work account is set using dsregcmd."""
    try:
        command = "dsregcmd /status"
        result = subprocess.run(command, capture_output=True, text=True, shell=True, check=False)
        output = result.stdout
        # Regex to find the work place tenant name
        tenant_match = re.search(r"WorkplaceTenantName\s*:\s*(.+)", output)
        if tenant_match:
            tenant_name = tenant_match.group(1).strip()
            logging.debug(f"Work Account is set, tenant name {tenant_name}")
            return tenant_name
        logging.debug(f"Work account not set dsregcmd")
        return ""
    except Exception as e:
        logging.error(f"Error checking work account with dsregcmd: {e}")
        return ""

def has_outlook_profiles():
    """Checks if any outlook profiles exists using the registry"""
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, OUTLOOK_PROFILES_REGISTRY) as key:
            _, subkeys, _ = reg.EnumValue(key, 0)
            if subkeys:
              logging.debug(f"Outlook profiles found")
              return True
        return False
    except FileNotFoundError:
        logging.debug("Registry key for outlook profiles not found.")
        return False
    except Exception as e:
        logging.error(f"Error checking outlook profiles registry: {e}")
        return False

def is_onedrive_setup():
    """Checks if OneDrive is setup by checking registry keys."""
    try:
      with reg.OpenKey(reg.HKEY_CURRENT_USER, ONEDRIVE_ACCOUNT_REGISTRY) as key:
        value, _ = reg.QueryValueEx(key, "ConfiguredTenantId")
        if value:
          logging.debug("Onedrive is setup")
          return True
      logging.debug(f"Onedrive is not setup")
      return False
    except FileNotFoundError:
        logging.debug("OneDrive registry key not found.")
        return False
    except Exception as e:
        logging.error(f"Error checking onedrive registry: {e}")
        return False


# Task Functions
def add_work_account(status_label=None):
    tenant_name = is_work_account_set()
    if tenant_name:
        update_info(
           f'Work or School account is already set your Tenant is "{tenant_name}".',
           ADD_WORK_SET_SCREENSHOT_PATH
           )
        logging.debug(f"log for add_work_account, work account is already set")
        if status_label:
            status_label.config(text="Done", fg="green")
        logs.append("Work or School account is already set.")
        return
    status = execute_command("start ms-settings:workplace", "Opened Work or School settings")
    update_info(f"{status}", ADD_WORK_SCREENSHOT_PATH)
    if status_label:
      status_label.config(text="Done", fg="green")
    logs.append("Add work account task completed")
    logging.debug(f"log for add_work_account")

def clear_outlook_profiles(status_label=None):
    if not has_outlook_profiles():
        status = "No email profiles found."
        update_info(status)
        if status_label:
           status_label.config(text="Done", fg="green")
        logs.append(status)
        logging.debug(f"log for clear_outlook_profiles, no profiles found")
        return
    
    try:
        reg_path = r"Software\Microsoft\Office\16.0\Outlook\Profiles"
        with reg.OpenKey(reg.HKEY_CURRENT_USER, reg_path, 0, reg.KEY_READ) as key:
          if reg.QueryInfoKey(key)[0] > 0:
            delete_registry_key_tree(reg.HKEY_CURRENT_USER, reg_path)
            status = "Success: Outlook profiles cleared."
            if status_label:
              status_label.config(text="Done", fg="green")
            logs.append("Clear outlook profiles task completed")
            logging.debug(f"log for clear_outlook_profiles")
          else:
            status = "No email profiles found."
            update_info(status)
            if status_label:
              status_label.config(text="Done", fg="green")
            logs.append("No email profiles found")
            logging.debug(f"log for clear_outlook_profiles, no profiles to clear")
    except PermissionError as e:
        status = f"Error: {e}"
        if status_label:
            status_label.config(text="Error", fg="red")
        logs.append(f"Clear outlook profiles task failed: {e}")
        logging.debug(f"log for clear_outlook_profiles with exception")
    except Exception as e:
        status = f"Error: {e}"
        logging.error(f"Error clearing outlook profiles: {e}")
        if status_label:
          status_label.config(text="Error", fg="red")
        logs.append(f"Error clearing outlook profiles: {e}")
        logging.debug(f"log for clear_outlook_profiles with error")
    update_info(status)


def clear_edge(status_label=None):
    """Closes Edge, deletes the user data, and then relaunches Edge for login."""
    update_info("Please sign in if prompted and click confirm continue.", EDGE_CON_SCREENSHOT_PATH)
    def edge_task():
        status = ""
        taskkill_output = execute_command("taskkill /IM msedge.exe /F", "Closed all instances of Edge")
        if taskkill_output:
           logging.debug(f"taskkill output error: {taskkill_output.stderr.strip()}" if taskkill_output.stderr.strip() else f"taskkill output: {taskkill_output.stdout.strip()}")

        rmdir_output = execute_command(f'rmdir /S /Q "%LOCALAPPDATA%\\Microsoft\\Edge\\User Data"', "Deleted Edge user data directories")
        if rmdir_output:
          if "The system cannot find the file specified" in rmdir_output.stderr:
             logging.debug(f"rmdir command output error: The system cannot find the file specified")
          elif "The process cannot access the file because it is being used by another process." in rmdir_output.stderr:
            logging.debug(f"rmdir command output error: The process cannot access the file because it is being used by another process")
          else:
              logging.debug(f"rmdir command output: {rmdir_output.stdout.strip()} and {rmdir_output.stderr.strip()}")

        edge_launch_status = ""
        try:
            if os.path.exists("C:\\Program Files\\Microsoft Edge\\Application\\msedge.exe"):
                 execute_command(r'start "" "C:\Program Files\Microsoft Edge\Application\msedge.exe"', "Launched Edge from program files folder")
                 logging.debug(f"Attempted to launch edge from C:\\Program Files\\Microsoft Edge\\Application\\msedge.exe")
            elif os.path.exists(r'%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe'):
                execute_command(r'start "" "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"', "Launched Edge from program files x86 folder")
                logging.debug(f"Attempted to launch edge from %ProgramFiles(x86)%\\Microsoft\\Edge\\Application\\msedge.exe")
            else:
              logging.debug("Edge not found in common directories")
        except Exception as e:
            logging.error(f"Error launching edge: {e}")
            
        
        status_list = subprocess.run('tasklist /FI "IMAGENAME eq msedge.exe"', capture_output=True, text=True, shell=True, check=False)
        if "msedge.exe" in status_list.stdout:
            status = "Success: Edge data cleared and Edge relaunched."
            if status_label:
              status_label.config(text="Done", fg="green")
            logs.append(status)
            logging.debug(f"log for clear_edge, edge relaunched successfully")
        else:
            status = "Error: Edge data cleared but Edge was not relaunched."
            if status_label:
              status_label.config(text="Error", fg="red")
            logs.append(status)
            logging.debug(f"log for clear_edge, edge failed to launch")

        if status_label:
           update_info(status, EDGE_CON_SCREENSHOT_PATH)
    threading.Thread(target=edge_task, daemon=True).start()


def is_edge_default():
        """Checks if Edge is the default browser using the registry."""
        try:
            with reg.OpenKey(reg.HKEY_CURRENT_USER, EDGE_DEFAULT_BROWSER_REGISTRY) as key:
                value, _ = reg.QueryValueEx(key, "ProgId")
                logging.debug(f"Current default browser ProgId: {value}")
                return "MSEdgeHTM" in value
        except FileNotFoundError:
            logging.debug("Registry key for default browser not found.")
            return False
        except Exception as e:
            logging.error(f"Error checking default browser registry: {e}")
            return False

def set_edge_default(status_label=None):
    """Opens Default Apps setting page, with instructions to set Edge as default."""
    logging.debug(f"log for set_edge_default, task started")
    try:
      if is_edge_default():
          logging.info("Edge is already the default browser.")
          if status_label:
              status_label.config(text="Done", fg="green")
          update_info(
              "Edge is already the default browser.",
              EDGE_SET_SCREENSHOT_PATH
          )
          logs.append("Edge was already the default browser.")
          logging.debug(f"log for set_edge_default, edge already default")
          return
      status = execute_command("start ms-settings:defaultapps", "Opened Default Apps settings")
      update_info(
        f"{status}\n\nFollow these steps:\n1. Search for Edge in the list of apps. \n2. Click 'Make Default' to set Edge as the default browser.",
        SET_EDGE_SCREENSHOT_PATH
        )
      if status_label:
        status_label.config(text="Checking", fg="blue")
      for _ in range(12):
          if is_edge_default():
              logging.info("Edge set as default browser successfully.")
              if status_label:
                status_label.config(text="Done", fg="green")
              update_info(
                "Edge set as default browser successfully.",
                EDGE_SET_SCREENSHOT_PATH
                )
              logs.append("Edge set to default.")
              logging.debug(f"log for set_edge_default, edge set to default")
              return
          time.sleep(5)
      logging.warning("Edge was not set to default within 60 seconds")
      if status_label:
          status_label.config(text="Failed", fg="red")
      update_info(
        "Edge was not set to default within 60 seconds, please click the button again if you want to try again." ,
        SET_EDGE_SCREENSHOT_PATH
        )
      logs.append("Edge was not set to default.")
      logging.debug(f"log for set_edge_default, edge was not set to default")

    except Exception as e:
      error_message = f"Error: Unable to open default apps settings. {e}"
      logging.error(error_message)
      if status_label:
            status_label.config(text="Error", fg="red")
      update_info(error_message)
      logs.append(f"Error unable to set edge to default, exception: {e}")
      logging.debug(f"log for set_edge_default, exception raised")

def setup_onedrive(status_label=None):
    """Launches OneDrive for initial setup."""
    try:
        if is_onedrive_setup():
          logging.info("Onedrive is already set up.")
          if status_label:
            status_label.config(text="Done", fg="green")
          update_info(
             "Onedrive is already setup.",
              ONEDRIVE_SET_SCREENSHOT_PATH
          )
          logs.append("Onedrive was already setup.")
          logging.debug(f"log for setup_onedrive, onedrive already setup")
          return
        # OneDrive executable path
        command = r'%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe'
        status = ""
        if os.path.exists("C:\\Program Files\\Microsoft OneDrive\\OneDrive.exe"):
            command = r'"C:\Program Files\Microsoft OneDrive\OneDrive.exe"'
            status = execute_command(command, "OneDrive setup initiated from program files")
        elif os.path.exists(r'%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe'):
            status = execute_command(command, "OneDrive setup initiated from localappdata")
        else:
            status = "Error: Unable to find the OneDrive executable."
            logging.error(status)
            update_info(status)
            if status_label:
               status_label.config(text="Error", fg="red")
            logs.append(status)
            return

        # Add required registry keys to configure onedrive
        execute_command(f'reg add "HKCU\\Software\\Microsoft\\OneDrive" /v SilentBusinessConfigCompleted /t REG_DWORD /d 1 /f', "Added SilentBusinessConfigCompleted registry key")
        execute_command(f'reg add "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Business1" /v ConfiguredTenantId /t REG_SZ /d "" /f', "Added ConfiguredTenantId registry key")
        execute_command(f'reg add "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Business1" /v LastKnownFolderBackupLaunchSource /t REG_DWORD /d 1 /f', "Added LastKnownFolderBackupLaunchSource registry key")
        execute_command(f'reg add "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Business1" /v LastKFMOptInSource /t REG_DWORD /d 1 /f', "Added LastKFMOptInSource registry key")
        execute_command(f'reg add "HKCU\\Software\\Microsoft\\OneDrive\\Accounts\\Business1" /v KfmFoldersProtectedNow /t REG_DWORD /d 3584 /f', "Added KfmFoldersProtectedNow registry key")

        # Update the info display
        update_info(f"{status}\n\nFollow the on-screen instructions to complete the OneDrive setup.")
        if status_label:
            status_label.config(text="Done", fg="green")
        logs.append("Setup Onedrive task completed.")
        logging.debug(f"log for setup_onedrive")
    except Exception as e:
        error_message = f"Error: Unable to initiate OneDrive setup. {e}"
        logging.error(error_message)
        if status_label:
            status_label.config(text="Error", fg="red")
        update_info(error_message)
        logs.append(f"Error setting up onedrive, exception {e}")
        logging.debug(f"log for setup_onedrive, exception raised")

# Registry Helper
def delete_registry_key_tree(key, subkey_path):
    """Recursively deletes a registry key and its subkeys."""
    try:
        with reg.OpenKey(key, subkey_path, 0, reg.KEY_ALL_ACCESS) as handle:
            while True:
                try:
                    subkey = reg.EnumKey(handle, 0)
                    delete_registry_key_tree(key, f"{subkey_path}\\{subkey}")
                except OSError:
                    break
        reg.DeleteKey(key, subkey_path)
    except FileNotFoundError:
        pass
    except Exception as e:
        logging.error(f"Failed to delete registry: {e}")
        raise  # Re-raise exception to be caught by run_task

# GUI Setup
root = tk.Tk()
root.title("Microsoft Post Migration Helper Wizard")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg="#f0f0f0")

# Header
header_frame = tk.Frame(root, bg=HEADER_BG_COLOR, height=80)
header_frame.pack(fill="x")

# Separate frame for the logo
logo_frame = tk.Frame(header_frame, bg="#f0f0f0")  # Neutral background for logo
logo_frame.pack(side="left", padx=TEXT_PADX, pady=TEXT_PADY)

try:
    # Load and resize the logo while maintaining its aspect ratio
    original_logo = Image.open(LOGO_PATH)
    logo_width, logo_height = original_logo.size
    new_height = 60  # Adjust the height as needed
    new_width = int((new_height / logo_height) * logo_width)  # Maintain aspect ratio
    logo = ImageTk.PhotoImage(original_logo.resize((new_width, new_height), Image.Resampling.LANCZOS))
    
    tk.Label(logo_frame, image=logo, bg="#f0f0f0").pack()
except Exception as e:
    logging.error(f"Error loading logo: {e}")
    tk.Label(logo_frame, text="LOGO", bg="#f0f0f0", font=(FONT_FAMILY, LARGE_FONT_SIZE, "bold")).pack()

# Separate frame for the header text
header_text_frame = tk.Frame(header_frame, bg=HEADER_BG_COLOR)
header_text_frame.pack(side="left", padx=TEXT_PADX)

tk.Label(header_text_frame, text="Microsoft Post Migration Helper Wizard", bg=HEADER_BG_COLOR, fg="white",
         font=(FONT_FAMILY, LARGE_FONT_SIZE, "bold")).pack(padx=TEXT_PADX, pady=TEXT_PADY)

# Buttons and Status
button_frame = tk.Frame(root, bg="#f0f0f0")
button_frame.pack(side="left", padx=FRAME_PADX, pady=FRAME_PADY)

buttons = [
    ("1. Add Work Account", add_work_account),
    ("2. Clear Outlook Profiles", clear_outlook_profiles),
    ("3. Reset Edge", clear_edge),
    ("4. Set Edge Default", set_edge_default),
    ("5. Setup OneDrive", setup_onedrive),
    ("View Logs", show_logs),
    ("Save Logs (prompt)", save_logs),
]
status_labels = []
for idx, (text, func) in enumerate(buttons):
    button = ttk.Button(button_frame, text=text, command=lambda f=func, l=idx: threaded_task(f, status_labels[l]), width=BUTTON_WIDTH)
    button.pack(pady=BUTTON_PADY, padx=BUTTON_PADX, side="top", fill='x')
    status_label = tk.Label(button_frame, text="", font=(FONT_FAMILY, FONT_SIZE), width=10, anchor="center", bg="#f0f0f0")
    status_label.pack(pady=STATUS_PADY, side="top", fill='x')
    status_labels.append(status_label)

def save_and_exit():
  """Validates all tasks and exits the application."""
  for label in status_labels:
    if label.cget("text") == "Error":
      messagebox.showerror("Error", "Please resolve all errors before saving.")
      return
    if label.cget("text") == "Running...":
      messagebox.showerror("Error", "Please wait for all tasks to complete before saving.")
      return
  exit_app(0) # Exit code 0 to report success

# Save Button
ttk.Button(button_frame, text="Save", command=save_and_exit, width=BUTTON_WIDTH).pack(pady=BUTTON_PADY, padx=BUTTON_PADX, side="top", fill='x')

# Exit Button
ttk.Button(button_frame, text="Exit", command=exit_app, width=BUTTON_WIDTH).pack(pady=BUTTON_PADY, padx=BUTTON_PADX, side="top", fill='x')

# Info Box
info_text = scrolledtext.ScrolledText(root, wrap="word", state="disabled", font=(FONT_FAMILY, FONT_SIZE), bg=INFO_BOX_BG_COLOR, relief="groove")
info_text.pack(side="right", fill="both", expand=True, padx=FRAME_PADX, pady=FRAME_PADY)

# Progress Bar
progress_bar = ttk.Progressbar(root, mode="indeterminate")
progress_bar.pack(fill="x", padx=FRAME_PADX, pady=TEXT_PADY)

# Logs
logs = []

root.protocol("WM_DELETE_WINDOW", exit_app)
root.mainloop()
