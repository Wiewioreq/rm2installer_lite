# rm2updater_ctk.py
import os
import sys
import ctypes
import zipfile
import subprocess
import customtkinter as ctk
from tkinter import messagebox, StringVar, filedialog   # messagebox/vars/file dialog
import win32serviceutil
import win32service
from datetime import datetime
import time
import re
import threading
import contextlib
import tempfile
import shutil
import xml.etree.ElementTree as ET

# ---------------- UI wygląd (jak w ociusinstall) ----------------
ctk.set_appearance_mode("dark")          # "light" | "dark" | "system"
ctk.set_default_color_theme("blue")      # np. "dark-blue", "green"


# ==================== Logger do pliku + UI ======================
UI_LOGGER = None  # ustawiane w main()

class UILogger:
    """
    Prosty, wątkowo-bezpieczny logger do CTkTextbox:
    - dopisuje linie z timestampem,
    - autoscroll,
    - ograniczenie liczby linii (default 1500).
    """
    def __init__(self, root: ctk.CTk, textbox: ctk.CTkTextbox, max_lines: int = 1500):
        self.root = root
        self.textbox = textbox
        self.max_lines = max_lines

    def write(self, message: str):
        # Harmonogram na wątku GUI (tk jest single-threaded)
        ts = datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")
        def _append():
            try:
                self.textbox.insert("end", ts + message + "\n")
                self.textbox.see("end")
                # przytnij linie jeśli potrzeba
                line_count = int(float(self.textbox.index('end-1c').split('.')[0]))
                if line_count > self.max_lines:
                    # usuń pierwsze 50 linii, żeby nie męczyć I/O
                    self.textbox.delete("1.0", "50.0")
            except Exception:
                pass
        self.root.after(0, _append)

def set_ui_logger(root, textbox):
    global UI_LOGGER
    UI_LOGGER = UILogger(root, textbox)


# ============================== Walidacja ==============================
class InputValidator:
    @staticmethod
    def validate_database_name(name):
        if not name or len(name) > 128:
            return False, "Database name is required and must be less than 128 characters"
        if not re.match(r'^[a-zA-Z0-9_]+$', name):
            return False, "Database name can only contain letters, numbers, and underscores"
        return True, ""

    @staticmethod
    def validate_server_name(name):
        if not name:
            return False, "Server name is required"
        if not re.match(r'^[a-zA-Z0-9\\_.\\-]+$', name):
            return False, "Invalid server name format"
        return True, ""

    @staticmethod
    def validate_credentials(username, password, use_windows_auth):
        if not use_windows_auth:
            if not username:
                return False, "Username is required for SQL Server authentication"
            if not password:
                return False, "Password is required for SQL Server authentication"
        return True, ""


# ============================ Okno postępu =============================
class ProgressWindow:
    def __init__(self, parent):
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Update Progress")
        self.window.geometry("500x210")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.geometry("+%d+%d" % (parent.winfo_rootx() + 60, parent.winfo_rooty() + 60))
        self.window.resizable(False, False)

        self.frame = ctk.CTkFrame(self.window, corner_radius=10)
        self.frame.pack(fill="both", expand=True, padx=16, pady=16)

        self.title_label = ctk.CTkLabel(self.frame, text="RM2 Update in progress…", font=("Segoe UI", 16, "bold"))
        self.title_label.pack(pady=(0, 10))

        self.progress = ctk.CTkProgressBar(self.frame)
        self.progress.pack(fill="x", padx=4, pady=(6, 10))
        self.progress.start()  # indeterminate

        self.status_label = ctk.CTkLabel(self.frame, text="Starting update process…", wraplength=440, justify="left")
        self.status_label.pack(pady=(4, 10))

        self.cancel_button = ctk.CTkButton(self.frame, text="Cancel", command=self.cancel)
        self.cancel_button.pack(pady=(2, 0))

        self.cancelled = False

    def update_status(self, message):
        if not self.cancelled:
            self.status_label.configure(text=message)
            # wrzuć też do live loga (jeśli jest)
            try:
                if UI_LOGGER:
                    UI_LOGGER.write(message)
            except Exception:
                pass
            self.window.update()

    def cancel(self):
        self.cancelled = True
        self.close()

    def close(self):
        try:
            self.progress.stop()
            self.window.destroy()
        except Exception:
            pass


# ====================== Engineer Utilities Window ======================
ENGINEER_PASSWORD = "ThisIsTheWay!"

BRANCH_CONFIG_TEMPLATE = """\
<configuration>
  <appSettings>
    <add key="Main.SQLConnectionString" value="Data Source={{SQL_INSTANCE}};Initial Catalog={{DB_NAME}};{{AUTH_STRING}}" />
    <add key="Main.HeadOfficeIP" value="{{HO_URL}}" />
    <add key="Main.HeadOfficeUserName" value="" />
    <add key="Main.HeadOfficePassword" value="" />
</appSettings>
  <startup useLegacyV2RuntimeActivationPolicy="true">
<supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0"/>
</startup>
</configuration>
"""

HEADOFFICE_CONFIG_TEMPLATE = """\
<configuration>
  <appSettings>
    <add key="Main.SQLConnectionString" value="Data Source={{SQL_INSTANCE}};Initial Catalog={{DB_NAME}};{{AUTH_STRING}}" />
    <add key="Main.HeadOfficeIP" value="{{HO_URL}}" />
    <add key="Main.HeadOfficeUserName" value="" />
    <add key="Main.HeadOfficePassword" value="" />
</appSettings>
  <startup useLegacyV2RuntimeActivationPolicy="true">
<supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.0"/>
</startup>
</configuration>
"""


class EngineerUtilitiesWindow:
    """
    Modal window with advanced config management tools.
    Password-protected — open only after successful authentication.
    """

    def __init__(self, parent):
        self.parent = parent
        self.window = ctk.CTkToplevel(parent)
        self.window.title("🔧 Engineer Utilities")
        self.window.geometry("900x700")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.geometry("+%d+%d" % (parent.winfo_rootx() + 30, parent.winfo_rooty() + 30))
        self.window.resizable(True, True)

        # ---- scrollable main container ----
        self.main_frame = ctk.CTkScrollableFrame(self.window)
        self.main_frame.pack(fill="both", expand=True, padx=16, pady=16)

        self._build_sql_section()
        self._build_branch_ho_section()
        self._build_preview_section()
        self._build_deploy_section()

    # ------------------------------------------------------------------
    # Section 1 — SQL Config
    # ------------------------------------------------------------------
    def _build_sql_section(self):
        sec = ctk.CTkFrame(self.main_frame, corner_radius=8)
        sec.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(sec, text="Section 1 — SQL Config", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        form = ctk.CTkFrame(sec, fg_color="transparent")
        form.pack(fill="x", padx=12, pady=(0, 12))
        form.grid_columnconfigure(1, weight=1)

        # SQL Instance
        ctk.CTkLabel(form, text="SQL Instance:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        self.ent_sql_instance = ctk.CTkEntry(form, placeholder_text="e.g. .\\SQLEXPRESS01")
        self.ent_sql_instance.grid(row=0, column=1, sticky="we", pady=(0, 8))

        # Database Name
        ctk.CTkLabel(form, text="Database Name:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        self.ent_db_name = ctk.CTkEntry(form, placeholder_text="e.g. lo01ho")
        self.ent_db_name.grid(row=1, column=1, sticky="we", pady=(0, 8))

        # Windows Auth checkbox
        self.win_auth_var = ctk.BooleanVar(value=True)
        self.chk_win_auth = ctk.CTkCheckBox(
            form, text="Windows Auth (Integrated Security)",
            variable=self.win_auth_var, command=self._on_win_auth_toggle
        )
        self.chk_win_auth.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 8))

        # Username
        ctk.CTkLabel(form, text="Username:").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        self.ent_username = ctk.CTkEntry(form, placeholder_text="SQL username")
        self.ent_username.grid(row=3, column=1, sticky="we", pady=(0, 8))

        # Password
        ctk.CTkLabel(form, text="Password:").grid(row=4, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        self.ent_password = ctk.CTkEntry(form, placeholder_text="SQL password", show="*")
        self.ent_password.grid(row=4, column=1, sticky="we", pady=(0, 8))

        # Start with Windows Auth enabled → disable Username/Password
        self._on_win_auth_toggle()

    def _on_win_auth_toggle(self):
        state = "disabled" if self.win_auth_var.get() else "normal"
        self.ent_username.configure(state=state)
        self.ent_password.configure(state=state)

    # ------------------------------------------------------------------
    # Section 2 — Branch / Head Office Mode
    # ------------------------------------------------------------------
    def _build_branch_ho_section(self):
        sec = ctk.CTkFrame(self.main_frame, corner_radius=8)
        sec.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(sec, text="Section 2 — Branch / Head Office Mode", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        radio_frame = ctk.CTkFrame(sec, fg_color="transparent")
        radio_frame.pack(fill="x", padx=12, pady=(0, 4))

        self.mode_var = StringVar(value="branch")

        self.rb_branch = ctk.CTkRadioButton(
            radio_frame, text="Branch", variable=self.mode_var, value="branch",
            command=self._on_mode_change
        )
        self.rb_branch.pack(side="left", padx=(0, 20))

        self.rb_ho = ctk.CTkRadioButton(
            radio_frame, text="Head Office", variable=self.mode_var, value="headoffice",
            command=self._on_mode_change
        )
        self.rb_ho.pack(side="left")

        # HO URL field
        ho_form = ctk.CTkFrame(sec, fg_color="transparent")
        ho_form.pack(fill="x", padx=12, pady=(4, 12))
        ho_form.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(ho_form, text="HeadOffice URL:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        self.ent_ho_url = ctk.CTkEntry(ho_form, placeholder_text="e.g. http://91.246.8.162:8085/Service1.asmx")
        self.ent_ho_url.grid(row=0, column=1, sticky="we", pady=(0, 8))

    def _on_mode_change(self):
        # Both modes use the same URL field; keep it always enabled.
        # Validation (required vs optional) is handled in _generate_config().
        pass

    # ------------------------------------------------------------------
    # Section 3 — Preview
    # ------------------------------------------------------------------
    def _build_preview_section(self):
        sec = ctk.CTkFrame(self.main_frame, corner_radius=8)
        sec.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(sec, text="Section 3 — Preview", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        btn_preview = ctk.CTkButton(
            sec, text="🔍  Preview Final Config",
            fg_color="#607D8B", hover_color="#455A64",
            command=self._do_preview
        )
        btn_preview.pack(anchor="w", padx=12, pady=(0, 8))

        self.preview_textbox = ctk.CTkTextbox(sec, height=180, state="disabled")
        self.preview_textbox.pack(fill="x", padx=12, pady=(0, 12))

    def _do_preview(self):
        config_str = self._generate_config()
        if config_str is None:
            return
        self.preview_textbox.configure(state="normal")
        self.preview_textbox.delete("1.0", "end")
        self.preview_textbox.insert("end", config_str)
        self.preview_textbox.configure(state="disabled")

    # ------------------------------------------------------------------
    # Section 4 — Deploy
    # ------------------------------------------------------------------
    def _build_deploy_section(self):
        sec = ctk.CTkFrame(self.main_frame, corner_radius=8)
        sec.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(sec, text="Section 4 — Deploy", font=("Segoe UI", 13, "bold")).pack(
            anchor="w", padx=12, pady=(10, 6)
        )

        btn_deploy = ctk.CTkButton(
            sec, text="💾  Write Config to RM2 Folder",
            fg_color="#607D8B", hover_color="#455A64",
            command=self._do_deploy
        )
        btn_deploy.pack(anchor="w", padx=12, pady=(0, 12))

    def _do_deploy(self):
        config_str = self._generate_config()
        if config_str is None:
            return

        target_folder = filedialog.askdirectory(title="Select RM2 Target Folder")
        if not target_folder:
            messagebox.showerror("Error", "No folder selected.", parent=self.window)
            return

        config_path = os.path.join(target_folder, "DryStockView.exe.config")

        # Backup existing config
        if os.path.exists(config_path):
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = config_path + f".bak_{ts}"
            try:
                shutil.copy2(config_path, backup_path)
                log_action(f"[Engineer] Backup created: {backup_path}")
            except Exception as e:
                log_action(f"[Engineer] Backup failed: {e}")
                messagebox.showerror("Error", f"Failed to create backup:\n{e}", parent=self.window)
                return

        # Write new config
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                f.write(config_str)
            log_action(f"[Engineer] Config written to: {config_path}")
            messagebox.showinfo("Success", f"Config written successfully to:\n{config_path}", parent=self.window)
        except Exception as e:
            log_action(f"[Engineer] Failed to write config: {e}")
            messagebox.showerror("Error", f"Failed to write config:\n{e}", parent=self.window)

    # ------------------------------------------------------------------
    # Config generation (template-based, no XML parsing)
    # ------------------------------------------------------------------
    def _generate_config(self):
        """
        Reads the appropriate template and replaces placeholders.
        Returns the final config string, or None on validation error.
        """
        sql_instance = self.ent_sql_instance.get().strip()
        db_name = self.ent_db_name.get().strip()

        # Validate SQL Instance
        ok, err = InputValidator.validate_server_name(sql_instance)
        if not ok:
            messagebox.showerror("Validation Error", f"SQL Instance: {err}", parent=self.window)
            return None

        # Validate Database Name
        ok, err = InputValidator.validate_database_name(db_name)
        if not ok:
            messagebox.showerror("Validation Error", f"Database Name: {err}", parent=self.window)
            return None

        use_windows_auth = self.win_auth_var.get()

        if use_windows_auth:
            auth_string = "Integrated Security=True;"
        else:
            username = self.ent_username.get().strip()
            password = self.ent_password.get().strip()
            ok, err = InputValidator.validate_credentials(username, password, False)
            if not ok:
                messagebox.showerror("Validation Error", err, parent=self.window)
                return None
            auth_string = f"User ID={username};Password={password};"

        mode = self.mode_var.get()  # "branch" or "headoffice"
        ho_url = self.ent_ho_url.get().strip()

        if mode == "headoffice":
            if not ho_url:
                messagebox.showerror("Validation Error", "HeadOffice URL is required.", parent=self.window)
                return None
            config_str = HEADOFFICE_CONFIG_TEMPLATE
        else:
            config_str = BRANCH_CONFIG_TEMPLATE

        config_str = config_str.replace("{{SQL_INSTANCE}}", sql_instance)
        config_str = config_str.replace("{{DB_NAME}}", db_name)
        config_str = config_str.replace("{{AUTH_STRING}}", auth_string)
        config_str = config_str.replace("{{HO_URL}}", ho_url)

        return config_str


_ENGINEER_ATTEMPTS = 0
_ENGINEER_MAX_ATTEMPTS = 5


def open_engineer_utilities(parent):
    """
    Shows a password dialog. If the password matches, opens EngineerUtilitiesWindow.
    Locks out after 5 consecutive failed attempts.
    """
    global _ENGINEER_ATTEMPTS

    if _ENGINEER_ATTEMPTS >= _ENGINEER_MAX_ATTEMPTS:
        messagebox.showerror(
            "Access Locked",
            "Too many failed attempts. Restart the application to try again.",
            parent=parent,
        )
        log_action("[Engineer] Access locked — too many failed attempts.")
        return

    pwd_win = ctk.CTkToplevel(parent)
    pwd_win.title("Engineer Access")
    pwd_win.geometry("360x160")
    pwd_win.transient(parent)
    pwd_win.grab_set()
    pwd_win.geometry("+%d+%d" % (parent.winfo_rootx() + 120, parent.winfo_rooty() + 120))
    pwd_win.resizable(False, False)

    frame = ctk.CTkFrame(pwd_win, corner_radius=10)
    frame.pack(fill="both", expand=True, padx=16, pady=16)

    ctk.CTkLabel(frame, text="Enter Engineer Password:", font=("Segoe UI", 13)).pack(pady=(0, 8))

    ent_pwd = ctk.CTkEntry(frame, show="*", width=260, placeholder_text="Password")
    ent_pwd.pack(pady=(0, 12))
    ent_pwd.focus()

    def _check():
        global _ENGINEER_ATTEMPTS
        entered = ent_pwd.get()
        if entered == ENGINEER_PASSWORD:
            _ENGINEER_ATTEMPTS = 0
            log_action("[Engineer] Engineer Utilities accessed successfully.")
            pwd_win.destroy()
            EngineerUtilitiesWindow(parent)
        else:
            _ENGINEER_ATTEMPTS += 1
            remaining = _ENGINEER_MAX_ATTEMPTS - _ENGINEER_ATTEMPTS
            log_action(
                f"[Engineer] Failed access attempt ({_ENGINEER_ATTEMPTS}/{_ENGINEER_MAX_ATTEMPTS})."
            )
            if remaining > 0:
                messagebox.showerror(
                    "Access Denied",
                    f"Incorrect password. {remaining} attempt(s) remaining.",
                    parent=pwd_win,
                )
                ent_pwd.delete(0, "end")
            else:
                messagebox.showerror(
                    "Access Locked",
                    "Too many failed attempts. Restart the application to try again.",
                    parent=pwd_win,
                )
                log_action("[Engineer] Access locked — too many failed attempts.")
                pwd_win.destroy()

    btn_ok = ctk.CTkButton(frame, text="OK", command=_check, width=120)
    btn_ok.pack()

    # Allow pressing Enter to confirm
    ent_pwd.bind("<Return>", lambda e: _check())


# ============================== Narzędzia ==============================
def get_base_dir():
    """
    Zwraca folder, w którym znajduje się EXE (w trybie frozen)
    albo folder skryptu (w trybie .py).
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_resource_path(relative_path):
    """
    Teraz zasoby (np. rm2.zip) są szukane obok EXE/skryptu,
    a nie w _MEIPASS (czyli nie są bundlowane).
    """
    return os.path.join(get_base_dir(), relative_path)

def get_log_path():
    return os.path.join(get_base_dir(), "update_log.txt")

def log_action(message):
    """Zapis do pliku + do live loga (jeśli ustawiony)."""
    # do pliku
    try:
        with open(get_log_path(), "a", encoding="utf-8") as log_file:
            log_file.write(
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')} UTC \n"
                f"{os.getenv('USERNAME', 'unknown')} \n"
                f"{message}\n"
            )
    except Exception as e:
        print(f"Logging error: {e}")
    # do UI
    try:
        if UI_LOGGER:
            UI_LOGGER.write(message)
    except Exception:
        pass

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def restart_as_admin():
    try:
        if getattr(sys, 'frozen', False):
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, "", None, 1)
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()
    except Exception as e:
        log_action(f"Failed to restart as admin: {e}")
        messagebox.showerror("Error", "Failed to restart with admin privileges")
        sys.exit(1)

def wait_for_service_status(service_name, desired_status, timeout=30):
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            status = win32serviceutil.QueryServiceStatus(service_name)[1]
            if status == desired_status:
                return True
            time.sleep(1)
        except Exception:
            return False
    return False

def service_exists(service_name):
    try:
        win32serviceutil.QueryServiceStatus(service_name)
        return True
    except Exception:
        return False

@contextlib.contextmanager
def service_manager(services, progress_callback=None):
    stopped_services = []
    timeout = 30
    try:
        for service in services:
            if progress_callback:
                progress_callback(f"Stopping service: {service}")
            try:
                if not service_exists(service):
                    log_action(f"Service {service} not found, skipping")
                    continue
                status = win32serviceutil.QueryServiceStatus(service)[1]
                if status == win32service.SERVICE_STOPPED:
                    log_action(f"Service {service} already stopped")
                    continue
                win32serviceutil.StopService(service)
                if wait_for_service_status(service, win32service.SERVICE_STOPPED, timeout):
                    stopped_services.append(service)
                    log_action(f"Successfully stopped {service}")
                else:
                    log_action(f"Timeout waiting for {service} to stop")
                    raise Exception(f"Timeout waiting for {service} to stop")
            except Exception as e:
                log_action(f"Error stopping {service}: {e}")
                raise Exception(f"Failed to stop {service}: {e}")
        if progress_callback:
            progress_callback("All services stopped successfully")
        yield stopped_services
    finally:
        for service in stopped_services:
            if progress_callback:
                progress_callback(f"Starting service: {service}")
            try:
                win32serviceutil.StartService(service)
                if wait_for_service_status(service, win32service.SERVICE_RUNNING, timeout):
                    log_action(f"Successfully restarted {service}")
                else:
                    log_action(f"Timeout waiting for {service} to start")
            except Exception as e:
                log_action(f"Error restarting {service}: {e}")

def verify_directory_writable(directory):
    try:
        test_file = os.path.join(directory, "write_test_temp")
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        return True
    except Exception as e:
        log_action(f"Directory {directory} is not writable: {e}")
        return False

def extract_zip_to_temp(progress_callback=None):
    try:
        if progress_callback:
            progress_callback("Locating update files...")

        zip_path = get_resource_path("rm2.zip")
        if not os.path.exists(zip_path):
            log_action(f"rm2.zip not found next to executable/script: {zip_path}")
            raise Exception(f"rm2.zip not found next to executable/script: {zip_path}")

        temp_dir = tempfile.mkdtemp(prefix="rm2_update_")
        log_action(f"Created temp directory: {temp_dir}")

        if progress_callback:
            progress_callback("Extracting update files to temporary location...")

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        log_action(f"Successfully extracted ZIP ({zip_path}) to temp {temp_dir}")

        if progress_callback:
            progress_callback("Files extracted to temporary location")
        return temp_dir
    except Exception as e:
        error_msg = f"Failed to extract files: {str(e)}"
        log_action(error_msg)
        raise Exception(error_msg)

def copy_from_temp_to_target(temp_dir, target_folder, progress_callback=None):
    try:
        os.makedirs(target_folder, exist_ok=True)
        skipped_files = []
        rm2_subdir = os.path.join(temp_dir, "rm2")
        source_root = rm2_subdir if os.path.isdir(rm2_subdir) else temp_dir
        for root, dirs, files in os.walk(source_root):
            rel_path = os.path.relpath(root, source_root)
            dest_dir = os.path.join(target_folder, rel_path) if rel_path != "." else target_folder
            os.makedirs(dest_dir, exist_ok=True)
            for file in files:
                src_file = os.path.join(root, file)
                dest_file = os.path.join(dest_dir, file)
                try:
                    shutil.copy2(src_file, dest_file)
                    log_action(f"Copied {src_file} to {dest_file}")
                except PermissionError as e:
                    log_action(f"Permission error copying {src_file}: {e}")
                    skipped_files.append(dest_file)
                    if progress_callback:
                        progress_callback(f"Skipping locked file: {file}")
                except Exception as e:
                    log_action(f"Error copying {src_file}: {e}")
                    raise Exception(f"Failed to copy {file}: {e}")
        if skipped_files:
            log_action(f"Skipped files due to locks: {skipped_files}")
        if progress_callback:
            progress_callback("Update files copied to target folder")
        return True
    except Exception as e:
        log_action(f"Copy error: {e}")
        raise

def extract_and_copy(destination_folder, progress_callback=None):
    temp_dir = extract_zip_to_temp(progress_callback)
    try:
        copy_from_temp_to_target(temp_dir, destination_folder, progress_callback)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        log_action(f"Temp directory {temp_dir} removed")

def read_sql_config(config_folder):
    config_path = os.path.join(config_folder, "DryStockView.exe.config")
    if not os.path.exists(config_path):
        log_action(f"Config file not found: {config_path}")
        return None, None
    try:
        tree = ET.parse(config_path)
        root = tree.getroot()
        conn_str = None
        for add_elem in root.findall(".//appSettings/add"):
            key = add_elem.get('key')
            if key and key.lower() == "main.sqlconnectionstring":
                conn_str = add_elem.get('value')
                break
        if not conn_str:
            log_action("Main.SQLConnectionString not found in config")
            return None, None
        instance, db_name = None, None
        for part in conn_str.split(';'):
            if part.strip().lower().startswith("data source="):
                instance = part.split('=', 1)[1].strip()
            elif part.strip().lower().startswith("initial catalog="):
                db_name = part.split('=', 1)[1].strip()
        log_action(f"Read config file: instance={instance}, db_name={db_name}")
        return instance, db_name
    except Exception as e:
        log_action(f"Error reading config file: {e}")
        return None, None

def validate_sql_connection(server_name, db_name, use_windows_auth, username, password, progress_callback=None):
    if progress_callback:
        progress_callback("Validating database connection...")
    valid, msg = InputValidator.validate_database_name(db_name)
    if not valid:
        raise Exception(msg)
    valid, msg = InputValidator.validate_server_name(server_name)
    if not valid:
        raise Exception(msg)
    valid, msg = InputValidator.validate_credentials(username, password, use_windows_auth)
    if not valid:
        raise Exception(msg)
    if use_windows_auth:
        auth_params = ["-E"]
    else:
        auth_params = ["-U", username, "-P", password]
    try:
        subprocess.run(
            ["sqlcmd", "-S", server_name, "-d", db_name, "-Q", "SELECT 1", "-t", "10"] + auth_params,
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, timeout=30
        )
        log_action("SQL Server connection validated successfully")
        if progress_callback:
            progress_callback("Database connection validated")
        return True
    except subprocess.TimeoutExpired:
        log_action("SQL Server connection timeout")
        raise Exception("Connection to SQL Server timed out")
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.strip() if e.stderr else "Unknown SQL error"
        log_action(f"SQL Server connection failed: {error_msg}")
        raise Exception(f"Failed to connect to SQL Server: {error_msg}")


# ------------------ AUTO: wykrywanie numerowanych skryptów ------------------
def discover_update_scripts(scripts_folder: str):
    """
    Zbiera pliki SQL w formacie: update<NUMBER>.sql (np. update1.sql, update_2.sql, update-10.sql),
    zwraca listę nazw posortowaną rosnąco po numerze.
    """
    import re
    candidates = []
    for name in os.listdir(scripts_folder):
        if not name.lower().endswith(".sql"):
            continue
        # Obsługujemy: update1.sql | update_1.sql | update-1.sql (case-insensitive)
        m = re.match(r'^update\s*[-_]?(\d+)\.sql$', name, re.IGNORECASE)
        if m:
            n = int(m.group(1))
            candidates.append((n, name))
    if not candidates:
        raise Exception(
            "No numbered update scripts found in 'Scripts' "
            "(expected files like update1.sql, update2.sql, ...)"
        )
    candidates.sort(key=lambda x: x[0])
    return [name for _, name in candidates]


def run_sql_scripts(destination_folder, server_name, db_name, use_windows_auth, username, password, progress_callback=None):
    # Szukamy folderu Scripts jak dotychczas
    possible_paths = [
        os.path.join(destination_folder, "Scripts"),
        os.path.join(destination_folder, "scripts"),
        os.path.join(destination_folder, "rm2", "scripts"),
        os.path.join(destination_folder, "RM2", "Scripts")
    ]
    scripts_folder = None
    for path in possible_paths:
        if os.path.exists(path):
            scripts_folder = path
            log_action(f"Found scripts folder: {scripts_folder}")
            break
    if not scripts_folder:
        log_action(f"Scripts folder not found in any of: {possible_paths}")
        raise Exception("Scripts folder not found in extracted files")

    # --- NOWE: auto-odkrywanie updateN.sql ---
    try:
        script_files = discover_update_scripts(scripts_folder)
    except Exception as e:
        # Fallback: jeśli nie ma numerowanych, spróbuj klasyczne 2 pliki
        log_action(f"Auto-discovery failed ({e}). Falling back to update1.sql, update2.sql")
        fallback = []
        for nm in ["update1.sql", "update2.sql"]:
            p = os.path.join(scripts_folder, nm)
            if os.path.exists(p):
                fallback.append(nm)
        if not fallback:
            raise
        script_files = fallback

    log_action(f"Final SQL execution order: {script_files}")

    # Połączenie
    sql_timeout = 300
    if use_windows_auth:
        conn_params = f'-S "{server_name}" -d "{db_name}" -E'
    else:
        conn_params = f'-S "{server_name}" -d "{db_name}" -U "{username}" -P "{password}"'

    # Wykonanie
    try:
        for i, script_name in enumerate(script_files, 1):
            if progress_callback:
                progress_callback(f"Executing script {i}/{len(script_files)}: {script_name}")
            script_path = os.path.normpath(os.path.join(scripts_folder, script_name))
            if not os.path.exists(script_path):
                log_action(f"Script file not found: {script_path}")
                raise Exception(f"Script file not found: {script_name}")

            start_time = datetime.utcnow()
            log_action(f"[SQL EXEC START] {script_name}")

            cmd = f'sqlcmd {conn_params} -i "{script_path}" -t 30'
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                shell=True,
                timeout=sql_timeout
            )

            duration = datetime.utcnow() - start_time
            log_action(f"Return code: {result.returncode}")
            if result.stdout.strip():
                log_action(f"STDOUT:\n{result.stdout.strip()}")
            if result.stderr.strip():
                log_action(f"STDERR:\n{result.stderr.strip()}")

            errors = []
            for line in result.stdout.splitlines():
                line_lower = line.lower()
                if "error:" in line_lower or "msg" in line_lower:
                    # Ignorowane, spodziewane komunikaty
                    if any(ign in line_lower for ign in [
                        "already contains", "already exists", "already has",
                        "duplicate", "cannot insert duplicate",
                        "invalid column name", "msg 207",
                        "could not find stored procedure",
                        "invalid object name", "does not exist"
                    ]):
                        log_action(f"Ignored expected error in {script_name}: {line.strip()}")
                        continue
                    errors.append(line.strip())

            if result.returncode != 0:
                log_action(f"⚠️ Script {script_name} completed with warnings")
            else:
                log_action(f"✅ Script {script_name} executed successfully")

            if errors:
                error_msg = f"Errors found in {script_name}:\n" + "\n".join(errors)
                log_action(error_msg)
                log_action("Continuing execution despite errors...")

            log_action(f"[SQL EXEC END] {script_name} \nDuration: {duration}")

        if progress_callback:
            progress_callback("All SQL scripts executed")
        log_action("All SQL Server scripts completed")
        return True

    except subprocess.TimeoutExpired:
        log_action("SQL script execution timed out")
        raise Exception("SQL script execution timed out")
    except Exception as e:
        log_action(f"SQL script execution failed: {e}")
        raise Exception(f"SQL script execution failed: {e}")


# ------------------------- Resolver ścieżki docelowej ------------------
def resolve_destination(selected_folder: str):
    """
    Jeśli wskażesz literę dysku (C:, C:\, D:, D:\) → kopiujemy do <DYSK>\RM2.
    W przeciwnym razie kopiujemy bezpośrednio do wskazanego folderu.
    """
    if selected_folder is None:
        return None, None

    p = selected_folder.strip(' "\'')
    p = os.path.normpath(p)

    abs_p = os.path.abspath(p)
    drive, tail = os.path.splitdrive(abs_p)

    if drive and (tail in ('', os.sep)):  # root dysku
        drive_root = drive + os.sep
        target_folder = os.path.join(drive_root, 'RM2')
        config_folder = target_folder
    else:
        target_folder = abs_p
        config_folder = abs_p

    return target_folder, config_folder


# =============================== Worker ================================
def worker(operation_type, root, fields, buttons, status_setter, server_var, db_name_var, username_var, password_var):
    """
    fields: dict z entry do blokowania
    buttons: dict z przyciskami do blokowania
    status_setter: funkcja (text, color_hex)
    """
    progress = None
    try:
        folder = filedialog.askdirectory(title="Select destination folder")
        if not folder:
            root.after(0, lambda: messagebox.showerror("Error", "No folder selected"))
            return
        target_folder, config_folder = resolve_destination(folder)
        log_action(f"Selected path: {folder} -> target: {target_folder}; config: {config_folder}")

        if operation_type == "update":
            instance, db_name = read_sql_config(config_folder)

            def set_gui_fields():
                if instance:
                    server_var.set(instance)
                if db_name:
                    db_name_var.set(db_name)
            root.after(0, set_gui_fields)

            db_name_value = db_name if db_name else db_name_var.get().strip()
            server_name_value = instance if instance else server_var.get().strip()
            username = username_var.get().strip()
            password = password_var.get().strip()

            if not db_name_value or not server_name_value:
                root.after(0, lambda: messagebox.showerror("Error", "Database name and server instance are required"))
                return

            use_windows_auth = not username or not password
        else:
            username = password = ""
            use_windows_auth = True
            server_name_value = db_name_value = ""

        progress = ProgressWindow(root)
        root.after(0, lambda: status_setter("Processing…", "#FFC107"))  # amber

        if operation_type == "update":
            services = ['ITSOvernightservice', 'RM2updateservice']
            with service_manager(services, progress.update_status):
                if progress.cancelled:
                    return
                extract_and_copy(target_folder, progress.update_status)
                if progress.cancelled:
                    return
                validate_sql_connection(server_name_value, db_name_value, use_windows_auth, username, password,
                                        progress.update_status)
                if progress.cancelled:
                    return
                run_sql_scripts(target_folder, server_name_value, db_name_value, use_windows_auth, username, password,
                                progress.update_status)
                if progress.cancelled:
                    return
        else:
            extract_and_copy(target_folder, progress.update_status)
            if progress.cancelled:
                return

        progress.update_status(f"{operation_type.capitalize()} completed successfully!")
        log_action(f"{operation_type.capitalize()} process completed successfully")
        root.after(0, lambda: status_setter("Completed successfully", "#4CAF50"))  # green
        root.after(0, lambda: messagebox.showinfo("Success", f"{operation_type.capitalize()} process completed successfully!"))

    except Exception as e:
        log_action(f"{operation_type.capitalize()} process failed: {e}")
        root.after(0, lambda: status_setter("Failed", "#F44336"))  # red
        root.after(0, lambda: messagebox.showerror("Error", f"{operation_type.capitalize()} process failed:\n{str(e)}"))
    finally:
        if progress:
            root.after(0, progress.close)
        # odblokuj UI
        def _unlock():
            for w in fields.values():
                try: w.configure(state="normal")
                except Exception: pass
            for b in buttons.values():
                try: b.configure(state="normal")
                except Exception: pass
        root.after(0, _unlock)


# =============================== GUI (CTk) =============================
def main():
    if not is_admin():
        restart_as_admin()

    root = ctk.CTk()
    root.title("RM2 Update Manager")

    window_width, window_height = 760, 650
    x = (root.winfo_screenwidth() - window_width) // 2
    y = (root.winfo_screenheight() - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.resizable(False, False)

    # wstępna konfiguracja (z c:\rm2)
    instance, db_name = read_sql_config(r"c:\rm2")
    db_name_var = StringVar(master=root, value=db_name if db_name else "")
    server_var  = StringVar(master=root, value=instance if instance else "")
    username_var = StringVar(master=root, value="")
    password_var = StringVar(master=root, value="")

    # --- główny frame ---
    main_frame = ctk.CTkFrame(root, corner_radius=12)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    title_label = ctk.CTkLabel(main_frame, text="RM2 Update Manager", font=("Segoe UI", 20, "bold"))
    title_label.pack(pady=(10, 12))

    # ---------- formularz ----------
    form = ctk.CTkFrame(main_frame)
    form.pack(fill="x", padx=10, pady=10)

    # piktogramy + etykiety
    lbl_db  = ctk.CTkLabel(form, text="🗄️  Database Name:", font=("Segoe UI", 11))
    lbl_srv = ctk.CTkLabel(form, text="🖥️  SQL Server Instance (np. localhost\\SQLEXPRESS):", font=("Segoe UI", 11))
    lbl_usr = ctk.CTkLabel(form, text="👤  SQL Server Username (leave blank for Windows Auth):", font=("Segoe UI", 11))
    lbl_pwd = ctk.CTkLabel(form, text="🔑  SQL Server Password:", font=("Segoe UI", 11))

    ent_db  = ctk.CTkEntry(form, textvariable=db_name_var, width=480)
    ent_srv = ctk.CTkEntry(form, textvariable=server_var,  width=480)
    ent_usr = ctk.CTkEntry(form, textvariable=username_var, width=480)
    ent_pwd = ctk.CTkEntry(form, textvariable=password_var, show="*", width=480)

    lbl_db.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    ent_db.grid(row=0, column=1, sticky="we", pady=(0, 8))

    lbl_srv.grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    ent_srv.grid(row=1, column=1, sticky="we", pady=(0, 8))

    lbl_usr.grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    ent_usr.grid(row=2, column=1, sticky="we", pady=(0, 8))

    lbl_pwd.grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
    ent_pwd.grid(row=3, column=1, sticky="we", pady=(0, 8))

    form.grid_columnconfigure(1, weight=1)

    # ---------- przyciski + status ----------
    buttons_frame = ctk.CTkFrame(main_frame)
    buttons_frame.pack(fill="x", padx=10, pady=(8, 6))

    status_label = ctk.CTkLabel(buttons_frame, text="Status: idle", text_color="#9E9E9E", font=("Segoe UI", 11))
    status_label.pack(side="bottom", pady=(10, 0))

    def set_status(txt: str, color_hex: str):
        status_label.configure(text=f"Status: {txt}", text_color=color_hex)        

    # live log (na dole)
    logbox = ctk.CTkTextbox(main_frame, width=700, height=240)
    logbox.pack(fill="both", expand=False, padx=10, pady=(6, 10))
    set_ui_logger(root, logbox)  # włącz log na żywo

    # referencje do komponentów dla blokady
    fields = {"db": ent_db, "srv": ent_srv, "usr": ent_usr, "pwd": ent_pwd}
    buttons = {}  # uzupełnimy po utworzeniu przycisków

    def lock_ui():
        for w in fields.values():
            try: w.configure(state="disabled")
            except Exception: pass
        for b in buttons.values():
            try: b.configure(state="disabled")
            except Exception: pass
        set_status("Processing…", "#FFC107")  # amber

    def unlock_ui():
        for w in fields.values():
            try: w.configure(state="normal")
            except Exception: pass
        for b in buttons.values():
            try: b.configure(state="normal")
            except Exception: pass
        set_status("Ready", "#9E9E9E")

    def execute_all_threaded():
        lock_ui()
        thread = threading.Thread(
            target=lambda: worker("update", root, fields, buttons, set_status,
                                  server_var, db_name_var, username_var, password_var)
        )
        thread.daemon = True
        thread.start()

    def install_to_slave():
        lock_ui()
        thread = threading.Thread(
            target=lambda: worker("slave", root, fields, buttons, set_status,
                                  server_var, db_name_var, username_var, password_var)
        )
        thread.daemon = True
        thread.start()

    def test_connection():
        db_name_value = db_name_var.get().strip()
        server_name_value = server_var.get().strip()
        username = username_var.get().strip()
        password = password_var.get().strip()
        if not db_name_value or not server_name_value:
            messagebox.showerror("Error", "Database name and server instance are required")
            set_status("Missing DB/Server", "#F44336")  # red
            return
        use_windows_auth = not username or not password
        try:
            if validate_sql_connection(server_name_value, db_name_value, use_windows_auth, username, password):
                messagebox.showinfo("Success", "Connection test successful!")
                set_status("Connection OK", "#4CAF50")  # green
        except Exception as e:
            messagebox.showerror("Error", f"Connection test failed: {e}")
            set_status("Connection FAILED", "#F44336")  # red

    # przyciski z piktogramami
    btn_run      = ctk.CTkButton(buttons_frame, text="▶️  Run Update Process", command=execute_all_threaded)
    btn_test     = ctk.CTkButton(buttons_frame, text="🔎  Test Connection",    command=test_connection)
    btn_slave    = ctk.CTkButton(buttons_frame, text="📦  Slave",              command=install_to_slave,
                                 fg_color="#FF9800", hover_color="#d17f00")
    btn_engineer = ctk.CTkButton(buttons_frame, text="🔧  Engineer Utilities",
                                 command=lambda: open_engineer_utilities(root),
                                 fg_color="#607D8B", hover_color="#455A64")

    btn_run.pack(side="left", expand=True, fill="x", padx=(0, 8))
    btn_test.pack(side="left", padx=8)
    btn_slave.pack(side="right", padx=(8, 0))
    btn_engineer.pack(side="right", padx=(8, 0))

    buttons.update({"run": btn_run, "test": btn_test, "slave": btn_slave, "engineer": btn_engineer})

    # info
    info_label = ctk.CTkLabel(
        main_frame,
        text="This application will stop services, extract files, run SQL scripts, and restart services.",
        font=("Segoe UI", 10),
        justify="center"
    )
    info_label.pack(pady=(0, 4))

    # --- PÓŁPRZEZROCZYSTA STOPKA ---
    footer_label = ctk.CTkLabel(
          main_frame,
          text="Made by Pawel",
          font=("Segoe UI", 14),
          text_color=("#888888", "#555555")  # jasny szary / ciemny szary
    )
    footer_label.pack(pady=(10, 5))

    # start
    set_status("Ready", "#9E9E9E")
    root.mainloop()


if __name__ == "__main__":
    main()