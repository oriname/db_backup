import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import datetime
import os
import json
import threading
import schedule
import time
import winreg
import sys

class DatabaseBackupApp:
    def get_resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def __init__(self, root):
        self.root = root
        self.root.title("Clarkeprint dbBackup Tool")
        self.root.geometry("600x550")
        
        self.scheduler_thread = None
        self.is_scheduler_running = False
        self.backup_in_progress = False
        self.stop_backup_flag = False
        
        app_data = os.path.join(os.environ['APPDATA'], 'SQLBackupTool')
        os.makedirs(app_data, exist_ok=True)
        self.config_file = os.path.join(app_data, "backup_config.json")
        
        self.load_config()
        self.create_gui()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        if self.config.get("scheduler_active", False):
            self.start_scheduler()
            self.schedule_button.config(text="Stop Scheduler")
    def load_config(self):
        """Load saved database configurations"""
        self.config = {
            "server": "localhost",
            "database": "",
            "username": "",
            "password": "",
            "trusted_connection": "yes",
            "backup_path": os.path.expanduser("~/Desktop/backups"),
            "backup_time": "23:00",
            "auto_start": False,
            "scheduler_active": False
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    saved_config = json.load(f)
                    self.config.update(saved_config)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {str(e)}")

    def save_settings(self):
        """Save current settings"""
        try:
            if not self.server_entry.get().strip():
                raise ValueError("Server name is required")
            if not self.db_entry.get().strip():
                raise ValueError("Database name is required")
            if self.auth_type.get() == "sql":
                if not self.user_entry.get().strip():
                    raise ValueError("Username is required for SQL Server Authentication")
                if not self.pass_entry.get().strip():
                    raise ValueError("Password is required for SQL Server Authentication")
            
            backup_path = self.backup_path.get().strip()
            if not backup_path:
                raise ValueError("Backup path is required")

            self.config["server"] = self.server_entry.get()
            self.config["database"] = self.db_entry.get()
            self.config["trusted_connection"] = "yes" if self.auth_type.get() == "windows" else "no"
            if self.auth_type.get() == "sql":
                self.config["username"] = self.user_entry.get()
                self.config["password"] = self.pass_entry.get()
            self.config["backup_path"] = self.backup_path.get()
            self.config["backup_time"] = self.backup_time.get()
            self.config["auto_start"] = self.auto_start_var.get()
            self.config["scheduler_active"] = self.is_scheduler_running
            
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f)
            messagebox.showinfo("Success", "Settings saved successfully!")
            return True
            
        except ValueError as e:
            messagebox.showerror("Validation Error", str(e))
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
            return False

    def save_config(self):
        """Save current configurations"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")
    
    def create_gui(self):
        """Create the GUI elements"""
        # Connection Frame
        connection_frame = ttk.LabelFrame(self.root, text="SQL Server Connection Settings")
        connection_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        # Server Settings
        ttk.Label(connection_frame, text="Server:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
        self.server_entry = ttk.Entry(connection_frame, width=40)
        self.server_entry.insert(0, self.config["server"])
        self.server_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=2)
        
        ttk.Label(connection_frame, text="Database:").grid(row=1, column=0, padx=5, pady=2, sticky="e")
        self.db_entry = ttk.Entry(connection_frame, width=40)
        self.db_entry.insert(0, self.config["database"])
        self.db_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=2)
        
        # Authentication Frame
        auth_frame = ttk.LabelFrame(connection_frame, text="Authentication")
        auth_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        
        self.auth_type = tk.StringVar(value="windows" if self.config["trusted_connection"] == "yes" else "sql")
        ttk.Radiobutton(auth_frame, text="Windows Authentication", 
                       variable=self.auth_type, value="windows",
                       command=self.toggle_auth).grid(row=0, column=0, padx=5, pady=2)
        ttk.Radiobutton(auth_frame, text="SQL Server Authentication", 
                       variable=self.auth_type, value="sql",
                       command=self.toggle_auth).grid(row=0, column=1, padx=5, pady=2)
        
        # SQL Authentication fields
        self.sql_auth_frame = ttk.Frame(auth_frame)
        self.sql_auth_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
        
        ttk.Label(self.sql_auth_frame, text="Username:").grid(row=0, column=0, padx=5, pady=2)
        self.user_entry = ttk.Entry(self.sql_auth_frame, width=30, show="*")
        self.user_entry.insert(0, self.config["username"])
        self.user_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(self.sql_auth_frame, text="Password:").grid(row=1, column=0, padx=5, pady=2)
        self.pass_entry = ttk.Entry(self.sql_auth_frame, width=30, show="*")
        self.pass_entry.insert(0, self.config["password"])
        self.pass_entry.grid(row=1, column=1, padx=5, pady=2)
        
        # Backup Settings
        backup_frame = ttk.LabelFrame(self.root, text="Backup Settings")
        backup_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        ttk.Label(backup_frame, text="Backup Location:").grid(row=0, column=0, padx=5, pady=5)
        self.backup_path = ttk.Entry(backup_frame, width=40)
        self.backup_path.insert(0, self.config["backup_path"])
        self.backup_path.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(backup_frame, text="Browse", command=self.browse_backup_location).grid(row=0, column=2, padx=5, pady=5)
        
        # Schedule Settings
        schedule_frame = ttk.LabelFrame(self.root, text="Schedule Settings")
        schedule_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        ttk.Label(schedule_frame, text="Backup Time (24h):").grid(row=0, column=0, padx=5, pady=5)
        self.backup_time = ttk.Entry(schedule_frame, width=10)
        self.backup_time.insert(0, self.config.get("backup_time", "23:00"))
        self.backup_time.grid(row=0, column=1, padx=5, pady=5)
        
        self.auto_start_var = tk.BooleanVar(value=self.config.get("auto_start", False))
        ttk.Checkbutton(schedule_frame, text="Start with Windows", 
                       variable=self.auto_start_var,
                       command=self.toggle_auto_start).grid(row=1, column=0, columnspan=2, pady=5)
        
        self.scheduler_status_var = tk.StringVar(value="Scheduler: Stopped")
        ttk.Label(schedule_frame, textvariable=self.scheduler_status_var).grid(row=2, column=0, columnspan=2, pady=5)
        
        # Status and Progress
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status_var).grid(row=3, column=0, columnspan=2, padx=5, pady=(5,0))
        
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.progress.grid_remove()
        
        # Button Frame
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Test Connection", 
                  command=self.test_connection).grid(row=0, column=0, padx=5)
        
        self.start_backup_button = ttk.Button(button_frame, text="Start Backup", 
                                            command=self.start_backup)
        self.start_backup_button.grid(row=0, column=1, padx=5)
        
        self.stop_backup_button = ttk.Button(button_frame, text="Stop Backup", 
                                           command=self.stop_backup)
        self.stop_backup_button.grid(row=0, column=2, padx=5)
        self.stop_backup_button.grid_remove()
        
        ttk.Button(button_frame, text="Save Settings", 
                  command=self.save_settings).grid(row=0, column=3, padx=5)
        
        self.schedule_button = ttk.Button(button_frame, text="Start Scheduler", 
                                        command=self.toggle_scheduler)
        self.schedule_button.grid(row=0, column=4, padx=5)
        
        # Initial auth type setup
        self.toggle_auth()


    def toggle_auth(self):
        """Toggle between Windows and SQL Server authentication"""
        if self.auth_type.get() == "windows":
            self.sql_auth_frame.grid_remove()
        else:
            self.sql_auth_frame.grid()

    def toggle_auto_start(self):
        """Toggle auto-start with Windows"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                            r"Software\Microsoft\Windows\CurrentVersion\Run", 
                            0, winreg.KEY_SET_VALUE)
            
            if self.auto_start_var.get():
                if getattr(sys, 'frozen', False):
                    application_path = f'"{sys.executable}"'
                else:
                    application_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
                
                winreg.SetValueEx(key, "SQLBackupTool", 0, winreg.REG_SZ, application_path)
            else:
                try:
                    winreg.DeleteValue(key, "SQLBackupTool")
                except WindowsError:
                    pass
            
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Auto-start error: {str(e)}")
            messagebox.showerror("Error", f"Failed to set auto-start: {str(e)}")

    def browse_backup_location(self):
        """Browse for backup location"""
        directory = filedialog.askdirectory(title="Select Backup Location")
        if directory:
            self.backup_path.delete(0, tk.END)
            self.backup_path.insert(0, directory)

    def get_connection_string(self):
        """Generate connection string based on authentication type"""
        if self.auth_type.get() == "windows":
            return f"DRIVER={{SQL Server}};SERVER={self.server_entry.get()};DATABASE={self.db_entry.get()};Trusted_Connection=yes;"
        else:
            return f"DRIVER={{SQL Server}};SERVER={self.server_entry.get()};DATABASE={self.db_entry.get()};UID={self.user_entry.get()};PWD={self.pass_entry.get()}"

    def test_connection(self):
        """Test database connection"""
        try:
            conn = pyodbc.connect(self.get_connection_string(), autocommit=True, timeout=10)
            conn.close()
            messagebox.showinfo("Success", "Connection test successful!")
        except Exception as e:
            messagebox.showerror("Error", f"Connection failed: {str(e)}")

    def on_closing(self):
        """Handle window closing"""
        try:
            if self.is_scheduler_running:
                if messagebox.askokcancel("Quit", "Scheduler is running. Stop it and exit?"):
                    self.stop_scheduler()
                else:
                    return

            if hasattr(self, 'stop_backup_flag') and self.stop_backup_flag:
                if messagebox.askokcancel("Quit", "Backup is in progress. Cancel it and exit?"):
                    self.stop_backup()
                else:
                    return
                    
            self.save_settings()
            self.root.destroy()
        except Exception as e:
            print(f"Error during cleanup: {str(e)}")
            self.root.destroy()

    def manage_backup_files(self, current_backup):
        """Keep only the two most recent backups"""
        try:
            backup_dir = os.path.dirname(current_backup)
            database_name = self.db_entry.get()
            
            backup_files = [f for f in os.listdir(backup_dir) 
                        if f.startswith(database_name) and f.endswith('.bak')]
            
            backup_files.sort(key=lambda x: os.path.getctime(os.path.join(backup_dir, x)), 
                            reverse=True)
            
            for old_backup in backup_files[2:]:
                try:
                    old_backup_path = os.path.join(backup_dir, old_backup)
                    os.remove(old_backup_path)
                    print(f"Removed old backup: {old_backup}")
                except Exception as e:
                    print(f"Error removing old backup {old_backup}: {str(e)}")
        
        except Exception as e:
            print(f"Error managing backup files: {str(e)}")
    
    def toggle_scheduler(self):
        if self.is_scheduler_running:
            self.stop_scheduler()
            self.schedule_button.config(text="Start Scheduler")
            self.scheduler_status_var.set("Scheduler: Stopped")
        else:
            self.start_scheduler()
            self.schedule_button.config(text="Stop Scheduler")
            self.scheduler_status_var.set("Scheduler: Running")

    def start_scheduler(self):
        if not self.is_scheduler_running:
            try:
                backup_time = self.backup_time.get()
                datetime.datetime.strptime(backup_time, "%H:%M")
                
                self.is_scheduler_running = True
                schedule.every().day.at(backup_time).do(self.start_backup)
                
                self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
                self.scheduler_thread.start()
                
                self.scheduler_status_var.set(f"Scheduler: Running (Next: {backup_time})")
            except ValueError:
                messagebox.showerror("Error", "Invalid time format. Please use HH:MM format (24-hour)")

    def stop_scheduler(self):
        self.is_scheduler_running = False
        schedule.clear()

    def run_scheduler(self):
        while self.is_scheduler_running:
            schedule.run_pending()
            time.sleep(60)

    def start_backup(self):
        self.progress.grid()
        self.progress.start()
        
        self.start_backup_button.config(state='disabled')
        self.stop_backup_button.grid()
        self.stop_backup_button.config(state='normal')
        
        thread = threading.Thread(target=self.perform_backup)
        thread.daemon = False
        thread.start()

    def stop_backup(self):
        try:
            self.stop_backup_flag = True
            self.status_var.set("Cancelling backup...")
            self.stop_backup_button.config(state='disabled')
            
            if hasattr(self, 'cancel_conn'):
                try:
                    cursor = self.cancel_conn.cursor()
                    cursor.execute(f"""
                    SELECT spid 
                    FROM master..sysprocesses 
                    WHERE dbid = DB_ID('{self.db_entry.get()}') 
                    AND cmd = 'BACKUP DATABASE'
                    """)
                    
                    row = cursor.fetchone()
                    if row:
                        cursor.execute(f"KILL {row[0]}")
                    cursor.close()
                except Exception as e:
                    print(f"Error stopping backup: {str(e)}")
        finally:
            if hasattr(self, 'cancel_conn'):
                try:
                    self.cancel_conn.close()
                except:
                    pass

    def perform_backup(self):
        conn = None
        backup_file = None
        try:
            self.stop_backup_flag = False
            self.status_var.set("Connecting to database...")
            conn = pyodbc.connect(self.get_connection_string(), autocommit=True, timeout=10)
            
            self.cancel_conn = pyodbc.connect(self.get_connection_string(), autocommit=True, timeout=10)
            cursor = conn.cursor()
            
            if self.stop_backup_flag:
                raise Exception("Backup cancelled by user")
                
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(
                self.backup_path.get(),
                f"{self.db_entry.get()}_{timestamp}.bak"
            )
            
            os.makedirs(os.path.dirname(backup_file), exist_ok=True)
            
            self.status_var.set("Starting backup process...")
            
            if self.stop_backup_flag:
                raise Exception("Backup cancelled by user")
                    
            backup_query = f"""
            BACKUP DATABASE [{self.db_entry.get()}]
            TO DISK = ?
            WITH FORMAT, STATS = 10
            """
            cursor.execute(backup_query, (backup_file,))
            
            while cursor.nextset():
                if self.stop_backup_flag:
                    raise Exception("Backup cancelled by user")
                pass
            
            if not self.stop_backup_flag:
                self.status_var.set("Backup completed successfully!")
                self.manage_backup_files(backup_file)
                
                if not self.is_scheduler_running:
                    messagebox.showinfo("Success", f"Backup completed successfully!\nSaved to: {backup_file}")
            
        except Exception as e:
            self.status_var.set("Backup failed!" if not self.stop_backup_flag else "Backup cancelled")
            error_msg = f"Backup failed: {str(e)}" if not self.stop_backup_flag else "Backup cancelled by user"
            print(error_msg)
            
            if not self.is_scheduler_running and not self.stop_backup_flag:
                messagebox.showerror("Error", error_msg)
            
            if self.stop_backup_flag and backup_file and os.path.exists(backup_file):
                try:
                    os.remove(backup_file)
                except:
                    pass
        
        finally:
            self.progress.stop()
            self.progress.grid_remove()
            if conn:
                try:
                    conn.close()
                except:
                    pass
            if hasattr(self, 'cancel_conn'):
                try:
                    self.cancel_conn.close()
                except:
                    pass
            self.start_backup_button.config(state='normal')
            self.stop_backup_button.grid_remove()
            self.status_var.set("Ready")

if __name__ == "__main__":
    root = tk.Tk()
    app = DatabaseBackupApp(root)
    root.mainloop()