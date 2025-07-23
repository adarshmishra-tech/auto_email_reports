import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd
import schedule
import time
import json
import os
import threading
import logging
import re
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# Setup logging
logging.basicConfig(filename='email_reports.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class AutoEmailReports:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Email Reports - Premium Edition")
        self.style = ttkb.Style(theme='flatly')
        self.root.geometry("900x650")
        self.root.resizable(False, False)

        self.config_file = 'email_config.json'
        self.config = self.load_config()

        self.running = False
        self.scheduling_thread = None
        self.attachments = self.config.get('attachments', [])

        self.setup_gui()

    def load_config(self):
        config = {
            'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
            'smtp_port': int(os.getenv('SMTP_PORT', 587)),
            'email': os.getenv('EMAIL', ''),
            'password': os.getenv('PASSWORD', ''),
            'recipients': [],
            'report_file': '',
            'schedule_time': '09:00',
            'subject': 'Daily Report',
            'attachments': []
        }
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    json_config = json.load(f)
                # Update config with non-sensitive info from JSON
                for key in ['recipients', 'report_file', 'schedule_time', 'subject', 'attachments']:
                    if key in json_config:
                        config[key] = json_config[key]
            except Exception as e:
                logging.warning(f"Failed to load config.json: {e}")
        return config

    def save_config(self):
        # Save only non-sensitive info to config file
        save_data = {
            'recipients': self.config.get('recipients', []),
            'report_file': self.config.get('report_file', ''),
            'schedule_time': self.config.get('schedule_time', '09:00'),
            'subject': self.config.get('subject', 'Daily Report'),
            'attachments': self.attachments
        }
        with open(self.config_file, 'w') as f:
            json.dump(save_data, f, indent=4)

    def setup_gui(self):
        frame = ttk.Frame(self.root, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        # SMTP Server (from .env but editable)
        ttk.Label(frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.smtp_server = ttk.Entry(frame, width=40)
        self.smtp_server.insert(0, self.config['smtp_server'])
        self.smtp_server.grid(row=0, column=1, pady=4)

        # SMTP Port (from .env but editable)
        ttk.Label(frame, text="SMTP Port:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.smtp_port = ttk.Entry(frame, width=40)
        self.smtp_port.insert(0, str(self.config['smtp_port']))
        self.smtp_port.grid(row=1, column=1, pady=4)

        # Sender Email - read-only from .env
        ttk.Label(frame, text="Sender Email (from .env):").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.email_label = ttk.Label(frame, text=self.config['email'] or "<Set EMAIL in .env>", foreground="blue")
        self.email_label.grid(row=2, column=1, sticky=tk.W, pady=4)

        # Password - NOT displayed for security, but can have a reset button if needed
        ttk.Label(frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.password_label = ttk.Label(frame, text="(hidden for security)", foreground="red")
        self.password_label.grid(row=3, column=1, sticky=tk.W, pady=4)

        # Recipients
        ttk.Label(frame, text="Recipients (comma-separated):").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.recipients = ttk.Entry(frame, width=40)
        self.recipients.insert(0, ','.join(self.config['recipients']))
        self.recipients.grid(row=4, column=1, pady=4)

        # Report File
        ttk.Label(frame, text="Report File (CSV):").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.report_file = ttk.Entry(frame, width=40)
        self.report_file.insert(0, self.config['report_file'])
        self.report_file.grid(row=5, column=1, pady=4)
        ttk.Button(frame, text="Browse", command=self.browse_report).grid(row=5, column=2, padx=6)

        # Attachments
        ttk.Label(frame, text="Attachments (optional):").grid(row=6, column=0, sticky=tk.W, pady=4)
        self.attachments_label = ttk.Label(frame, text=", ".join(self.attachments), wraplength=400)
        self.attachments_label.grid(row=6, column=1, sticky=tk.W)
        ttk.Button(frame, text="Add Attachments", command=self.add_attachments).grid(row=6, column=2, padx=6)

        # Schedule Time
        ttk.Label(frame, text="Schedule Time (HH:MM, 24h):").grid(row=7, column=0, sticky=tk.W, pady=4)
        self.schedule_time = ttk.Entry(frame, width=40)
        self.schedule_time.insert(0, self.config['schedule_time'])
        self.schedule_time.grid(row=7, column=1, pady=4)

        # Email Subject
        ttk.Label(frame, text="Email Subject:").grid(row=8, column=0, sticky=tk.W, pady=4)
        self.subject = ttk.Entry(frame, width=40)
        self.subject.insert(0, self.config['subject'])
        self.subject.grid(row=8, column=1, pady=4)

        # Buttons
        ttk.Button(frame, text="Save Configuration", style="primary.TButton", command=self.save_config_gui).grid(row=9, column=0, pady=15)
        ttk.Button(frame, text="Preview Report", style="info.TButton", command=self.preview_report).grid(row=9, column=1)
        ttk.Button(frame, text="Send Test Email", style="info.TButton", command=self.test_email).grid(row=10, column=0, pady=10)
        self.start_button = ttk.Button(frame, text="Start Scheduler", style="success.TButton", command=self.start_scheduler)
        self.start_button.grid(row=10, column=1, pady=10)

        # Status Label
        self.status_label = ttk.Label(frame, text="Status: Stopped", font=('Segoe UI', 10, 'italic'))
        self.status_label.grid(row=11, column=0, columnspan=3, pady=15)

    def browse_report(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.report_file.delete(0, tk.END)
            self.report_file.insert(0, path)

    def add_attachments(self):
        files = filedialog.askopenfilenames(title="Select attachment files")
        if files:
            self.attachments.extend([f for f in files if f not in self.attachments])
            self.attachments_label.config(text=", ".join(self.attachments))

    def save_config_gui(self):
        if not self.validate_inputs():
            return
        # Use SMTP and port from entries (allow editing)
        self.config['smtp_server'] = self.smtp_server.get()
        self.config['smtp_port'] = int(self.smtp_port.get())
        # Email and password come from .env, no changes here
        self.config['recipients'] = [r.strip() for r in self.recipients.get().split(',') if r.strip()]
        self.config['report_file'] = self.report_file.get()
        self.config['schedule_time'] = self.schedule_time.get()
        self.config['subject'] = self.subject.get()
        self.attachments = list(set(self.attachments))  # deduplicate
        self.config['attachments'] = self.attachments
        self.save_config()
        messagebox.showinfo("Success", "Configuration saved successfully!")

    def validate_inputs(self):
        email_regex = r"[^@]+@[^@]+\.[^@]+"

        # Validate email from .env exists and is valid
        sender_email = self.config.get('email', '')
        if not sender_email or not re.match(email_regex, sender_email):
            messagebox.showerror("Validation Error", "Sender email in .env is invalid or missing.")
            return False

        recipients_list = [r.strip() for r in self.recipients.get().split(',')]
        if not recipients_list or any(not re.match(email_regex, r) for r in recipients_list):
            messagebox.showerror("Validation Error", "One or more recipient emails are invalid.")
            return False
        try:
            port = int(self.smtp_port.get())
            if port <= 0 or port > 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Validation Error", "SMTP port must be a valid number (1-65535).")
            return False
        if not os.path.isfile(self.report_file.get()):
            messagebox.showerror("Validation Error", "Report file path is invalid or does not exist.")
            return False
        schedule_time = self.schedule_time.get()
        if not re.match(r'^\d{2}:\d{2}$', schedule_time):
            messagebox.showerror("Validation Error", "Schedule time must be in HH:MM 24-hour format.")
            return False
        hour, minute = map(int, schedule_time.split(':'))
        if not (0 <= hour < 24 and 0 <= minute < 60):
            messagebox.showerror("Validation Error", "Schedule time is not a valid time.")
            return False
        if not self.subject.get().strip():
            messagebox.showerror("Validation Error", "Email subject cannot be empty.")
            return False
        return True

    def preview_report(self):
        path = self.report_file.get()
        if not os.path.isfile(path):
            messagebox.showerror("Error", "Report file not found.")
            return
        try:
            df = pd.read_csv(path)
            preview_text = df.head(30).to_string(index=False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV report: {e}")
            return
        # Show preview window
        preview_win = tk.Toplevel(self.root)
        preview_win.title("Report Preview")
        preview_win.geometry("700x400")
        text_widget = tk.Text(preview_win, wrap='none', font=("Consolas", 10))
        text_widget.pack(expand=True, fill=tk.BOTH)
        text_widget.insert(tk.END, preview_text)
        text_widget.config(state=tk.DISABLED)
        # Add scrollbars
        y_scroll = ttk.Scrollbar(preview_win, orient=tk.VERTICAL, command=text_widget.yview)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=y_scroll.set)
        x_scroll = ttk.Scrollbar(preview_win, orient=tk.HORIZONTAL, command=text_widget.xview)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        text_widget.config(xscrollcommand=x_scroll.set)

    def send_email(self):
        try:
            if not self.config['email'] or not self.config['password']:
                raise Exception("Sender email or password missing in environment variables.")

            # Compose email
            msg = MIMEMultipart()
            msg['From'] = self.config['email']
            msg['To'] = ', '.join(self.config['recipients'])
            msg['Subject'] = self.config['subject']

            # Attach report content as plain text
            df = pd.read_csv(self.config['report_file'])
            report_text = df.to_string(index=False)
            msg.attach(MIMEText(report_text, 'plain'))

            # Attach additional files if any
            for filepath in self.attachments:
                if os.path.isfile(filepath):
                    with open(filepath, 'rb') as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(filepath))
                    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filepath)}"'
                    msg.attach(part)

            # Connect & send email
            with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                server.starttls()
                server.login(self.config['email'], self.config['password'])
                server.send_message(msg)

            logging.info(f"Email sent successfully to {', '.join(self.config['recipients'])}")
            self.root.after(0, lambda: self.status_label.config(
                text=f"Status: Email sent successfully at {datetime.now().strftime('%H:%M:%S')}"))
        except Exception as e:
            logging.error(f"Failed to send email: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to send email: {str(e)}"))

    def test_email(self):
        if not self.validate_inputs():
            return
        self.save_config_gui()
        self.send_email()

    def run_scheduler(self):
        schedule.clear()
        schedule.every().day.at(self.config['schedule_time']).do(self.send_email)
        while self.running:
            schedule.run_pending()
            time.sleep(1)

    def start_scheduler(self):
        if self.running:
            messagebox.showinfo("Scheduler", "Scheduler is already running.")
            return
        if not self.validate_inputs():
            return
        self.save_config_gui()
        self.running = True
        self.scheduling_thread = threading.Thread(target=self.run_scheduler, daemon=True)
        self.scheduling_thread.start()
        self.start_button.config(text="Stop Scheduler", style="danger.TButton", command=self.stop_scheduler)
        self.status_label.config(text=f"Status: Scheduler running, next run at {self.config['schedule_time']}")

    def stop_scheduler(self):
        if not self.running:
            messagebox.showinfo("Scheduler", "Scheduler is not running.")
            return
        self.running = False
        schedule.clear()
        self.start_button.config(text="Start Scheduler", style="success.TButton", command=self.start_scheduler)
        self.status_label.config(text="Status: Stopped")

if __name__ == "__main__":
    root = ttkb.Window()
    app = AutoEmailReports(root)
    root.mainloop()

