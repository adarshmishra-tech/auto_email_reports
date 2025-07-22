import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import schedule
import time
import json
import os
from datetime import datetime
import threading
import logging

# Setup logging
logging.basicConfig(filename='email_reports.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class AutoEmailReports:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Email Reports")
        self.style = ttkb.Style(theme='flatly')  # Modern theme
        self.root.geometry("800x600")
        
        # Load configurations
        self.config_file = 'email_config.json'
        self.config = self.load_config()
        
        # Initialize scheduling thread
        self.scheduling_thread = None
        self.running = False
        
        # Setup GUI
        self.setup_gui()
        
    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as f:
                return json.load(f)
        return {
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'email': '',
            'password': '',
            'recipients': [],
            'report_file': '',
            'schedule_time': '09:00',
            'subject': 'Daily Report'
        }
    
    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=4)
    
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Email Settings
        ttk.Label(main_frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.smtp_server = ttk.Entry(main_frame, width=40)
        self.smtp_server.insert(0, self.config['smtp_server'])
        self.smtp_server.grid(row=0, column=1, pady=2)
        
        ttk.Label(main_frame, text="SMTP Port:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.smtp_port = ttk.Entry(main_frame, width=40)
        self.smtp_port.insert(0, str(self.config['smtp_port']))
        self.smtp_port.grid(row=1, column=1, pady=2)
        
        ttk.Label(main_frame, text="Email:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.email = ttk.Entry(main_frame, width=40)
        self.email.insert(0, self.config['email'])
        self.email.grid(row=2, column=1, pady=2)
        
        ttk.Label(main_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.password = ttk.Entry(main_frame, width=40, show="*")
        self.password.insert(0, self.config['password'])
        self.password.grid(row=3, column=1, pady=2)
        
        # Recipients
        ttk.Label(main_frame, text="Recipients (comma-separated):").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.recipients = ttk.Entry(main_frame, width=40)
        self.recipients.insert(0, ','.join(self.config['recipients']))
        self.recipients.grid(row=4, column=1, pady=2)
        
        # Report File
        ttk.Label(main_frame, text="Report File:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.report_file = ttk.Entry(main_frame, width=40)
        self.report_file.insert(0, self.config['report_file'])
        self.report_file.grid(row=5, column=1, pady=2)
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(row=5, column=2, padx=5)
        
        # Schedule
        ttk.Label(main_frame, text="Schedule Time (HH:MM):").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.schedule_time = ttk.Entry(main_frame, width=40)
        self.schedule_time.insert(0, self.config['schedule_time'])
        self.schedule_time.grid(row=6, column=1, pady=2)
        
        # Subject
        ttk.Label(main_frame, text="Email Subject:").grid(row=7, column=0, sticky=tk.W, pady=2)
        self.subject = ttk.Entry(main_frame, width=40)
        self.subject.insert(0, self.config['subject'])
        self.subject.grid(row=7, column=1, pady=2)
        
        # Buttons
        ttk.Button(main_frame, text="Save Config", style="primary.TButton", command=self.save_config_gui).grid(row=8, column=0, columnspan=2, pady=10)
        ttk.Button(main_frame, text="Test Email", style="info.TButton", command=self.test_email).grid(row=9, column=0, columnspan=2, pady=5)
        self.start_button = ttk.Button(main_frame, text="Start Scheduler", style="success.TButton", command=self.start_scheduler)
        self.start_button.grid(row=10, column=0, columnspan=2, pady=5)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="Status: Stopped")
        self.status_label.grid(row=11, column=0, columnspan=2, pady=5)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            self.report_file.delete(0, tk.END)
            self.report_file.insert(0, file_path)
    
    def save_config_gui(self):
        self.config = {
            'smtp_server': self.smtp_server.get(),
            'smtp_port': int(self.smtp_port.get()),
            'email': self.email.get(),
            'password': self.password.get(),
            'recipients': [r.strip() for r in self.recipients.get().split(',')],
            'report_file': self.report_file.get(),
            'schedule_time': self.schedule_time.get(),
            'subject': self.subject.get()
        }
        self.save_config()
        messagebox.showinfo("Success", "Configuration saved successfully!")
    
    def send_email(self):
        try:
            # Read report data
            if not os.path.exists(self.config['report_file']):
                raise FileNotFoundError("Report file not found")
            
            df = pd.read_csv(self.config['report_file'])
            report_content = df.to_string(index=False)
            
            # Setup email
            msg = MIMEMultipart()
            msg['From'] = self.config['email']
            msg['To'] = ', '.join(self.config['recipients'])
            msg['Subject'] = self.config['subject']
            msg.attach(MIMEText(report_content, 'plain'))
            
            # Connect to SMTP server
            with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                server.starttls()
                server.login(self.config['email'], self.config['password'])
                server.send_message(msg)
                
            logging.info(f"Email sent successfully to {', '.join(self.config['recipients'])}")
            self.root.after(0, lambda: self.status_label.config(text=f"Status: Email sent at {datetime.now().strftime('%H:%M:%S')}"))
        except Exception as e:
            logging.error(f"Failed to send email: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to send email: {str(e)}"))
    
    def test_email(self):
        self.save_config_gui()
        self.send_email()
    
    def run_scheduler(self):
        schedule.every().day.at(self.config['schedule_time']).do(self.send_email)
        while self.running:
            schedule.run_pending()
            time.sleep(60)
    
    def start_scheduler(self):
        if not self.running:
            self.save_config_gui()
            self.running = True
            self.scheduling_thread = threading.Thread(target=self.run_scheduler, daemon=True)
            self.scheduling_thread.start()
            self.start_button.config(text="Stop Scheduler", style="danger.TButton", command=self.stop_scheduler)
            self.status_label.config(text=f"Status: Scheduler running, next at {self.config['schedule_time']}")
        else:
            messagebox.showinfo("Info", "Scheduler is already running!")
    
    def stop_scheduler(self):
        self.running = False
        self.start_button.config(text="Start Scheduler", style="success.TButton", command=self.start_scheduler)
        self.status_label.config(text="Status: Stopped")
        schedule.clear()

if __name__ == "__main__":
    root = ttkb.Window()
    app = AutoEmailReports(root)
    root.mainloop()
