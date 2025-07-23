✅ How to Use the "Auto Email Reports" Tool
Professional Setup Guide for Windows (PowerShell)

📥 Step 1: Download the Project ZIP from GitHub
Open your browser and go to:
👉 https://github.com/adarshmishra-tech/auto_email_reports

Click the green Code button → Click Download ZIP.

Extract the downloaded ZIP file.
You should now have a folder like this:

makefile
Copy
Edit
D:\auto_email_reports-main\auto_email_reports-main\
📂 Step 2: Open the Project Folder in PowerShell
Press Shift + Right-click in the extracted folder.

Select Open PowerShell window here.

Or run manually:

powershell
Copy
Edit
cd "D:\auto_email_reports-main\auto_email_reports-main"
📄 Step 3: Check Files Are Present
Run:

powershell
Copy
Edit
ls
You should see these files:

auto_email_reports.py

.env

requirements.txt

email_reports.log

README.md

LICENSE

🧪 Step 4: Create & Activate a Virtual Environment (Optional but Recommended)
powershell
Copy
Edit
python -m venv venv
.\venv\Scripts\activate
This keeps your dependencies isolated for this project only.

📦 Step 5: Install Required Python Packages
Install from requirements.txt:

powershell
Copy
Edit
pip install -r requirements.txt
If .env usage fails, install python-dotenv:

powershell
Copy
Edit
pip install python-dotenv
📝 Step 6: Configure Email Credentials in .env File
Open .env file in Notepad:

powershell
Copy
Edit
notepad .env
Add your email settings like this:

ini
Copy
Edit
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
EMAIL=your_email@gmail.com
PASSWORD=your_app_password
🔐 Use an app password if using Gmail with 2FA: https://myaccount.google.com/apppasswords

🟢 Step 7: Run the App with GUI
Run this command:

powershell
Copy
Edit
python auto_email_reports.py
If everything is configured correctly, a Tkinter-based GUI will launch.

🧰 Step 8: Use the GUI
In the GUI, you can:

Set SMTP Server, Port

Input your Email & Password

Add Recipients

Select your CSV Report File

Add Attachments (optional)

Set Schedule Time (e.g., 09:00)

Define the Email Subject

📨 Step 9: Send Test Email
Click Send Test Email to verify your setup.
Check the "Status" label for success/failure messages.

⏱ Step 10: Start the Email Scheduler
Click Start Scheduler to begin daily email automation.

The app will run in background and send emails every day at the scheduled time.

🧾 Logs
Check the email_reports.log file in the folder for:

Email delivery logs

Any error messages

Debug information

🧠 Bonus: Troubleshooting
Problem	Solution
ModuleNotFoundError: dotenv	Run pip install python-dotenv
Password not working (Gmail)	Use an App Password if 2FA is enabled
GUI doesn't open	Make sure tkinter, ttkbootstrap, and pandas are installed properly
Emails not sending	Check SMTP server settings and internet/firewall

✅ You're now fully set up to use the Auto Email Reports Tool!

Would you like me to generate this into a formatted Word or PDF file? Just say:
Make PDF or Make Word.
