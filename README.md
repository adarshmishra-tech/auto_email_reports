git clone https://github.com/yourusername/auto-email-reports.git
 cd auto-email-reports
 pip install -r requirements.txt
 Create a .env file in the root:
create .env in same folder and insert these,
EMAIL_USER=youremail@example.com
EMAIL_PASS=yourpassword
python auto_email_reports.py
