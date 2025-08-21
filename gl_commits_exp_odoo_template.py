import requests
import openpyxl
import urllib.parse
import os
from dateutil import parser
from dotenv import load_dotenv

# load env 
load_dotenv()

# Config
GITLAB_API = "https://gitlab.com/api/v4"
# PROJECT_ID = urllib.parse.quote_plus("hitdigitalconsulting/py-approval-system-api")
# PROJECT_ID = urllib.parse.quote_plus("hitdigitalconsulting/mobile-approval-ionic")
PROJECT_ID = urllib.parse.quote_plus("hitdigitalconsulting/ebkm-ionic")
BRANCH = "development"
TOKEN = os.getenv("GITLAB_TOKEN")

# Workbook
wb = openpyxl.Workbook()

# Sheet 1: Raw commits
ws1 = wb.active
ws1.title = "Commits"
ws1.append(["sha", "author", "email", "date_time", "message"])

# Sheet 2: Odoo Import format
ws2 = wb.create_sheet("Odoo Import")
ws2.append(["Start Date", "Project", "Task", "Description", "Hours Spent"])

# Request commits
url = f"{GITLAB_API}/projects/{PROJECT_ID}/repository/commits"
params = {
    "ref_name": BRANCH,
    "since": "2025-01-01T00:00:00Z",
    "until": "2025-08-01T23:59:59Z"
}
headers = {"PRIVATE-TOKEN": TOKEN}

r = requests.get(url, headers=headers, params=params)

if r.status_code != 200:
    print("❌ Error:", r.status_code, r.text)
    exit()

commits = r.json()

for c in commits:
    email = c["author_email"].strip().lower()   # normalisasi email
    #filter by email 
    if email == "fryatama@gmail.com":
        sha = c["id"]
        author = c["author_name"]
        raw_date = c["created_at"]
        dt = parser.isoparse(raw_date)
        date_time = dt.strftime("%Y-%m-%d %H:%M:%S")
        message = c["message"].replace("\n", " ")

        # Raw commits (hanya Fryatama yg masuk sheet1)
        ws1.append([sha, author, email, date_time, message])

        # Filter email khusus untuk sheet 2
        ws2.append([
            dt.strftime("%Y-%m-%d"),   # Start Date
            "Internal",                # Project
            "Vallo Approval",          # Task 
            message,                   # Description = commit message
            1                          # Hours Spent (default 1 jam per commit)
        ])

wb.save("gitlab_commits_ebkm_Mobile.xlsx")
print("✅ Export selesai")
