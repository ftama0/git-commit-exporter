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


# Buat workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Odoo Import"

# Tambahkan header sekali
ws.append(["Start Date", "Project", "Task", "Description", "Hours Spent"])

for c in commits:
    email = c["author_email"].strip().lower()
    
    if email == "fryatama@gmail.com":   # filter khusus Fryatama
        raw_date = c["created_at"]
        dt = parser.isoparse(raw_date)
        message = c["message"].replace("\n", " ")

        ws.append([
            dt.strftime("%Y-%m-%d %H:%M:%S"),   # Start Datetime
            "Internal",                # Project
            "Vallo Approval",          # Task
            message,                   # Description (commit message)
            8                          # Hours Spent (default 8 jam / commit)
        ])

# Simpan file
wb.save("gitlab_commits_ebkm_Mobile.xlsx")
print("✅ Export selesai")