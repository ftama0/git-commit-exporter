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
PROJECT_ID = urllib.parse.quote_plus("hitdigitalconsulting/mobile-approval-ionic")
BRANCH = "dev-ferry"
TOKEN = os.getenv("GITLAB_TOKEN")
# Workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Commits"
ws.append(["sha", "author", "email", "date_time", "message"])

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
    sha = c["id"]
    author = c["author_name"]
    email = c["author_email"]
    raw_date = c["created_at"]
    dt = parser.isoparse(raw_date)
    date_time = dt.strftime("%Y-%m-%d %H:%M:%S")
    message = c["message"].replace("\n", " ")

    ws.append([sha, author, email, date_time, message])

wb.save("gitlab_commits.xlsx")
print("✅ Export selesai: gitlab_commits.xlsx")
