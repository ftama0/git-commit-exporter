import requests
import os
from openpyxl import Workbook
from dateutil import parser
from dotenv import load_dotenv

# load env 
load_dotenv()

# --- CONFIG ---
OWNER = "Hasnur-Informasi-Teknologi"
REPO = "ess-hasnurgroup"
BRANCH = "dev-ferry"
TOKEN = os.getenv("GITHUB_TOKEN")

SINCE = "2025-01-01T00:00:00Z"
UNTIL = "2025-08-21T23:59:59Z"

# --- FUNCTION AMBIL SEMUA COMMIT ---
def fetch_commits(owner, repo, branch, since=None, until=None, token=None):
    url = f"https://api.github.com/repos/{owner}/{repo}/commits"
    headers = {"Authorization": f"token {token}"} if token else {}

    page = 1
    commits_all = []

    while True:
        params = {
            "sha": branch,
            "per_page": 100,
            "page": page
        }
        if since:
            params["since"] = since
        if until:
            params["until"] = until

        res = requests.get(url, headers=headers, params=params)
        commits = res.json()

        if isinstance(commits, dict) and "message" in commits:
            print("⚠️ Error dari GitHub API:", commits["message"])
            break

        if not commits:
            break

        commits_all.extend(commits)
        print(f"✅ Page {page} → {len(commits)} commit diambil")

        page += 1

    return commits_all


# --- MAIN ---
commits = fetch_commits(OWNER, REPO, BRANCH, SINCE, UNTIL, TOKEN)

# --- TULIS KE EXCEL ---
wb = Workbook()
ws = wb.active
ws.title = f"Commits-{BRANCH}"

# Header sesuai permintaan
ws.append(["Start Date", "Project", "Task", "Description", "Hours Spent"])

for c in commits:
    # ambil raw date dari GitHub
    raw_date = c["commit"]["author"]["date"]
    # parse ISO8601 ke datetime object
    dt = datetime.strptime(raw_date, "%Y-%m-%dT%H:%M:%SZ")
    # format ulang untuk Odoo
    date = dt.strftime("%Y-%m-%d %H:%M:%S")

    project = "Internal"   # bisa di-hardcode
    task = 'ESS'     # ambil nama branch biar tahu task
    description = c["commit"]["message"].replace("\n", " ")
    hours_spent = 8   # default fix
    
    ws.append([date, project, task, description, hours_spent])

filename = f"commits_custom_{BRANCH}_{SINCE[:10]}_{UNTIL[:10]}.xlsx"
wb.save(filename)

print(f"\n📁 File '{filename}' berhasil dibuat dengan {len(commits)} commit!")
