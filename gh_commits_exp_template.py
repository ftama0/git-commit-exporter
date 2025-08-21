import requests
import os
from openpyxl import Workbook
from dateutil import parser
from dotenv import load_dotenv

# load env 
load_dotenv()
# --- CONFIG ---
OWNER = "Hasnur-Informasi-Teknologi"          # ganti username org/repo kamu
REPO = "ess-hasnurgroup"            # nama repo
BRANCH = "dev-ferry"             # ganti
TOKEN = os.getenv("GITHUB_TOKEN")            # isi dengan token kalau repo private, contoh: "ghp_xxxxxx"

# filter tanggal (ISO 8601, UTC)
SINCE = "2025-01-01T00:00:00Z"   # atau None kalau tidak dipakai
UNTIL = "2025-08-21T23:59:59Z"   # atau None kalau tidak dipakai

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

        if not commits:  # kalau kosong, berhenti
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
ws.append(["SHA", "Author", "Email", "Date", "Message"])

for c in commits:
    sha = c["sha"]
    author = c["commit"]["author"]["name"]
    email = c["commit"]["author"]["email"]
    date = c["commit"]["author"]["date"]
    message = c["commit"]["message"].replace("\n", " ")
    ws.append([sha, author, email, date, message])

filename = f"commits_{BRANCH}_{SINCE[:10]}_{UNTIL[:10]}.xlsx"
wb.save(filename)

print(f"\n📁 File '{filename}' berhasil dibuat dengan {len(commits)} commit!")