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
TOKEN = os.getenv("GITHUB_TOKEN")            # isi dengan token kalau repo private, contoh: "ghp_xxxxxx"

# --- FETCH DATA DARI GITHUB API ---
url = f"https://api.github.com/repos/{OWNER}/{REPO}/commits"
headers = {"Authorization": f"token {TOKEN}"} if TOKEN else {}

res = requests.get(url, headers=headers)
commits = res.json()

# --- CEK RESPONSE ---
if isinstance(commits, dict) and "message" in commits:
    print("⚠️ Error dari GitHub API:", commits["message"])
    exit()

# --- TULIS KE EXCEL ---
wb = Workbook()
ws = wb.active
ws.title = "Commits"
ws.append(["SHA", "Author", "Email", "Date", "Message"])

for c in commits:
    sha = c["sha"]
    author = c["commit"]["author"]["name"]
    email = c["commit"]["author"]["email"]
    date = c["commit"]["author"]["date"]
    message = c["commit"]["message"].replace("\n", " ")
    ws.append([sha, author, email, date, message])

wb.save("commits_report.xlsx")
print("✅ commits_report.xlsx berhasil dibuat!")
