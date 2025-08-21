from openpyxl import Workbook

# data dummy
data = [
    ["ID", "Nama", "Email"],
    [1, "Hecatte", "hecatte@example.com"],
    [2, "Alice", "alice@example.com"],
    [3, "Bob", "bob@example.com"],
]

# buat workbook baru
wb = Workbook()
ws = wb.active
ws.title = "Report"

# tulis data
for row in data:
    ws.append(row)

# simpan ke file
wb.save("dummy_report.xlsx")

print("File dummy_report.xlsx berhasil dibuat!")
