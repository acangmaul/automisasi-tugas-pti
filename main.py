from datetime import datetime, date
from gspread.spreadsheet import ExportFormat
import gspread
import sys
import locale

locale.setlocale(locale.LC_ALL,
                 "id-ID" if sys.platform == "win32" else "id_ID.UTF-8")


def main():
    service_account = gspread.service_account(filename="creds.json")
    doc = service_account.open_by_key("1CKSNp_cAqZ7wIO-SDjOqh7FqrBSEF6L2VbE3Dz19X7E")

    with open(f"Laporan {date.today().strftime('%d-%b-%Y')}.pdf", "wb") as f:
        pdf = doc.export(ExportFormat.PDF)
        f.write(pdf)

    # kalau jam 12 malam
    # update tanggal dan kosongkan spreadsheet
    if datetime.now().hour == 0:
        sheet = doc.sheet1

        range_tanggal = "A3:F3"
        range_nama_npm = "B5:C44"

        sheet.update(range_tanggal, date.today().strftime("%A, %d %B %Y").upper())
        sheet.batch_clear(range_nama_npm)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)