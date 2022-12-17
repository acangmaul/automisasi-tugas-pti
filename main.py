from datetime import date
from gspread.auth import Path
from gspread.spreadsheet import ExportFormat
from typing import List, Any
import gspread
import locale
import sys


locale.setlocale(locale.LC_ALL, "id-ID" if sys.platform == "win32" else "id_ID.UTF-8")


def flatten(l: List[List[Any]]) -> List:
    return [item for sublist in l for item in sublist]


def main():
    service_account = gspread.service_account(filename=Path("creds.json"))
    doc = service_account.open_by_key("1glDtHOViEcNfRVGuItSiuYv7iguWBaOQUFKXpI46Zug")

    sheet = doc.sheet1

    range_tanggal = "A3:F3"
    tanggal = flatten(sheet.get(range_tanggal))

    assert len(tanggal) == 1, "Format tanggal tidak cocok"

    tanggal = tanggal[0]

    with open(f"LAPORAN {tanggal}.pdf", "wb") as f:
        pdf = doc.export(ExportFormat.PDF)
        f.write(pdf)

    sheet.update(range_tanggal, date.today().strftime("%A, %d %B %Y").upper())


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
