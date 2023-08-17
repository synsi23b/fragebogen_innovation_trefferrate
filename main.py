from pathlib import Path
import openpyxl

FOLDER = Path(__file__).resolve().parent
INFILE = FOLDER / "w3_700-22.xlsx"
OUTFILE = FOLDER / "result.xlsx"

inwb = openpyxl.load_workbook(INFILE)
inws = inwb.active


print(inws["A1:ABX1"])

fieldnames = [
    "Field_{}: Ausweichoption (negativ) oder Anzahl ausgewählter Optionen",
    "Field_{}: Tiere &amp; Biologie",
    "Field_{}: Technologie",
    "Field_{}: Kleidung",
    "Field_{}: Energie",
    "Field_{}: Essen",
    "Field_{}: Gartenarbeit",
    "Field_{}: Gesundheit / Medizin",
    "Field_{}: Haushalt",
    "Field_{}: Lernen",
    "Field_{}: Beleuchtung",
    "Field_{}: App oder Intelligenz",
    "Field_{}: Nanotechnologie",
    "Field_{}: Sport",
    "Field_{}: andere"
    ]

