from pathlib import Path
import openpyxl

FOLDER = Path(__file__).resolve().parent
INFILE = FOLDER / "w3_700-22.xlsx"
OUTFILE = FOLDER / "result.xlsx"

inwb = openpyxl.load_workbook(INFILE)
inws = inwb.active


print(inws["A1:ABX1"])

fieldnames = [
    "Field_1: Ausweichoption (negativ) oder Anzahl ausgewählter Optionen",
    "Field_1: Tiere &amp; Biologie",
    "Field_1: Technologie",
    "Field_1: Kleidung",
    "Field_1: Energie",
    "Field_1: Essen",
    "Field_1: Gartenarbeit",
    "Field_1: Gesundheit / Medizin",
    "Field_1: Haushalt",
    "Field_1: Lernen",
    "Field_1: Beleuchtung",
    "Field_1: App oder Intelligenz",
    "Field_1: Nanotechnologie",
    "Field_1: Sport",
    "Field_1: andere"
    ]