from pathlib import Path
import pandas
from helper import correct_inc, correct_rad

FOLDER = Path(__file__).resolve().parent
INFILE = FOLDER / "w3_700-22.xlsx"
OUTFILE = FOLDER / "result.xlsx"


interview = "Interview-Nummer (fortlaufend)"
condition = "Fragebogen, der im Interview verwendet wurde"
fieldnames = [
    "Field_{}: Ausweichoption (negativ) oder Anzahl ausgewählter Optionen",
    "Field_{}: Tiere & Biologie",
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

def calc_hitrate(df, index, fieldnum, correct):
    try:
        field = fieldnames[0].format(fieldnum)
        select_count = int(df[field][index])
        correct = correct[fieldnum]
        select = []
        for idx, field in enumerate(fieldnames[1:], 1):
            fname = field.format(fieldnum)
            if df[fname][index] == 2:
                select.append(idx)
        cor_count = 0
        for sel in select:
            if sel in correct:
                cor_count += 1
        return float(cor_count / select_count)
    except ValueError:
        #return None
        return "#NULL!"

df = pandas.read_excel(INFILE)
print(df.columns)

result = []
for ind in df.index:
    inr = df[interview][ind]
    cnd = df[condition][ind]
    if cnd == "I":
        correct = correct_inc
    elif cnd == "R":
        correct = correct_rad
    else:
        raise ValueError("Unknown condition")
    data = [inr]
    for i in range(1, 51):
        data.append(calc_hitrate(df, ind, i, correct))
        if type(data[-1]) is float:
            print(f"Interview {inr} Field {i:02} Hitrate: {data[-1]}")
    result.append(data)

colnames = [interview] + [f"hitrate_{i}" for i in range(1, 51)]
outframe = pandas.DataFrame(result, columns=colnames)
outframe.to_excel(OUTFILE, index=False)