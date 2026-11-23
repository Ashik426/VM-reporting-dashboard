"""Builds dashboard/sample-data.js: an offline snapshot of every source, in the same raw
shape the Apps Script API returns ({sheets:[{name, source, values}]}).
Usage: python build_sample.py <google_Sheet1.csv> <google_Sheet2.csv> <Protistha.xlsx> <out sample-data.js>"""
import csv, datetime, json, sys, openpyxl

def csv_values(path):
    with open(path, encoding="utf-8") as f:
        return [row for row in csv.reader(f)]

def val(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%dT00:00:00")
    return v

sheets = [
    {"name": "Sheet1", "source": "Google Sheet", "values": csv_values(sys.argv[1])},
    {"name": "Sheet2", "source": "Google Sheet", "values": csv_values(sys.argv[2])},
]
wb = openpyxl.load_workbook(sys.argv[3], data_only=True)
for ws in wb.worksheets:
    vals = [[val(c) for c in r] for r in ws.iter_rows(values_only=True)]
    sheets.append({"name": ws.title, "source": "Protistha VM Report 2026.xlsx", "values": vals})

open(sys.argv[4], "w", encoding="utf-8").write(
    "// Offline snapshot (Google Sheet 'VM Dissemination Report (Jan-Current)' + 'Protistha VM Report 2026.xlsx').\n"
    "// Used only when no API_URL is set. Rebuild with tools/build_sample.py.\nwindow.SAMPLE_DATA=" +
    json.dumps({"sheets": sheets, "updated": "snapshot 23 Nov 2026"}, ensure_ascii=False, default=str) + ";\n")
print([(s["name"], len(s["values"])) for s in sheets])
