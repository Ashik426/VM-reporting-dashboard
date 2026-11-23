"""Converts the Protistha VM Report Excel into dashboard/sample-data.js (same shape the API returns).
Usage: python excel_to_sample.py <report.xlsx> <out sample-data.js>"""
import json, sys, datetime, openpyxl

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
ws = wb.worksheets[[w.title for w in wb.worksheets].index("Sheet1")]
rows = list(ws.iter_rows(values_only=True))
hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
def val(v):
    return v.strftime("%Y-%m-%dT00:00:00") if isinstance(v, datetime.datetime) else v
broadcasts = [{h: val(v) for h, v in zip(hdr, r) if h} for r in rows[1:] if any(x not in (None, "") for x in r)]

ben, on = [], False
for r in wb["Sheet2"].iter_rows(values_only=True):
    if r[0] and str(r[0]).strip().lower().startswith("unique beneficiar"):
        on = True; continue
    if on and r[0] and r[1] and isinstance(r[2], (int, float)):
        ben.append({"Project": "Protistha", "District": r[0], "Upazila": r[1], "Unique Beneficiaries": r[2]})
open(sys.argv[2], "w", encoding="utf-8").write(
    "// Offline sample generated from 'Protistha VM Report 2026.xlsx'. Used only when no API URL is set.\nwindow.SAMPLE_DATA=" +
    json.dumps({"broadcasts": broadcasts, "beneficiaries": ben, "updated": "Excel snapshot"}, ensure_ascii=False) + ";\n")
print(len(broadcasts), len(ben))
