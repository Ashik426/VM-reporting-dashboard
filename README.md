# VM Reporting Dashboard

A free, browser-based reporting platform for **voice message (VM) dissemination** projects in Bangladesh.
It reads every tab of a Google Sheet (and any Excel/CSV you import), cleans the data, and produces filtered
reports, charts and a project area map. Reports download as Excel, PDF or Word in the same layout as the
original workbook.

Projects covered: GFFO, RIMES, UNICEF, ECHO, Caritas, Protistha, PROVATi3, Concern, Enrich, and any new
project added to the sheet.

```
Google Sheet (every tab)  ⇄  Apps Script Web App (free API)  ⇄  Dashboard (dashboard/index.html)
Excel / CSV import  ─────────────────────────────────────────▶  filters · map · charts · exports
```

No server, database or paid service is needed. The dashboard is a static web page, and the API is a free
Google Apps Script web app attached to the sheet.

---

## Features

| Feature | Details |
|---|---|
| **Filters** | Project, timeframe (month, quarter, monsoon, last 30/90 days, custom range), zila, upazila, VM content, media |
| **Overview** | KPIs (broadcasts, total dissemination, success rate, zilas covered), monthly trend, calls by zila, media and project charts |
| **Dissemination log** | Same 9 columns as the Google Sheet: Date, Districts, Upazila, Message content, Total Dissimination, Success, Failed, Project, Disseminated Media |
| **Summary** | Same layout as the summary sheet in the Excel workbook: Total Broadcasts, Total Calls, calls by zila/upazila, Unique Beneficiary block, per-project block with success rate |
| **Exports** | Excel (`Dissemination log` + `Summary` sheets), PDF and Word; you can add the map and charts to the PDF and Word files |
| **Project area map** | GIS-style layout: upazilas named in broadcasts are filled blue over grey district boundaries, with a legend, north arrow, scale bar and lat/long frame. Click an area to filter. Download as PNG (A4 at 300 dpi) or GeoJSON |
| **Auto-correction** | Rain / Raainfall → *Rainfall*, Lightening → *Lightning*, CTG → *Chittagong*, Bagura → *Bogra*, Sador → *Sadar*, Unicef → *UNICEF*, upazila names typed in the zila column, and more. Uses a gazetteer of 64 zilas and 544 upazilas |
| **De-duplication** | A broadcast logged in two tabs (same date, total, success, failed, project and media) is counted once |
| **Data check** | Data tab lists every source, every auto-fix, duplicates removed and rows needing attention |
| **Data entry** | Add-a-broadcast form, Excel/CSV import with "Save to Google Sheet", a short URL API, and unique-beneficiary entry |
| **Themes** | Light and dark mode |

---

## Quick start (offline preview)

The dashboard runs on a built-in data snapshot until it is connected to the Google Sheet.

```bash
python -m http.server 8765 --directory dashboard
# open http://localhost:8765
```

Opening `dashboard/index.html` directly also works.

## Connect to the Google Sheet (about 10 minutes)

1. Open the Google Sheet, then go to **Extensions → Apps Script**. Paste in `apps-script/Code.gs`.
   In ⚙ Project Settings, tick *Show "appsscript.json"* and paste in `apps-script/appsscript.json`.
2. Run **`setup`** once and approve the permissions. It creates a **Beneficiaries** tab and an **API key**
   (shown in the Execution log).
3. Go to **Deploy → New deployment → Web app**. Set *Execute as* **Me** and *Who has access* **Anyone**.
   Copy the `/exec` URL.
4. Put that URL in `dashboard/config.js` → `API_URL`. The status pill turns green ("Live"), and the
   dashboard re-reads the sheet every `REFRESH_MINUTES`.

New rows go to the tab named `WRITE_SHEET` (default `Sheet2`, the all-projects log). After you edit
`Code.gs`, use **Deploy → Manage deployments → Edit → New version**. The URL stays the same.

## Using Excel files

The dashboard does not stay linked to an Excel file. Excel is used in two ways:

- **Import:** drop an `.xlsx`/`.csv` on the **Data** tab. Every sheet is read, and the rows are cleaned and
  added to the dashboard straight away. Press **Save to Google Sheet** to store them permanently. Until you
  save, imported rows are lost when the page reloads.
- **Export:** **Excel** downloads a workbook in the original layout (`Dissemination log` + `Summary`).

The Google Sheet is the single live source. To add a project kept in Excel (for example Protistha), import
its file once and save it to the sheet, or paste its rows into a new tab named after the project.

## Writing data: short API

```
# same order as the sheet: date, district, upazila, content, total, success, failed, project, media
<URL>?action=add&key=KEY&q=23-11-2026,Sylhet,Kanaighat,Lightning,1005,454,551,UNICEF,Kotha

# short form: date, upazila, content, total, media, project
<URL>?action=add&key=KEY&q=23-11-2026,Ulipur,Flood,2579,GP,Protistha

# unique beneficiaries
<URL>?action=beneficiary&key=KEY&project=Protistha&upazila=Ulipur&count=7303

# read everything (used by the dashboard)
<URL>?action=data
```

Several rows can be sent at once with a POST (`Content-Type: text/plain`):
`{"key":"KEY","rows":[{...} | "csv line"]}`.

**Security:** writing needs the API key. Anyone with the `/exec` URL can read the data. To restrict reading,
set the deployment access to your organisation (viewers then need a Google sign-in). Never commit the API key.

---

## Project structure

```
apps-script/Code.gs          Free API: returns every tab; appends rows in the sheet's column order
apps-script/appsscript.json  Manifest (Asia/Dhaka time zone, web-app access)
dashboard/index.html         The dashboard (open this)
dashboard/app.js             Parsing, auto-correction, de-duplication, filters, map, exports
dashboard/config.js          Settings: API URL, refresh interval, title, organisation
dashboard/geo.js             64 zilas + 544 upazilas + neighbouring countries
dashboard/sample-data.js     Offline data snapshot
tools/build_geo.py           Rebuilds geo.js from geoBoundaries ADM2/ADM3
tools/build_neighbours.py    Adds neighbouring-country land (Natural Earth) to geo.js
tools/build_sample.py        Rebuilds the snapshot from CSV exports + the Protistha Excel
tools/snapshot/              CSV copies of the Google Sheet tabs
```

## Built with

[Leaflet](https://leafletjs.com) · [Chart.js](https://www.chartjs.org) · [jsPDF](https://github.com/parallax/jsPDF) + AutoTable ·
[docx](https://docx.js.org) · [SheetJS](https://sheetjs.com) · Google Apps Script. Libraries load from public CDNs,
so there is no build step.

## Data sources and credits

- Administrative boundaries: [geoBoundaries](https://www.geoboundaries.org) BGD ADM2/ADM3 (CC BY 4.0)
- Neighbouring countries: [Natural Earth](https://www.naturalearthdata.com) 1:50m (public domain)
- Basemap: © [OpenStreetMap](https://www.openstreetmap.org/copyright) contributors
- Dissemination data: RIMES VM dissemination logs
