# DATAVIS v5 ULTIMATE

Data analysis & visualization suite for students, academic work, and office use.
Built with Streamlit + Plotly.
Coded by **Arun**, a.k.a. **KIRITO_SERFORT**.

## Run it
```
pip install -r requirements.txt
streamlit run app.py
```
Or use the desktop launcher (requires PyQt5):
```
python launcher.py
```

## What's new in v5
- **ℹ️ About tab** — creator credit (Arun / KIRITO_SERFORT), version history, what the
  suite is for, and a feature-highlights list, all inside the app.
- Refreshed header with a byline badge and matching hero/pill styling for the About tab.
- Everything renamed from v4 → v5 throughout the app (window title, page title,
  exported report filenames, launcher window/labels).

## What's new in v5 (carried over from v3)
- **🏠 Dashboard tab** — KPI stat row (rows, columns, numeric metrics, missing
  cells) plus free, instant "Auto Insights" bullets generated straight from the
  data (no API key needed), a file-tag overview when multiple files are loaded,
  and a one-click quick chart.
- **🧹 Data Cleaning tab** — remove duplicate rows, trim whitespace, drop
  fully-empty columns, fill missing values (mean/median/zero/forward-fill) per
  column, and an IQR-based outlier detector that flags suspicious rows before
  you chart them. Every change is logged, and the cleaned data downloads as CSV.
- **📤 Export Report tab** — bundles every uploaded file's raw data, summary
  statistics, and the correlation matrix into a single multi-sheet Excel
  workbook.
- **🔎 Searchable Raw Data tab** — a global search box filters every loaded
  table (still labeled and separated by source file) at once, with a per-file
  CSV download button.
- **🎨 Theme presets** — five color themes (Midnight Violet, Deep Ocean,
  Charcoal Rose, Pure Light, Warm Paper) plus a compact-layout toggle, on top
  of the existing dark/light mode and chart color palettes.
- **Launcher** — now titled "DATAVIS v3", auto-opens the app in your browser
  once the server is ready, and has a manual "Open in Browser" button.

## Existing capabilities (carried over from v2.5)
- CSV / Excel / PDF (table-extraction) uploads, multiple files at once.
- Multi-file column-compatibility check with a picker for shared columns when
  headers don't match exactly.
- Explore & Customize: 9 chart types, radar chart, 3D scatter explorer.
- Stats & Correlation: describe(), missing-value chart, correlation heatmap,
  distribution explorer.
- Personal Analysis (per-entity comparison vs. the rest of the dataset).
- Head-to-Head comparison between two entities.
- Multi-File Growth tracking across semesters/periods.
- Compound growth Projection calculator.
- Optional AI Insights tab (Claude API key required).

## Notes on file compatibility (multi-file analysis)
When you upload more than one file, the **Multi-File Growth** tab automatically
checks whether every file shares the exact same columns. If they don't match
(different headers or column counts), you're prompted to pick which columns are
common and meaningful across all of them before anything is combined — nothing
is auto-merged silently.
