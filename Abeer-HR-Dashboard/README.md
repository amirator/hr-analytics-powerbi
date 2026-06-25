# Abeer Medical Group — HR Analytics Dashboard (PBIP)

A complete, ready-to-open Power BI Project for HR analytics modelled on
**Abeer Medical Group** (Jeddah-based healthcare network, est. 1999, facilities
across Jeddah, Riyadh, Dammam and Makkah).

## How to open

1. Install **Power BI Desktop** (latest version from Microsoft Store).
2. Double-click **`Abeer HR Analytics.pbip`** (or File ▸ Open in Power BI Desktop).
   > If prompted, enable *Preview features ▸ Power BI Project (.pbip) save option*
   > in Options — recent Desktop versions open PBIP natively with no setting needed.
3. Click **Refresh** once so the embedded data loads into the model.
4. Explore. Save As `.pbix` if you want a single-file copy to share.

All data is **embedded in the semantic model** (compressed Enter-Data queries) —
no external files, gateways or credentials are needed. The data is realistic
**synthetic sample data** (680 employee records, ~500 active) seeded to mirror a
KSA healthcare group: nationality mix, Saudization levels, clinical vs
non-clinical departments, SAR salary bands, Nitaqat targets.

## Pages

| Page | What it answers |
|---|---|
| **Executive Overview** | Headcount, hires, attrition, Saudization, payroll at a glance; trend and facility/department breakdowns |
| **Demographics** | Age bands, gender split, nationality, tenure; department scorecard table |
| **Attrition & Retention** | Hires vs terminations by month, attrition trend, reasons, hotspot facilities |
| **Recruitment** | Open requisitions, time-to-fill, offer acceptance, sourcing channels, requisition register |
| **Saudization** | Saudization rate vs the 40% Nitaqat target (gauge), by facility/department, Saudi vs non-Saudi trend |
| **Rewards & Learning** | Salary by grade/department, training hours & completion, leave mix, absenteeism, performance |

Every page carries **Facility / Department / Year** slicers wired through proper
dimension tables, so all visuals cross-filter.

## Model

- 8 tables: `Employees`, `Recruitment`, `Training`, `Leave`, `DimDepartment`,
  `DimFacility`, a DAX `Date` table, and a dedicated `HR Measures` table.
- 36 DAX measures in display folders (Workforce, Attrition, Saudization,
  Recruitment, Compensation, Training, Leave, Performance). Headcount/tenure/payroll
  are **point-in-time** measures (they respect the Date slicer correctly, not just
  row counts).
- Custom **Abeer brand theme** (teal/navy healthcare palette, rounded cards,
  soft shadows) in `Abeer HR Analytics.Report/StaticResources/RegisteredResources/AbeerTheme.json`.

## Excel VBA edition — `Abeer HR Analytics.xlsm`

A self-contained macro-enabled workbook with the same data and a VBA automation
engine. Open it, click **Enable Content**, and use the dashboard buttons:

- **REFRESH DATA** — re-runs the in-memory engine (~53,000 employee-month
  evaluations via Variant arrays + Scripting.Dictionary, ~1 second) and rebuilds
  snapshots, KPI cards, charts, pivots and alerts.
- **REBUILD ALL** — wipes and reconstructs the entire dashboard layout from code.
- **FIND EMPLOYEE** — instant search across ID/name/department/title/facility.
- **EXPORT PDF** — one-click dashboard PDF.
- **ALERTS** — auto-generated compliance register (Saudization gaps, attrition
  ceilings, ageing requisitions, performance watchlist, sick-leave outliers,
  training no-shows) with thresholds driven by the `Config` sheet.
- **PIVOT EXPLORER** — code-generated pivot tables + slicers over the raw data.

Sheets: `Dashboard`, `PivotExplorer`, `Alerts`, `Snapshots` (month x facility x
department fact table built by VBA), raw `Data_*` tables, `Config` (edit the
targets and refresh). VBA source lives in `excel/vba/*.bas`; rebuild the whole
workbook with:

```
python excel/export_data.py
powershell -ExecutionPolicy Bypass -File excel/build_xlsm.ps1
```

(The builder temporarily enables Excel's "Trust access to the VBA project object
model" setting and restores it afterwards.)

## Regenerating / customizing

`build_pbip.py` regenerates the whole project (data, model, report) from scratch:

```
python build_pbip.py
python validate.py   # integrity checks
```

Tune the constants at the top (facilities, departments, salary bands,
headcount, Saudization shares) and re-run. To use **real HRIS data**, replace
each table's partition M query in
`Abeer HR Analytics.SemanticModel/model.bim` with a connector query
(Excel/SQL/API) that returns the same column names and types — the report and
measures will work unchanged.
