# Project: Inbound Logistics BOD Dashboard

## Purpose
Daily operational dashboard for reporting inbound logistics KPIs to the Board of Directors (BOD).

## Data Source
- **File**: `Data.xlsx`
- **Sheet**: `Logistics Data` — 2108 rows × 27 columns
- **Sheet**: `Summary` — pre-aggregated KPI metrics
- **Date range**: Jan 2026 – Oct 2026 (with Dec 2025 tail)

## Key Columns
| Column | Description |
|--------|-------------|
| Actual Time of Arrival | Vessel arrival at port |
| Customs Clearance Date | Date cargo cleared customs |
| Container Status | Delivered / In Transit / At Port |
| Declaration Status | Correct / Incorrect |
| Customs Line | Green / Yellow / Red channel |
| On Time | Yes / No |
| In Full | Yes / No |
| OTIF | Yes / No (On Time In Full) |
| Carrier | 9 carriers (COSCO, OOCL, Evergreen, ONE, Maersk, CMA CGM, MSC, Hapag-Lloyd, Wan Hai) |
| Plant | 7 plants across Vietnam |
| Destination Port | 5 ports (Hai Phong, Lach Huyen, Cai Mep, Cat Lai, Da Nang) |
| Customer Complaint | No Complaint / Late Delivery / Damaged Cargo / etc. |

## KPI Benchmarks (YTD 2026)
- Total Shipments: 2,108
- OTIF: 93.97%
- On Time: 97.11%
- In Full: 95.92%
- Declaration Accuracy: 97.77%
- Avg Lead Time (Arrival → Clearance): ~3.6 days
- Containers at Port: 50
- In Transit: 132

## Dashboard Style
- **Theme**: Dark (dark navy/charcoal background)
- **Audience**: BOD — professional, high-level, visual
- **Format**: Single-page HTML with Chart.js from CDN
- **Inspiration**: KPI Digital Dashboard – Customs & Delivery Services style

## Dashboard File
- **Output**: `Dashboard.html` in this folder

## Data Update Workflow
- When `Data.xlsx` gets new rows or columns, **re-read the file first**, then regenerate `Dashboard.html` with updated embedded data
- Re-check KPI Benchmarks section and update the values to reflect the latest dataset
- If new columns are added, assess relevance and add new KPI cards or charts as appropriate
- Always re-embed all data directly in HTML after any data change (no stale hardcodes)

## Conventions
- Always embed data directly in HTML (no external data files)
- Use Chart.js from cdnjs.cloudflare.com only
- KPI cards at top, charts below
- Color palette: dark navy bg (#0a0e1a), card bg (#12192c), accent blues/teals/amber
- Months filter should cover Jan–Oct 2026
