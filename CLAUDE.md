# Daily Report — PDF-to-Excel Web App

## Purpose

Local Flask web app for the DP World daily reporting workflow. Reads PDF files placed in
`Input/AN/`, `Input/INV/`, `Input/PKL/`, extracts structured line-item data, and writes
Excel files into the matching `Output/` sub-folder following the exact column layout of
`template.xls`.

---

## Folder Layout

```
Daily Report/
├── app.py                         # Flask entry point — run this
├── extractor.py                   # PDF → list-of-dicts (text path + Claude vision path)
├── excel_writer.py                # list-of-dicts → .xlsx (matches template.xls)
├── convert_invoice.py             # Legacy CLI reference only — do not extend
├── mapping rule.xlsx              # Reference guide: PDF label → Excel column name
├── template.xls                   # Output format reference (columns, sheet name)
├── templates/
│   └── index.html                 # Single-page web UI (vanilla JS, no build step)
├── Input/
│   ├── AN/                        # Arrival Notice PDFs  → Output/AN/
│   ├── INV/                       # Commercial Invoice PDFs → Output/INV/
│   └── PKL/                       # Packing List PDFs   → Output/PKL/
│   └── Charge Description Invoice/ # Legacy training samples (do not delete)
└── Output/
    ├── AN/
    ├── INV/
    └── PKL/
```

---

## Running the App

```powershell
# 1. Set API key (required for scanned/image-based PDFs in INV and PKL)
$env:ANTHROPIC_API_KEY = "sk-ant-api03-..."

# 2. Start server
cd "c:\Users\v1411\OneDrive - DP World\Desktop\Daily Report"
python app.py

# 3. Open browser
start http://localhost:5000
```

---

## Install Dependencies

```powershell
pip install flask pdfplumber openpyxl anthropic pymupdf
```

| Package | Purpose |
|---|---|
| `flask` | Web server |
| `pdfplumber` | Extract text from text-based PDFs |
| `pymupdf` (fitz) | Render scanned PDF pages to PNG for vision API |
| `anthropic` | Claude vision API — reads scanned/image PDFs |
| `openpyxl` | Write `.xlsx` output |

---

## Web App Routes

| Route | Method | Description |
|---|---|---|
| `/` | GET | Serves `index.html` |
| `/api/status` | GET | Returns per-folder PDF list, job statuses, output files |
| `/api/process/<folder>` | POST | Kicks off background conversion of all PDFs in that folder |
| `/api/process/all` | POST | Kicks off conversion across all three folders |
| `/api/download/<folder>/<filename>` | GET | Streams `.xlsx` file as download |

The UI polls `/api/status` every 4 seconds and shows live job chips
(`Pending → Processing… → Done / Warning / Error`) without page reload.

---

## Excel Output Format

Strictly matches `template.xls` Sheet1 layout:

| Col | Header | Notes |
|---|---|---|
| A | STT | Sequential row number (int, auto-incremented by writer) |
| B | Mã hàng  | Product / item code or model number (string) |
| C | Mã HS | HS / customs tariff code (string, blank if absent) |
| D | Tên hàng | Full goods description (string, left-aligned, wrap) |
| E | Xuất xứ | 2-letter ISO country of origin (string, blank if absent) |
| F | Số lượng 1 | Primary quantity (string) |
| G | Đơn vị tính 1 | Primary UOM — kg, pcs, UNITS, CBM, SET … (string) |
| H | Số lượng 2 | Secondary quantity (string, blank unless source provides it) |
| I | Đơn vị tính 2 | Secondary UOM (string, blank unless source provides it) |
| J | Đơn giá | Unit price, kept as **string** with original comma formatting |
| K | Tổng trị giá | Line total, stored as **float** (numeric cell, matches template ctype=2) |

**Sheet name**: `Sheet1`
**Row 1**: Blue (`#1F4E79`) header row, white bold text, thin border
**Row 2+**: Data rows, thin border, row height 35, frozen pane at A2
**Output file naming**: `<pdf_stem>.xlsx` (same stem as source PDF)

---

## Mapping Rule (`mapping rule.xlsx`)

This file is a **reference guide only** — it documents which PDF labels correspond to which
Excel columns. It is not loaded at runtime.

| Excel column | Look for in PDF |
|---|---|
| Mã hàng | Product Code / Commodity Code / Part No / Item Code / SKU / Model |
| Mã HS | HS code / Tariff code / Customs Tariff XXXXXX |
| Tên hàng | Description of goods / Good description / Product name / Designation |
| Số lượng 1 | Quantity / Qty / Net Weight (PKL) |
| Đơn vị tính 1 | Unit / UOM / PCS / SET / kg / UNITS |
| Số lượng 2 | Leave blank unless source explicitly provides a second measure |
| Đơn vị tính 2 | Leave blank unless source explicitly provides a second measure |
| Đơn giá | Unit price / Rate (before VAT for AN) |
| Tổng trị giá | Amount / Total amount / Line amount / Total after VAT (AN) |

---

## Document Types

### AN — Arrival Notice

- **Format**: Text-based PDF (pdfplumber extracts text directly — no API call needed)
- **Parser**: `_parse_an_text()` in `extractor.py` — regex on charge lines
- **Charge line pattern**: `<NAME>  <QTY><UNIT>  <RATE>  <N>%VAT  <TOTAL>  (USD)`
- **Field mapping**:
  - `ten_hang` = charge name (CFS CHARGE, THC FEE, DOCUMENT FEE, …)
  - `so_luong_1` / `don_vi_tinh_1` = quantity + unit (CBM, SHIPMENT, BL)
  - `don_gia` = rate **before** VAT
  - `tong_tri_gia` = total **after** VAT
  - `ma_hang`, `ma_hs`, `xuat_xu` = always blank for AN
- **Fallback**: if regex yields nothing, falls through to Claude vision

**Known AN subtypes**

| Issuer | Marker | Notes |
|---|---|---|
| ASPRESS Shipping | `THÔNG BÁO HÀNG ĐẾN` / `Arrival Notice` | Text-based; H-B/L in header |

---

### INV — Commercial Invoice

- **Format**: Scanned image PDF — Claude vision API required
- **Parser**: `_call_claude_vision()` with `_PROMPTS["INV"]` in `extractor.py`
- **Field mapping**:
  - `ma_hang` = Product Code column or Model column
  - `ma_hs` = HS code / Customs Tariff number (often embedded in description)
  - `ten_hang` = Description / product name
  - `xuat_xu` = Country of origin (2-letter ISO; look for "CERTIFY THAT GOODS OF X ORIGIN")
  - `so_luong_1` / `don_vi_tinh_1` = Quantity + Unit
  - `don_gia` = Unit Price column
  - `tong_tri_gia` = Amount / Line Total column

**Known INV subtypes**

| Issuer | Marker | Structure |
|---|---|---|
| Robertet Asia (SG) | ROBERTET logo, GST reg 201411292W | Single product block; Customs Tariff buried in description |
| Cooltec (TH) | COOLTEC COMMERCIAL REFRIGERATION header | Table: ITEM / DESCRIPTION / MODEL / QUANTITY / UNIT PRICE / AMOUNT |

---

### PKL — Packing List

- **Format**: Scanned image PDF — Claude vision API required
- **Parser**: `_call_claude_vision()` with `_PROMPTS["PKL"]` in `extractor.py`
- **Field mapping**:
  - `ma_hang` = Product code column
  - `ten_hang` = Designation / product description
  - `so_luong_1` = Net Weight (numeric string)
  - `don_vi_tinh_1` = Weight unit (kg)
  - `tong_tri_gia` = Gross Weight (numeric string)
  - `don_gia`, `ma_hs`, `xuat_xu`, `so_luong_2`, `don_vi_tinh_2` = blank
- **Do NOT include** total/summary rows — one row per physical package

**Known PKL subtypes**

| Issuer | Marker | Notes |
|---|---|---|
| Robertet Asia (SG) | ROBERTET logo, PACKING LIST header | Per-drum rows with Batch Nr, Net/Gross Weight |

---

## Extractor Logic (`extractor.py`)

```
extract(pdf_path, doc_type)
  ├─ doc_type == "AN" and PDF has text (>50 chars)
  │    └─ _parse_an_text()  →  regex on charge lines
  │         └─ if no items found → fall through to vision
  └─ all other cases
       └─ _pages_as_base64_png()  →  fitz renders at 200 dpi
            └─ _call_claude_vision()  →  claude-opus-4-5, max_tokens=4096
                 └─ parses JSON array from response
```

Parser return type — list of dicts with **exactly** these keys:
```python
{
    "ma_hang": str,       # product/item code
    "ma_hs": str,         # HS tariff code ("" if absent)
    "ten_hang": str,      # goods description
    "xuat_xu": str,       # 2-letter ISO origin ("" if absent)
    "so_luong_1": str,    # primary quantity
    "don_vi_tinh_1": str, # primary UOM
    "so_luong_2": str,    # secondary quantity ("" if absent)
    "don_vi_tinh_2": str, # secondary UOM ("" if absent)
    "don_gia": str,       # unit price string (keep comma formatting)
    "tong_tri_gia": str,  # line total string (writer converts to float)
}
```

---

## Key Conventions

- **Never use `xlrd` / `xlwt`** — `template.xls` is reference-only; all output is `.xlsx` via `openpyxl`.
- **`tong_tri_gia`** is written as a float (numeric cell) to match template.xls `ctype=2`.
- **`don_gia`** is written as a string to preserve comma formatting (e.g. `"2,349.00"`).
- **Do not add a merged title row** above the headers — template.xls starts directly with the header row.
- **Sheet name must be `Sheet1`** (matches template.xls).
- **Output file name** = `<pdf_stem>.xlsx` placed in `Output/<FOLDER>/`.
- **Do not include grand-total / summary rows** from the source PDF in the item list.
- **Số lượng 2 / Đơn vị tính 2** — leave blank unless the source document explicitly provides a secondary quantity/unit for the same line.
- **ANTHROPIC_API_KEY** must be set in the environment before starting `app.py` for INV/PKL processing.

---

## Verified Sample Conversions

| Source PDF | Output | Rows | Notes |
|---|---|---|---|
| `Input/AN/AN_1.pdf` | `Output/AN/AN_1.xlsx` | 5 | Text-based; 5 charges (CFS, CIC, THC, DOCUMENT, HANDLING) |
| `Input/INV/INV_1.pdf` | `Output/INV/INV_1.xlsx` | 1 | Robertet; product 848552; HS 330210 |
| `Input/INV/INV_2.pdf` | `Output/INV/INV_2.xlsx` | 3 | Cooltec; total 520,772 USD verified against amount-in-words |
| `Input/PKL/PKL_1.pdf` | `Output/PKL/PKL_1.xlsx` | 4 | Robertet; 4 drums × 25 kg net / 26.54 kg gross |
