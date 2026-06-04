# PDF to Excel Converter — DP World

Convert scanned Invoice (INV) and Packing List (PKL) PDFs into structured Excel files.

## Features
- AI-powered data extraction (Groq Llama 4 Vision — free)
- Cloud storage via Supabase
- Web UI + REST API
- Deploy on Vercel

## Setup

### 1. Clone & install
```bash
git clone https://github.com/tonypham-2003/PDF-Extraction.git
cd PDF-Extraction
pip install -r requirements.txt
```

### 2. Environment variables
Copy `.env.example` to `.env` and fill in:
```
GROQ_API_KEY=gsk_...        # free at console.groq.com
SUPABASE_URL=https://...
SUPABASE_SERVICE_KEY=eyJ...
```

### 3. Run locally
```bash
python app.py
# Open http://localhost:5000
```

## Deploy on Vercel
1. Import this repo on [vercel.com](https://vercel.com)
2. Add environment variables in Project Settings
3. Deploy

## API
`POST /api/convert` — upload PDF + type (INV/PKL) → returns Excel download URL  
`GET  /api/download/<folder>/<filename>` — download converted Excel  
`GET  /api/history` — list recent conversions  
`GET  /api/config` — check AI + Supabase status
