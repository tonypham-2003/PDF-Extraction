"""extractor.py — detect PDF type and extract structured line items.

AI backend (first configured key wins):
  1. ANTHROPIC_API_KEY → Claude Opus 4        (paid)
  2. GROQ_API_KEY      → Llama 3.2 Vision     (free, 7,000 req/day)
  3. GEMINI_API_KEY    → Gemini 2.0 Flash     (free, 1,500 req/day)
"""
import os, re, json, base64
import pdfplumber
import fitz          # PyMuPDF


# ── AI client helpers ─────────────────────────────────────────────────────────

def _env(key: str) -> str:
    """Get env var, stripping BOM + whitespace that PowerShell/Windows can add."""
    return os.environ.get(key, "").strip().lstrip("﻿​")


def _active_backend() -> str:
    """Return 'claude', 'groq', 'gemini', or raise if nothing is configured."""
    key_ant = _env("ANTHROPIC_API_KEY")
    key_grq = _env("GROQ_API_KEY")
    key_gem = _env("GEMINI_API_KEY")

    if key_ant.startswith("sk-ant-") and len(key_ant) > 80:
        return "claude"
    if key_grq.startswith("gsk_") and len(key_grq) > 20:
        return "groq"
    if key_gem.startswith("AIza") or (key_gem.startswith("AQ.") and len(key_gem) > 20):
        return "gemini"
    raise RuntimeError(
        "No valid AI key found. Set one of these in .env:\n"
        "  GROQ_API_KEY=gsk_...      (free, get at console.groq.com)\n"
        "  ANTHROPIC_API_KEY=sk-ant-... (paid)"
    )


_claude_client = None
def _claude():
    global _claude_client
    if _claude_client is None:
        import anthropic
        _claude_client = anthropic.Anthropic(api_key=_env("ANTHROPIC_API_KEY"))
    return _claude_client


_gemini_client = None
def _gemini():
    global _gemini_client
    if _gemini_client is None:
        from google import genai
        _gemini_client = genai.Client(api_key=_env("GEMINI_API_KEY"))
    return _gemini_client


# ── PDF helpers ───────────────────────────────────────────────────────────────

def _has_text(pdf_path: str) -> bool:
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if len((page.extract_text() or "").strip()) > 50:
                return True
    return False


def _full_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def _pages_as_base64_png(pdf_path: str, dpi: int = 200) -> list[str]:
    doc = fitz.open(str(pdf_path))
    result = []
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        result.append(base64.standard_b64encode(pix.tobytes("png")).decode())
    return result


def _clean_json(raw: str) -> list[dict]:
    """Strip markdown fences and parse JSON array."""
    raw = re.sub(r"^```[^\n]*\n?", "", raw.strip())
    raw = re.sub(r"\n?```$",       "", raw)
    return json.loads(raw.strip())


# ── Claude vision ─────────────────────────────────────────────────────────────

def _call_claude(images_b64: list[str], prompt: str) -> list[dict]:
    content = [
        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img}}
        for img in images_b64
    ]
    content.append({"type": "text", "text": prompt})

    resp = _claude().messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    return _clean_json(resp.content[0].text)


# ── Groq vision ───────────────────────────────────────────────────────────────

_groq_client = None
def _groq():
    global _groq_client
    if _groq_client is None:
        from groq import Groq
        _groq_client = Groq(api_key=_env("GROQ_API_KEY"))
    return _groq_client


def _call_groq(images_b64: list[str], prompt: str) -> list[dict]:
    content = [
        {"type": "image_url",
         "image_url": {"url": f"data:image/png;base64,{img}"}}
        for img in images_b64
    ]
    content.append({"type": "text", "text": prompt})

    resp = _groq().chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role": "user", "content": content}],
        max_tokens=4096,
        temperature=0,
    )
    return _clean_json(resp.choices[0].message.content)


# ── Gemini vision ─────────────────────────────────────────────────────────────

def _call_gemini(images_b64: list[str], prompt: str) -> list[dict]:
    from google.genai import types

    parts = []
    for img in images_b64:
        parts.append(types.Part.from_bytes(
            data=base64.b64decode(img),
            mime_type="image/png",
        ))
    parts.append(types.Part.from_text(text=prompt))

    resp = _gemini().models.generate_content(
        model="gemini-2.0-flash",
        contents=parts,
    )
    return _clean_json(resp.text)


# ── Prompt templates ──────────────────────────────────────────────────────────

_PROMPTS = {
    "INV": """Extract every line item from this commercial invoice as a JSON array.
Each object must have exactly these keys:
  ma_hang       — product / item code (string)
  ma_hs         — HS / customs tariff code (string, "" if absent)
  ten_hang      — full product description (string)
  xuat_xu       — country of origin 2-letter ISO (string, "" if absent)
  so_luong_1    — quantity as a plain number string (e.g. "100", "2.5")
  don_vi_tinh_1 — unit of measure (kg, pcs, set, ltr, …)
  so_luong_2    — "" (leave empty)
  don_vi_tinh_2 — "" (leave empty)
  don_gia       — unit price as plain number string (no currency symbol)
  tong_tri_gia  — line total as plain number string (no currency symbol)

Do NOT include summary/total rows.
Return ONLY valid JSON — no markdown, no explanation.""",

    "PKL": """Extract every individual package/packing row from this packing list as a JSON array.
Each object must have exactly these keys:
  ma_hang       — product / item code (string)
  ma_hs         — HS code (string, "" if absent)
  ten_hang      — product description / designation (string)
  xuat_xu       — country of origin (string, "" if absent)
  so_luong_1    — net weight as plain number string (e.g. "25", "25.000")
  don_vi_tinh_1 — weight unit (kg, g, …) or quantity unit (pcs, set, …)
  so_luong_2    — "" (leave empty)
  don_vi_tinh_2 — "" (leave empty)
  don_gia       — "" (packing lists have no unit price)
  tong_tri_gia  — gross weight as plain number string, or "" if absent

Include one row per physical package. Do NOT include TOTAL rows.
Return ONLY valid JSON — no markdown, no explanation.""",

    "AN": """Extract every charge line from this arrival notice as a JSON array.
Each object must have exactly these keys:
  ma_hang       — "" (always empty)
  ma_hs         — "" (always empty)
  ten_hang      — charge name exactly as printed (e.g. "CFS CHARGE", "THC FEE")
  xuat_xu       — "" (always empty)
  so_luong_1    — quantity as plain number string
  don_vi_tinh_1 — unit (CBM, SHIPMENT, BL, …)
  so_luong_2    — "" (always empty)
  don_vi_tinh_2 — "" (always empty)
  don_gia       — rate BEFORE VAT as plain number string
  tong_tri_gia  — total AFTER VAT as plain number string

Do NOT include the grand total row.
Return ONLY valid JSON — no markdown, no explanation.""",
}


# ── Text-based AN parser (no API call needed) ─────────────────────────────────

_RE_AN_CHARGE = re.compile(
    r"^([A-Z][A-Z &/]+?)\s+"
    r"(\d[\d.]*)\s*"
    r"([A-Z]+)\s+"
    r"([\d.]+)\s+"
    r"\d+%VAT\s+"
    r"([\d.]+)",
    re.MULTILINE,
)

def _parse_an_text(pdf_path: str) -> list[dict]:
    text = _full_text(pdf_path)
    items = []
    for m in _RE_AN_CHARGE.finditer(text):
        name, qty, unit, rate, total = m.groups()
        items.append({
            "ma_hang": "", "ma_hs": "",
            "ten_hang": name.strip(), "xuat_xu": "",
            "so_luong_1": qty, "don_vi_tinh_1": unit,
            "so_luong_2": "", "don_vi_tinh_2": "",
            "don_gia": rate, "tong_tri_gia": total,
        })
    return items


# ── Public API ────────────────────────────────────────────────────────────────

def active_backend() -> str:
    """Return the name of the AI backend that will be used ('claude' or 'gemini')."""
    try:
        return _active_backend()
    except RuntimeError:
        return "none"


def extract(pdf_path: str, doc_type: str) -> list[dict]:
    """Return list of row-dicts for the given PDF and document type (INV/PKL/AN)."""
    doc_type = doc_type.upper()

    # AN: try text parser first (free, no API)
    if doc_type == "AN" and _has_text(pdf_path):
        items = _parse_an_text(pdf_path)
        if items:
            return items

    images  = _pages_as_base64_png(pdf_path)
    prompt  = _PROMPTS.get(doc_type, _PROMPTS["INV"])
    backend = _active_backend()

    if backend == "claude":
        return _call_claude(images, prompt)
    if backend == "groq":
        return _call_groq(images, prompt)
    return _call_gemini(images, prompt)
