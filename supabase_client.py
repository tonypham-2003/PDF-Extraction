"""supabase_client.py — Supabase REST + Storage via raw HTTP (no supabase-py).

Works with both old JWT keys (eyJ...) and new sb_secret_ / sb_publishable_ keys.
"""
import os
import json
import urllib.request
import urllib.error
from pathlib import Path

PDF_BUCKET  = "pdf-uploads"
XLSX_BUCKET = "excel-outputs"


# ── Config ────────────────────────────────────────────────────────────────────

def _url() -> str:
    return os.environ.get("SUPABASE_URL", "").rstrip("/")

def _key() -> str:
    return os.environ.get("SUPABASE_SERVICE_KEY", "")

def is_configured() -> bool:
    return bool(_url() and _key())

def _headers(extra: dict | None = None) -> dict:
    h = {
        "apikey":        _key(),
        "Authorization": f"Bearer {_key()}",
        "Content-Type":  "application/json",
    }
    if extra:
        h.update(extra)
    return h


# ── HTTP helpers ──────────────────────────────────────────────────────────────

def _get(path: str, params: dict | None = None) -> list | dict:
    qs = ""
    if params:
        qs = "?" + "&".join(f"{k}={v}" for k, v in params.items())
    req = urllib.request.Request(_url() + path + qs, headers=_headers())
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def _post(path: str, body: dict, extra_headers: dict | None = None) -> list | dict:
    data = json.dumps(body).encode()
    req = urllib.request.Request(
        _url() + path,
        data=data,
        headers=_headers(extra_headers),
        method="POST",
    )
    with urllib.request.urlopen(req) as r:
        text = r.read()
        return json.loads(text) if text.strip() else {}


def _patch(path: str, body: dict, params: dict | None = None) -> None:
    qs = ""
    if params:
        qs = "?" + "&".join(f"{k}={v}" for k, v in params.items())
    data = json.dumps(body).encode()
    req = urllib.request.Request(
        _url() + path + qs,
        data=data,
        headers=_headers({"Prefer": "return=minimal"}),
        method="PATCH",
    )
    with urllib.request.urlopen(req):
        pass


# ── Database ──────────────────────────────────────────────────────────────────

def insert_conversion(filename: str, doc_type: str,
                      pdf_storage_path: str | None = None) -> str:
    """Insert a row in public.conversions. Returns the new UUID."""
    res = _post("/rest/v1/conversions", {
        "filename":         filename,
        "doc_type":         doc_type,
        "status":           "pending",
        "pdf_storage_path": pdf_storage_path,
    }, extra_headers={"Prefer": "return=representation"})
    rows = res if isinstance(res, list) else [res]
    return rows[0]["id"]


def update_conversion(conv_id: str, **fields) -> None:
    """Update any columns on a conversions row by UUID."""
    _patch(f"/rest/v1/conversions", fields, params={"id": f"eq.{conv_id}"})


def list_conversions(limit: int = 50) -> list[dict]:
    """Return the most recent conversions (newest first)."""
    return _get("/rest/v1/conversions", {
        "select":   "*",
        "order":    "created_at.desc",
        "limit":    str(limit),
    })


# ── Storage ───────────────────────────────────────────────────────────────────

def upload_file(bucket: str, storage_path: str, local_path: str | Path,
                content_type: str = "application/octet-stream") -> str:
    """Upload a local file to a Supabase Storage bucket. Returns storage_path."""
    with open(local_path, "rb") as fh:
        data = fh.read()

    # Supabase Storage upload URL
    upload_url = f"{_url()}/storage/v1/object/{bucket}/{storage_path}"
    req = urllib.request.Request(
        upload_url,
        data=data,
        headers={
            "apikey":        _key(),
            "Authorization": f"Bearer {_key()}",
            "Content-Type":  content_type,
            "x-upsert":      "true",
        },
        method="POST",
    )
    with urllib.request.urlopen(req):
        pass
    return storage_path


def download_file(bucket: str, storage_path: str) -> bytes:
    """Download a file from Supabase Storage and return its raw bytes.

    Uses the service key — no expiry, always accessible as long as the
    file exists in the bucket.
    """
    req = urllib.request.Request(
        f"{_url()}/storage/v1/object/{bucket}/{storage_path}",
        headers={
            "apikey":        _key(),
            "Authorization": f"Bearer {_key()}",
        },
    )
    with urllib.request.urlopen(req) as r:
        return r.read()
