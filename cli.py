"""cli.py — command-line PDF → Excel converter (no web server needed).

Usage:
  python cli.py <pdf_file> <INV|PKL> [--output path/to/output.xlsx]

Examples:
  python cli.py invoice.pdf INV
  python cli.py "Input/PKL/PKL_1.pdf" PKL --output "Output/PKL/result.xlsx"
"""
import argparse
import sys
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

import extractor
import excel_writer
import supabase_client as supa


def main():
    parser = argparse.ArgumentParser(
        description="DP World — PDF to Excel CLI converter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("pdf",  help="Path to the input PDF file")
    parser.add_argument("type", choices=["INV", "PKL"], help="Document type")
    parser.add_argument("-o", "--output", help="Output .xlsx path (default: same folder as PDF)")
    parser.add_argument("--no-upload", action="store_true",
                        help="Skip Supabase upload even if configured")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)
    if pdf_path.suffix.lower() != ".pdf":
        print(f"[ERROR] Expected a .pdf file, got: {pdf_path.suffix}")
        sys.exit(1)

    out_path = Path(args.output) if args.output else pdf_path.with_suffix(".xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # ── Check API key ────────────────────────────────────────────────────────
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("[WARN] ANTHROPIC_API_KEY not set — vision extraction will fail.")

    # ── Extract ──────────────────────────────────────────────────────────────
    print(f"[INFO] Extracting  {pdf_path.name}  ({args.type}) …")
    try:
        items = extractor.extract(str(pdf_path), args.type)
    except Exception as e:
        print(f"[ERROR] Extraction failed: {e}")
        sys.exit(1)

    if not items:
        print("[WARN] No items extracted — check the PDF.")
        sys.exit(1)

    # ── Write Excel ──────────────────────────────────────────────────────────
    print(f"[INFO] Writing {len(items)} rows → {out_path}")
    excel_writer.write_excel(items, out_path)
    print(f"[OK]   Saved locally: {out_path}")

    # ── Upload to Supabase ───────────────────────────────────────────────────
    if not args.no_upload and supa.is_configured():
        storage_path = f"{args.type}/{out_path.name}"
        try:
            conv_id = supa.insert_conversion(pdf_path.name, args.type)
            supa.upload_file(supa.XLSX_BUCKET, storage_path, out_path,
                             content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            supa.update_conversion(conv_id,
                status="done",
                message=f"{len(items)} rows",
                rows_extracted=len(items),
                xlsx_storage_path=storage_path)
            url = supa.signed_url(supa.XLSX_BUCKET, storage_path)
            print(f"[OK]   Uploaded to Supabase: {storage_path}")
            print(f"[OK]   Signed URL (1h):  {url}")
        except Exception as e:
            print(f"[WARN] Supabase upload failed: {e}")
    elif args.no_upload:
        print("[INFO] Supabase upload skipped (--no-upload).")
    else:
        print("[INFO] Supabase not configured — local file only.")


if __name__ == "__main__":
    main()
