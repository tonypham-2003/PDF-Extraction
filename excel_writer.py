"""excel_writer.py — write extracted items to .xlsx following template.xls layout.

Template structure (matches template.xls exactly):
  Row 1 : column headers  (blue background, white bold text)
  Row 2+: data rows
  Columns: STT | Mã hàng | Mã HS | Tên hàng | Xuất xứ |
           Số lượng 1 | Đơn vị tính 1 | Số lượng 2 | Đơn vị tính 2 |
           Đơn giá | Tổng trị giá

Đơn giá    → stored as string (preserves formatting, e.g. "13.50", "2,349.00")
Tổng trị giá → stored as float (numeric, matches template.xls ctype=2)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HEADER_COLOR = "1F4E79"
COL_WIDTHS   = [6, 20, 14, 60, 10, 13, 16, 13, 16, 18, 18]
HEADERS = [
    "STT", "Mã hàng ", "Mã HS", "Tên hàng", "Xuất xứ",
    "Số lượng 1", "Đơn vị tính 1", "Số lượng 2", "Đơn vị tính 2",
    "Đơn giá", "Tổng trị giá",
]

_BORDER = None
def _thin():
    global _BORDER
    if _BORDER is None:
        s = Side(style="thin")
        _BORDER = Border(left=s, right=s, top=s, bottom=s)
    return _BORDER


def _to_number(val):
    """Convert a value to float if possible, else return as-is."""
    if val is None or val == "":
        return ""
    try:
        return float(str(val).replace(",", ""))
    except (ValueError, TypeError):
        return val


def write_excel(items: list[dict], out_path) -> None:
    """Write items to xlsx following template.xls structure."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # ── Row 1: headers (matches template.xls row 0) ───────────────────────────
    hfill  = PatternFill("solid", fgColor=HEADER_COLOR)
    hfont  = Font(bold=True, color="FFFFFF", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = _thin()

    for ci, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = hfont; cell.fill = hfill
        cell.alignment = center; cell.border = border
    ws.row_dimensions[1].height = 30

    # ── Data rows from row 2 ──────────────────────────────────────────────────
    left  = Alignment(horizontal="left",  vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    for stt, item in enumerate(items, 1):
        ri = stt + 1
        vals = [
            stt,
            item.get("ma_hang",       ""),
            item.get("ma_hs",         ""),
            item.get("ten_hang",      ""),
            item.get("xuat_xu",       ""),
            item.get("so_luong_1",    ""),
            item.get("don_vi_tinh_1", ""),
            item.get("so_luong_2",    ""),
            item.get("don_vi_tinh_2", ""),
            item.get("don_gia",       ""),          # string, keep formatting
            _to_number(item.get("tong_tri_gia","")),# numeric, matches template
        ]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font      = Font(size=10)
            cell.border    = border
            cell.alignment = (left  if ci == 4 else
                              right if ci in (10, 11) else center)
        ws.row_dimensions[ri].height = 35

    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)
