from copy import copy
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import quote, unquote
import argparse
import html
import sys
import webbrowser
import re
import csv
import unicodedata
from openpyxl import Workbook

# Cho phép chạy script trong thư mục này dù virtualenv được tạo ở máy khác.
for site_packages in (Path(__file__).resolve().parent / ".venv" / "lib").glob(
    "python*/site-packages"
):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

# ===== THƯ MỤC INPUT/OUTPUT =====
input_dir = Path("Excel")
output_dir = Path("Đã gộp")
max_upload_bytes = 50 * 1024 * 1024

# ===== CỘT DỮ LIỆU =====
# A = No.
# B = Ngày
# C = Mã khách hàng
# D = Ca
# E = Số lượng
# F = Tên món ăn
# H = Khối lượng định mức
# J = Khối lượng đi chợ
col_ngay = 2
col_ma_kh = 3
col_ca = 4
col_so_luong = 5
col_ten_mon = 6
col_khoi_luong = 8
col_khoi_luong_di_cho = 10

source_start_row = 2
title_row = 1
header_row = 2
body_start_row = 3

title_font_size = 18
header_font_size = 14
body_font_size = 13

visible_column_widths = {
    "C": 10,
    "D": 5.5,
    "E": 6,
    "F": 15,
    "G": 18,
    "H": 7,
    "I": 8,
    "J": 9,
    "K": 5.5,
}

hidden_print_columns = ("A", "B", "L")

A4_OUTPUT_HEADERS = [
    "Khách hàng",
    "Ca",
    "Site",
    "Loại món ăn",
    "Món ăn",
    "Số lượng",
    "Định mức\ncam kết",
    "Nguyên liệu",
    "Định\nmức",
    "Đơn vị\ntính",
    "KL yêu cầu\nsản xuất",
    "Đơn vị tính\nmua hàng",
    "KL duyệt đi chợ",
]

A4_LAST_COL = len(A4_OUTPUT_HEADERS)
A4_HEADER_ROW = 2
A4_DATA_START_ROW = 3

APPROVAL_FORMAT_MODE = "duyet_dinh_muc"
APPROVAL_OUTPUT_HEADERS = [
    "ID",
    "Trạng thái",
    "Ngày duyệt",
    "Cần duyệt",
    "Mã món",
    "Nhóm món",
    "Tên món ăn",
    "Nhóm NVL",
    "Tên nguyên vật liệu",
    "STT",
    "ĐVT",
    "ĐVT chợ",
    "Định mức theo KH",
    "Version",
]
APPROVAL_LAST_COL = len(APPROVAL_OUTPUT_HEADERS)
APPROVAL_HEADER_ROW = 2
APPROVAL_DATA_START_ROW = 3


def ensure_visible_worksheet(workbook, preferred=None):
    """Make an output workbook valid even when its source sheets were hidden."""
    worksheets = workbook.worksheets
    if not worksheets:
        raise ValueError("File Excel không có trang tính dữ liệu để xử lý.")

    visible_sheets = [ws for ws in worksheets if ws.sheet_state == "visible"]
    if visible_sheets:
        active_sheet = preferred if preferred in visible_sheets else visible_sheets[0]
    else:
        active_sheet = preferred if preferred in worksheets else worksheets[0]
        active_sheet.sheet_state = "visible"

    workbook.active = active_sheet
    return active_sheet


def normalized_value(raw_value, previous_value=None):
    if raw_value is None:
        return previous_value
    if isinstance(raw_value, str) and raw_value.strip() == "":
        return previous_value
    return raw_value


def key_value(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return " ".join(value.split()).casefold()
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return value


def format_date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    if value is None:
        return ""
    return str(value)


def extract_date_from_filename(file_path, year=None):
    if year is None:
        year = datetime.now().year
    stem = file_path.stem
    m_full = re.search(r"(\d{4})[._-](\d{1,2})[._-](\d{1,2})", stem)
    if m_full:
        return f"{m_full.group(3).zfill(2)}/{m_full.group(2).zfill(2)}/{m_full.group(1)}"
    m_full_rev = re.search(r"(\d{1,2})[._-](\d{1,2})[._-](\d{4})", stem)
    if m_full_rev:
        return f"{m_full_rev.group(1).zfill(2)}/{m_full_rev.group(2).zfill(2)}/{m_full_rev.group(3)}"
    m_short = re.search(r"[_\s](\d{1,2})[._-](\d{1,2})(?:$|[_\s])", stem)
    if not m_short:
        m_short = re.search(r"(\d{1,2})[._-](\d{1,2})", stem)
    if m_short:
        day = int(m_short.group(1))
        month = int(m_short.group(2))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{str(day).zfill(2)}/{str(month).zfill(2)}/{year}"
    return None


def parse_date_value(val):
    if isinstance(val, (datetime, date)):
        return val
    if not val:
        return None
    val_str = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(val_str, fmt)
        except ValueError:
            continue
    return None

def collect_date_text(ws, start_row, max_row, col_ngay_idx=None):
    if col_ngay_idx is None:
        col_ngay_idx = col_ngay
    if col_ngay_idx is None:
        return ""
    dates = []
    seen = set()

    for row in range(start_row, max_row + 1):
        value = ws.cell(row, col_ngay_idx).value
        dt_val = parse_date_value(value)
        text = format_date(dt_val) if dt_val else str(value).strip()
        if text and text not in seen:
            seen.add(text)
            dates.append((dt_val or value, text))

    if not dates:
        return ""
    if len(dates) == 1:
        return dates[0][1]

    sortable_dates = [item for item in dates if isinstance(item[0], (datetime, date))]
    if len(sortable_dates) == len(dates):
        sortable_dates.sort(key=lambda item: item[0])
        return f"{sortable_dates[0][1]} - {sortable_dates[-1][1]}"

    return ", ".join(text for _, text in dates[:3])


def kitchen_name_from_file(file_path):
    name = file_path.stem.replace("_", " ").replace("-", " ")
    name = " ".join(name.split())
    if name.casefold().startswith("bom "):
        name = name[4:].strip()
    if name:
        name = name[:1].upper() + name[1:]
    return name or file_path.stem


def unmerge_group_columns(ws, start_row, col_ma_kh_idx=None, col_ca_idx=None, col_so_luong_idx=None, col_ten_mon_idx=None):
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row >= start_row:
            try:
                ws.unmerge_cells(str(merged_range))
            except Exception:
                pass


def get_lixil_site_order(site_name):
    if not site_name:
        return 99
    s = str(site_name).lower().replace(" ", "").replace("-", "")
    if "fab1" in s:
        return 1
    if "fab2" in s:
        return 2
    if "vpc" in s:
        return 3
    return 99


def text_key(value):
    text = "" if value is None else str(value)
    text = text.replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text.casefold())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())


def get_co_cau_order(co_cau_val, customer_val=None):
    if not co_cau_val:
        return 99
    s = text_key(co_cau_val)
    order = [
        (("mon man", "suat man", "suat com", "com", "cai thien", "cai tihen", "an sang"), 1),
        (("mon nuoc", "suat nuoc", "nuoc", "chao"), 2),
        (("mon chay", "suat chay", "chay"), 3),
    ]
    for keywords, index in order:
        if any(keyword in s for keyword in keywords):
            return index
    return 99


def meal_type_label(co_cau_val):
    order = get_co_cau_order(co_cau_val)
    if order == 1:
        return "Mặn"
    if order == 2:
        return "Nước"
    if order == 3:
        return "Chay"
    return co_cau_val or ""


def is_generic_meal_label(value):
    key = text_key(value)
    return key in {
        "man",
        "nuoc",
        "chay",
        "mon man",
        "mon nuoc",
        "mon chay",
        "suat man",
        "suat nuoc",
        "suat chay",
    }


def sort_rows_by_groups(ws, max_row, max_col, col_ngay_idx, col_ca_idx, col_ma_kh_idx, col_site_an_idx, col_co_cau_idx, col_ten_mon_idx, col_so_luong_idx):
    prev_ngay = None
    prev_ca = None
    prev_company = None
    prev_site = None
    prev_co_cau = None
    prev_dish = None
    prev_quantity = None
    
    rows = []
    for row_idx in range(source_start_row, max_row + 1):
        values = [ws.cell(row_idx, col).value for col in range(1, max_col + 1)]
        
        if col_ngay_idx is not None and col_ngay_idx <= len(values):
            prev_ngay = normalized_value(values[col_ngay_idx - 1], prev_ngay)
            values[col_ngay_idx - 1] = prev_ngay
            
        if col_ca_idx is not None and col_ca_idx <= len(values):
            prev_ca = normalized_value(values[col_ca_idx - 1], prev_ca)
            values[col_ca_idx - 1] = prev_ca
            
        if col_ma_kh_idx is not None and col_ma_kh_idx <= len(values):
            prev_company = normalized_value(values[col_ma_kh_idx - 1], prev_company)
            values[col_ma_kh_idx - 1] = prev_company
            
        if col_site_an_idx is not None and col_site_an_idx <= len(values):
            prev_site = normalized_value(values[col_site_an_idx - 1], prev_site)
            values[col_site_an_idx - 1] = prev_site
            
        if col_co_cau_idx is not None and col_co_cau_idx <= len(values):
            prev_co_cau = normalized_value(values[col_co_cau_idx - 1], prev_co_cau)
            values[col_co_cau_idx - 1] = prev_co_cau
            
        if col_ten_mon_idx is not None and col_ten_mon_idx <= len(values):
            prev_dish = normalized_value(values[col_ten_mon_idx - 1], prev_dish)
            values[col_ten_mon_idx - 1] = prev_dish
            
        if col_so_luong_idx is not None and col_so_luong_idx <= len(values):
            prev_quantity = normalized_value(values[col_so_luong_idx - 1], prev_quantity)
            values[col_so_luong_idx - 1] = prev_quantity

        rows.append((row_idx, values))

    def get_row_sort_key(item):
        row_idx, values = item
        
        ngay_val = values[col_ngay_idx - 1] if col_ngay_idx else ""
        dt_val = parse_date_value(ngay_val)
        if dt_val:
            date_key = (0, dt_val)
        else:
            date_key = (1, str(ngay_val))
        
        ca_val = values[col_ca_idx - 1] if col_ca_idx else ""
        ca_key = get_shift_order(normalize_shift(ca_val))
        
        kh_val = str(values[col_ma_kh_idx - 1] or "").strip().upper() if col_ma_kh_idx else ""
        
        site_val = str(values[col_site_an_idx - 1] or "").strip() if col_site_an_idx else ""
        if kh_val == "LIXIL":
            site_key = (get_lixil_site_order(site_val), site_val.casefold())
        else:
            site_key = (0, site_val.casefold())
            
        if col_co_cau_idx:
            co_cau_raw = str(values[col_co_cau_idx - 1] or "").strip()
            co_cau_key = (get_co_cau_order(co_cau_raw, kh_val), co_cau_raw.lower())
        else:
            co_cau_key = (99, "")
            
        dish_val = str(values[col_ten_mon_idx - 1] or "").strip().lower() if col_ten_mon_idx else ""
        
        return (date_key, ca_key, kh_val, site_key, co_cau_key, dish_val, row_idx)

    rows.sort(key=get_row_sort_key)

    for new_row, (orig_row_idx, values) in enumerate(rows, start=source_start_row):
        for col, value in enumerate(values, start=1):
            ws.cell(new_row, col).value = value
        ws.cell(new_row, 1).value = new_row - source_start_row + 1


def set_center(ws, row, col):
    ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")


def merge_group_columns(ws, max_row, column_indices, col_ngay_idx, col_ca_idx, col_ma_kh_idx, col_site_an_idx, col_ten_mon_idx, col_so_luong_idx, page_break_rows=None):
    valid_cols = [c for c in column_indices if c is not None]
    if not valid_cols:
        return
    page_break_rows = set(page_break_rows or ())

    row_keys = {}
    for row in range(body_start_row, max_row + 1):
        key = []
        for c in valid_cols:
            val = ws.cell(row, c).value
            key.append(val)
        row_keys[row] = key

    for col_idx_in_hierarchy, col in enumerate(valid_cols):
        start_row = body_start_row
        current_prefix = [row_keys[body_start_row][i] for i in range(col_idx_in_hierarchy + 1)]
        
        for row in range(body_start_row + 1, max_row + 2):
            prefix = [row_keys[row][i] for i in range(col_idx_in_hierarchy + 1)] if row <= max_row else None
            
            if row in page_break_rows or prefix != current_prefix:
                end_row = row - 1
                if start_row < end_row:
                    ws.merge_cells(
                        start_row=start_row,
                        start_column=col,
                        end_row=end_row,
                        end_column=col,
                    )
                # Apply alignment to merged range start cell
                if col in (col_ngay_idx, col_ca_idx, col_ma_kh_idx, col_site_an_idx, col_so_luong_idx):
                    ws.cell(start_row, col).alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                else:
                    ws.cell(start_row, col).alignment = Alignment(
                        horizontal="left", vertical="center", wrap_text=True
                    )
                current_prefix = prefix
                start_row = row


def add_group_page_breaks(ws, max_row, column_indices, max_body_rows=22):
    """Keep printed groups readable and prevent merged cells crossing pages."""
    valid_cols = [c for c in column_indices if c is not None]
    ws.row_breaks.brk = []
    if not valid_cols or max_row < body_start_row + max_body_rows:
        return set()

    def key_at(row, depth):
        return tuple(ws.cell(row, col).value for col in valid_cols[:depth])

    # Prefer a new page at a Site boundary; fall back to a dish boundary when
    # one site contains more rows than a page can comfortably show.
    major_depth = min(4, len(valid_cols))
    major_boundaries = {
        row for row in range(body_start_row + 1, max_row + 1)
        if key_at(row, major_depth) != key_at(row - 1, major_depth)
    }
    detail_boundaries = {
        row for row in range(body_start_row + 1, max_row + 1)
        if key_at(row, len(valid_cols)) != key_at(row - 1, len(valid_cols))
    }

    break_rows = set()
    page_start = body_start_row
    while page_start + max_body_rows <= max_row:
        latest_start = page_start + max_body_rows
        earliest_start = page_start + max(10, max_body_rows // 2)
        preferred = [
            row for row in major_boundaries
            if earliest_start <= row <= latest_start
        ]
        fallback = [
            row for row in detail_boundaries
            if earliest_start <= row <= latest_start
        ]
        next_page = max(preferred or fallback or [latest_start])
        if next_page <= page_start:
            next_page = latest_start
        break_rows.add(next_page)
        ws.row_breaks.append(Break(id=next_page - 1))
        page_start = next_page

    return break_rows


def visible_last_column(ws, max_col):
    for col in range(max_col, 0, -1):
        value = ws.cell(header_row, col).value
        if isinstance(value, str) and value.strip().casefold() == "sourceid":
            continue
        return col
    return max_col


def apply_print_layout(ws, max_row, max_col, title_text, col_ma_kh_idx=None, col_ca_idx=None, col_so_luong_idx=None, col_ten_mon_idx=None, col_khoi_luong_idx=None, col_khoi_luong_di_cho_idx=None, col_ngay_idx=None, col_site_an_idx=None, col_co_cau_idx=None):
    if col_ma_kh_idx is None: col_ma_kh_idx = col_ma_kh
    if col_ca_idx is None: col_ca_idx = col_ca
    if col_so_luong_idx is None: col_so_luong_idx = col_so_luong
    if col_ten_mon_idx is None: col_ten_mon_idx = col_ten_mon
    if col_khoi_luong_idx is None: col_khoi_luong_idx = col_khoi_luong
    if col_khoi_luong_di_cho_idx is None: col_khoi_luong_di_cho_idx = col_khoi_luong_di_cho

    last_visible_col = visible_last_column(ws, max_col)

    # Keep the requested grouping columns visible. Only the technical ID column
    # and an optional SourceId column are excluded from the printed report.
    hidden_cols_indices = [1]
    col_12_val = ws.cell(header_row, 12).value
    if col_12_val and str(col_12_val).strip().lower() == "sourceid":
        hidden_cols_indices.append(12)

    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = c in hidden_cols_indices

    first_visible_col = min(c for c in range(1, max_col + 1) if c not in hidden_cols_indices)
    first_visible_letter = get_column_letter(first_visible_col)
    last_visible_letter = get_column_letter(last_visible_col)

    ws.merge_cells(
        start_row=title_row,
        start_column=first_visible_col,
        end_row=title_row,
        end_column=last_visible_col,
    )
    title_cell = ws.cell(title_row, first_visible_col)
    title_cell.value = title_text
    title_cell.font = Font(name="Calibri", bold=True, size=title_font_size, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="404040")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[title_row].height = 28

    # Set column widths dynamically based on role mapping (scaled for size 13 font)
    role_widths = {
        col_ngay_idx: 14,
        col_ma_kh_idx: 18,
        col_ca_idx: 8,
        col_so_luong_idx: 10,
        col_ten_mon_idx: 24,
        col_khoi_luong_idx: 10,
        col_khoi_luong_di_cho_idx: 14,
    }
    if col_site_an_idx:
        role_widths[col_site_an_idx] = 14
    if col_co_cau_idx:
        role_widths[col_co_cau_idx] = 14
    
    col_nvl_idx = None
    col_dvt_idx = None
    col_dvt_dc_idx = None
    
    for c in range(1, max_col + 1):
        val = ws.cell(header_row, c).value
        if not val:
            continue
        v_str = str(val).lower()
        if "nguyên vật liệu" in v_str or "nguyen vat lieu" in v_str or "nvl" in v_str:
            col_nvl_idx = c
        elif "đơn vị tính" in v_str or "don vi tinh" in v_str or "dvt" in v_str:
            if "đi chợ" in v_str or "di cho" in v_str or "mua hàng" in v_str or "mua hang" in v_str:
                col_dvt_dc_idx = c
            else:
                col_dvt_idx = c
                
    if col_nvl_idx:
        role_widths[col_nvl_idx] = 30
    if col_dvt_idx:
        role_widths[col_dvt_idx] = 12
    if col_dvt_dc_idx:
        role_widths[col_dvt_dc_idx] = 12

    for c, width in role_widths.items():
        if c is not None and c <= max_col:
            ws.column_dimensions[get_column_letter(c)].width = width

    thin = Side(style="thin", color="D9D9D9")
    medium = Side(style="medium", color="B7B7B7")
    header_fill = PatternFill("solid", fgColor="C0C0C0")
    group_fill = PatternFill("solid", fgColor="F2F2F2")

    for col in range(1, last_visible_col + 1):
        cell = ws.cell(header_row, col)
        cell.font = Font(name="Calibri", bold=True, size=header_font_size, color="000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=medium, bottom=medium)
    ws.row_dimensions[header_row].height = 64

    previous_group_key = None
    for row in range(body_start_row, max_row + 1):
        group_key = (
            ws.cell(row, col_ngay_idx).value if col_ngay_idx else None,
            ws.cell(row, col_ca_idx).value if col_ca_idx else None,
            ws.cell(row, col_ma_kh_idx).value,
            ws.cell(row, col_site_an_idx).value if col_site_an_idx else None,
            ws.cell(row, col_co_cau_idx).value if col_co_cau_idx else None,
            ws.cell(row, col_ten_mon_idx).value,
        )
        is_group_start = group_key != previous_group_key
        previous_group_key = group_key

        for col in range(1, last_visible_col + 1):
            cell = ws.cell(row, col)
            updated_font = copy(cell.font)
            updated_font.name = "Calibri"
            updated_font.size = body_font_size
            cell.font = updated_font
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(
                left=thin,
                right=thin,
                top=medium if is_group_start else thin,
                bottom=thin,
            )
            if is_group_start and col in (
                col_ngay_idx,
                col_ma_kh_idx,
                col_ca_idx,
                col_site_an_idx,
                col_co_cau_idx,
                col_ten_mon_idx,
            ):
                cell.fill = group_fill

        ws.row_dimensions[row].height = 24

    for row in range(body_start_row, max_row + 1):
        for col in (col_ngay_idx, col_ma_kh_idx, col_ca_idx, col_site_an_idx, col_co_cau_idx):
            if col is not None and col <= max_col:
                ws.cell(row, col).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
        ws.cell(row, col_ten_mon_idx).alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True
        )

    for row in range(body_start_row, max_row + 1):
        ws.cell(row, col_so_luong_idx).number_format = "#,##0.###"
        if col_khoi_luong_idx is not None and col_khoi_luong_idx <= max_col:
            ws.cell(row, col_khoi_luong_idx).number_format = "#,##0.###"
        if col_khoi_luong_di_cho_idx is not None and col_khoi_luong_di_cho_idx <= max_col:
            ws.cell(row, col_khoi_luong_di_cho_idx).number_format = "#,##0.###"
        ws.cell(row, col_so_luong_idx).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        for col in (col_khoi_luong_idx, col_khoi_luong_di_cho_idx):
            if col is not None and col <= max_col:
                ws.cell(row, col).alignment = Alignment(
                    horizontal="right", vertical="top", wrap_text=True
                )

    ws.freeze_panes = f"{get_column_letter(first_visible_col)}3"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"{first_visible_letter}{header_row}:{last_visible_letter}{max_row}"
    ws.print_area = f"{first_visible_letter}{title_row}:{last_visible_letter}{max_row}"
    ws.print_title_rows = "1:2"

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.blackAndWhite = True
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.pageOrder = "downThenOver"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15
    ws.oddHeader.center.text = ""
    ws.oddFooter.center.text = ""
    ws.sheet_view.zoomScale = 85


def process_sheet(ws, file_path, date_mode="auto", selected_kitchen=None):
    max_row = ws.max_row
    if max_row < source_start_row:
        return
    max_col = ws.max_column

    # Build col_map dynamically from row 1 (header row)
    col_map = {
        "ngay": col_ngay,
        "ma_kh": col_ma_kh,
        "ca": col_ca,
        "so_luong": col_so_luong,
        "ten_mon": col_ten_mon,
        "khoi_luong": col_khoi_luong,
        "khoi_luong_di_cho": col_khoi_luong_di_cho,
        "site_an": None,
        "co_cau": None
    }
    
    header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    has_any_header = any(h is not None for h in header)
    if has_any_header:
        for idx, col in enumerate(header, start=1):
            if col is None:
                continue
            c_lower = str(col).lower().strip()
            if any(x in c_lower for x in ["ngày", "ngay", "date"]):
                col_map["ngay"] = idx
            elif any(x in c_lower for x in ["khách hàng", "khach hang", "mã kh", "ma kh", "khach_hang", "mã khách hàng"]) and "site" not in c_lower:
                col_map["ma_kh"] = idx
            elif any(x in c_lower for x in ["ca", "shift"]):
                col_map["ca"] = idx
            elif any(x in c_lower for x in ["site ăn", "site an", "site"]):
                col_map["site_an"] = idx
            elif any(x in c_lower for x in ["cơ cấu món ăn", "co cau mon an", "cơ cấu", "co cau", "nhóm món", "nhom mon"]):
                col_map["co_cau"] = idx
            elif any(x in c_lower for x in ["món ăn", "mon an", "tên món", "ten mon"]) and "cơ cấu" not in c_lower and "co cau" not in c_lower:
                col_map["ten_mon"] = idx
            elif any(x in c_lower for x in ["số lượng", "so luong", "s.lượng", "qty", "quantity"]):
                col_map["so_luong"] = idx
            elif any(x in c_lower for x in ["định mức", "dinh muc", "khối lượng", "định lượng"]) and "đi chợ" not in c_lower and "di cho" not in c_lower and "yêu cầu" not in c_lower and "yeu cau" not in c_lower and "mua hàng" not in c_lower and "mua hang" not in c_lower:
                col_map["khoi_luong"] = idx
            elif any(x in c_lower for x in ["khối lượng đi chợ", "đi chợ", "di cho", "định lương đi chợ", "yêu cầu sản xuất", "yeu cau san xuat", "kl yêu cầu", "kl yeu cau", "mua hàng", "mua hang"]) and "đơn vị tính" not in c_lower and "don vi tinh" not in c_lower and "dvt" not in c_lower and "đvt" not in c_lower:
                col_map["khoi_luong_di_cho"] = idx

    col_ngay_idx = col_map["ngay"]
    col_ma_kh_idx = col_map["ma_kh"]
    col_ca_idx = col_map["ca"]
    col_so_luong_idx = col_map["so_luong"]
    col_ten_mon_idx = col_map["ten_mon"]
    col_khoi_luong_idx = col_map["khoi_luong"]
    col_khoi_luong_di_cho_idx = col_map["khoi_luong_di_cho"]
    col_site_an_idx = col_map["site_an"]
    col_co_cau_idx = col_map["co_cau"]

    # CSV imports arrive as text. Store dates and quantities as real Excel
    # values so printouts are clean and users can continue sorting/calculating.
    for row in range(source_start_row, max_row + 1):
        if col_ngay_idx:
            parsed_date = parse_date_value(ws.cell(row, col_ngay_idx).value)
            if parsed_date:
                ws.cell(row, col_ngay_idx).value = parsed_date
                ws.cell(row, col_ngay_idx).number_format = "dd/mm/yyyy"
        for numeric_col in (
            col_so_luong_idx,
            col_khoi_luong_idx,
            col_khoi_luong_di_cho_idx,
        ):
            if numeric_col:
                ws.cell(row, numeric_col).value = to_numeric(
                    ws.cell(row, numeric_col).value
                )

    if date_mode == "none":
        date_text = ""
    else:
        date_text = collect_date_text(ws, source_start_row, max_row, col_ngay_idx)
        
    if selected_kitchen:
        kitchen_name = selected_kitchen
    else:
        kitchen_name = kitchen_name_from_file(file_path)
        
    title_text = f"Bếp: {kitchen_name}"
    if date_text:
        title_text = f"{title_text} | Ngày: {date_text}"

    unmerge_group_columns(ws, source_start_row)
    sort_rows_by_groups(
        ws, max_row, max_col, 
        col_ngay_idx, col_ca_idx, col_ma_kh_idx, col_site_an_idx, col_co_cau_idx, col_ten_mon_idx, col_so_luong_idx
    )

    ws.insert_rows(title_row)
    max_row = ws.max_row

    column_indices = [col_ngay_idx, col_ca_idx, col_ma_kh_idx, col_site_an_idx, col_co_cau_idx, col_ten_mon_idx]
    apply_print_layout(
        ws, max_row, max_col, title_text, 
        col_ma_kh_idx, col_ca_idx, col_so_luong_idx, col_ten_mon_idx, col_khoi_luong_idx, col_khoi_luong_di_cho_idx, col_ngay_idx,
        col_site_an_idx, col_co_cau_idx
    )
    page_break_rows = add_group_page_breaks(ws, max_row, column_indices)
    merge_group_columns(
        ws, max_row, column_indices, col_ngay_idx, col_ca_idx,
        col_ma_kh_idx, col_site_an_idx, col_ten_mon_idx, col_so_luong_idx,
        page_break_rows=page_break_rows,
    )


def save_workbook(wb, output_file):
    ensure_visible_worksheet(wb)
    try:
        wb.save(output_file)
        return output_file
    except PermissionError:
        for index in range(1, 100):
            suffix = "mới" if index == 1 else f"mới {index}"
            fallback_file = output_file.with_name(
                f"{output_file.stem} - {suffix}{output_file.suffix}"
            )
            if not fallback_file.exists():
                wb.save(fallback_file)
                return fallback_file
        raise


def safe_filename(filename):
    filename = (filename or "file.xlsx").replace("\\", "/").split("/")[-1].strip()
    filename = filename.replace("\r", "").replace("\n", "")
    return filename or "file.xlsx"


def output_filename(filename):
    input_name = safe_filename(filename)
    suffix = Path(input_name).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        suffix = ".xlsx"
    return f"{Path(input_name).stem} - da dinh dang{suffix}"


def output_filename_for_format(filename, format_mode):
    input_name = safe_filename(filename)
    stem = Path(input_name).stem
    if format_mode == "format2":
        return f"{stem} - giay di cho.xlsx"
    if format_mode == APPROVAL_FORMAT_MODE:
        return f"{stem} - duyet dinh muc.xlsx"
    return output_filename(filename)


def content_disposition_filename(filename):
    filename = safe_filename(filename)
    ascii_fallback = filename.encode("ascii", "ignore").decode("ascii") or "bom.xlsx"
    ascii_fallback = ascii_fallback.replace('"', "")
    return f'attachment; filename="{ascii_fallback}"; filename*=UTF-8\'\'{quote(filename)}'


def clean_header(value):
    return text_key(value).replace("\n", " ")


def choose_a4_dish_column(ws, header_row, candidates):
    lookup = merged_value_lookup(ws)
    best = None
    max_sample_row = min(ws.max_row, header_row + 80)

    for col in candidates:
        header_text = clean_header(ws.cell(header_row, col).value)
        nonempty = 0
        generic = 0
        distinct = set()
        for row in range(header_row + 1, max_sample_row + 1):
            value = effective_cell_value(ws, lookup, row, col)
            value_key = text_key(value)
            if not value_key:
                continue
            nonempty += 1
            distinct.add(value_key)
            if is_generic_meal_label(value):
                generic += 1
        header_bonus = 8 if "ten mon" in header_text else 0
        score = nonempty + len(distinct) + header_bonus - generic * 4
        option = (score, col)
        if best is None or option > best:
            best = option

    return best[1] if best else None


def is_numeric_like(value):
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return bool(re.fullmatch(r"-?\d+(?:[,.]\d+)?", str(value).strip()))


def choose_a4_ingredient_column(ws, header_row, candidates):
    lookup = merged_value_lookup(ws)
    best = None
    max_sample_row = min(ws.max_row, header_row + 80)
    expanded = {}

    for source_col in candidates:
        for col in (source_col, source_col + 1):
            if 1 <= col <= ws.max_column:
                distance = abs(col - source_col)
                expanded[col] = min(distance, expanded.get(col, distance))

    for col, distance in sorted(expanded.items()):
        header_text = clean_header(ws.cell(header_row, col).value)
        nonempty = 0
        textish = 0
        numeric = 0
        short_values = 0
        distinct = set()

        for row in range(header_row + 1, max_sample_row + 1):
            value = effective_cell_value(ws, lookup, row, col)
            value_key = text_key(value)
            if not value_key:
                continue
            nonempty += 1
            distinct.add(value_key)
            if is_numeric_like(value):
                numeric += 1
            elif any(ch.isalpha() for ch in value_key):
                textish += 1
                if len(value_key) <= 2:
                    short_values += 1

        header_bonus = 0
        if "ten nguyen" in header_text or "ten nvl" in header_text:
            header_bonus += 24
        elif "nguyen lieu" in header_text or "nguyen vat lieu" in header_text or "nvl" in header_text:
            header_bonus += 8

        header_penalty = 0
        if any(pattern in header_text for pattern in ("ma nvl", "ma nguyen", "stt", "so thu tu", "no.", "id")):
            header_penalty += 30
        if any(pattern in header_text for pattern in ("dinh muc", "dinh luong", "don vi", "dvt", "so luong", "kl ")):
            header_penalty += 30

        score = (
            header_bonus
            + nonempty
            + textish * 5
            + len(distinct) * 2
            - numeric * 8
            - short_values * 2
            - header_penalty
            - distance
        )
        option = (score, textish, len(distinct), -distance, -col)
        if best is None or option > best[0]:
            best = (option, col)

    return best[1] if best else None


def find_a4_source_columns(ws):
    def matches(header_text, patterns, excluded=()):
        return any(pattern in header_text for pattern in patterns) and not any(
            pattern in header_text for pattern in excluded
        )

    required_keys = {
        "date",
        "customer",
        "shift",
        "site",
        "structure",
        "dish",
        "quantity",
        "ingredient",
        "norm",
        "unit",
        "required",
        "purchase_unit",
    }
    optional_keys = {"approved", "committed_norm"}

    for row in range(1, min(ws.max_row, 10) + 1):
        columns = {}
        dish_candidates = []
        ingredient_candidates = []
        for col in range(1, ws.max_column + 1):
            header_text = clean_header(ws.cell(row, col).value)
            if not header_text:
                continue
            if matches(header_text, ("ngay", "date")):
                columns.setdefault("date", col)
            elif matches(header_text, ("khach hang", "ma kh", "ma khach hang"), ("site",)):
                columns.setdefault("customer", col)
            elif matches(header_text, ("co cau", "nhom mon")):
                columns.setdefault("structure", col)
            elif header_text in {"ca", "shift"} or header_text.startswith("ca "):
                columns.setdefault("shift", col)
            elif matches(header_text, ("site an", "site")):
                columns.setdefault("site", col)
            elif matches(header_text, ("mon an", "ten mon"), ("co cau", "nhom mon", "loai mon")):
                dish_candidates.append(col)
            elif matches(header_text, ("so luong", "qty", "quantity")):
                columns.setdefault("quantity", col)
            elif matches(header_text, ("nguyen vat lieu", "nguyen lieu", "nvl")):
                ingredient_candidates.append(col)
            elif matches(
                header_text,
                ("dinh muc cam ket", "dm cam ket", "cam ket"),
                ("yeu cau", "di cho", "mua hang"),
            ):
                columns.setdefault("committed_norm", col)
            elif matches(
                header_text,
                ("dinh muc", "dinh luong"),
                ("cam ket", "yeu cau", "di cho", "mua hang"),
            ):
                columns.setdefault("norm", col)
            elif matches(
                header_text,
                ("don vi tinh", "dvt", "dvt"),
                ("mua hang", "di cho", "yeu cau"),
            ):
                columns.setdefault("unit", col)
            elif matches(header_text, ("yeu cau san xuat", "kl yeu cau", "khoi luong di cho")):
                columns.setdefault("required", col)
            elif matches(header_text, ("don vi tinh mua hang", "dvt mua hang", "don vi mua hang")):
                columns.setdefault("purchase_unit", col)
            elif matches(header_text, ("duyet di cho", "kl duyet")):
                columns.setdefault("approved", col)

        if dish_candidates:
            columns["dish"] = choose_a4_dish_column(ws, row, dish_candidates)
        if ingredient_candidates:
            columns["ingredient"] = choose_a4_ingredient_column(ws, row, ingredient_candidates)

        if required_keys.issubset(columns):
            for key in optional_keys:
                columns.setdefault(key, None)
            return row, columns

    missing = ", ".join(sorted(required_keys))
    raise ValueError(f"Không tìm thấy đủ cột nguồn để xuất Format 1 A4: {missing}.")


def merged_value_lookup(ws):
    lookup = {}
    for merged_range in ws.merged_cells.ranges:
        value = ws.cell(merged_range.min_row, merged_range.min_col).value
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                lookup[(row, col)] = value
    return lookup


def effective_cell_value(ws, lookup, row, col):
    if col is None:
        return None
    return lookup.get((row, col), ws.cell(row, col).value)


def parse_print_number(value):
    if not isinstance(value, str):
        return value
    text = value.strip()
    if not re.fullmatch(r"-?\d+(?:[,.]\d+)?", text):
        return value
    number = float(text.replace(",", "."))
    return int(number) if number.is_integer() else number


def a4_row_empty(values):
    return all(value is None or str(value).strip() == "" for value in values)


def collect_a4_records(ws):
    header_row, columns = find_a4_source_columns(ws)
    lookup = merged_value_lookup(ws)
    records = []
    previous = {key: None for key in ("date", "customer", "shift", "site", "structure", "dish", "quantity")}

    for source_row in range(header_row + 1, ws.max_row + 1):
        raw = {
            key: effective_cell_value(ws, lookup, source_row, col)
            for key, col in columns.items()
        }
        for key in previous:
            raw[key] = normalized_value(raw.get(key), previous[key])
            previous[key] = raw[key]

        values = [
            raw["date"],
            raw["customer"],
            raw["shift"],
            raw["site"],
            meal_type_label(raw.get("structure")),
            raw["dish"],
            parse_print_number(raw["quantity"]),
            parse_print_number(raw.get("committed_norm")),
            clean_ingredient_name(raw["ingredient"]),
            parse_print_number(raw["norm"]),
            raw["unit"],
            parse_print_number(raw["required"]),
            raw["purchase_unit"],
            parse_print_number(raw.get("approved")),
        ]
        if a4_row_empty(values):
            continue
        records.append(
            {
                "values": values,
                "meal_order": get_co_cau_order(raw.get("structure"), raw.get("customer")),
                "source_index": len(records),
            }
        )

    return sort_a4_records(records)


def sort_a4_records(records):
    groups = []
    start = 0
    while start < len(records):
        group_key = tuple(text_key(records[start]["values"][idx]) for idx in range(4))
        end = start + 1
        while end < len(records):
            next_key = tuple(text_key(records[end]["values"][idx]) for idx in range(4))
            if next_key != group_key:
                break
            end += 1
        group = records[start:end]
        group.sort(key=lambda record: (record["meal_order"], record["source_index"]))
        first_values = group[0]["values"]
        groups.append(
            {
                "date": first_values[0],
                "customer": first_values[1],
                "shift": first_values[2],
                "site": first_values[3],
                "source_index": group[0]["source_index"],
                "rows": group,
            }
        )
        start = end

    sorted_records = []
    start = 0
    while start < len(groups):
        outer_key = tuple(text_key(groups[start][key]) for key in ("date", "customer", "shift"))
        end = start + 1
        while end < len(groups):
            next_outer_key = tuple(text_key(groups[end][key]) for key in ("date", "customer", "shift"))
            if next_outer_key != outer_key:
                break
            end += 1

        outer_groups = groups[start:end]
        if "lixil" in text_key(outer_groups[0]["customer"]):
            outer_groups.sort(
                key=lambda group: (
                    get_lixil_site_order(group["site"]),
                    text_key(group["site"]),
                    group["source_index"],
                )
            )
        for group in outer_groups:
            sorted_records.extend(group["rows"])
        start = end
    return [record["values"] for record in sorted_records]


def a4_date_title(rows):
    dates = []
    seen = set()
    for row in rows:
        date_value = row[0] if row else None
        dt_val = parse_date_value(date_value)
        text = format_date(dt_val) if dt_val else str(date_value or "").strip()
        if text and text not in seen:
            seen.add(text)
            dates.append((dt_val or date_value, text))

    if not dates:
        return "Ngày:"
    if len(dates) == 1:
        return f"Ngày: {dates[0][1]}"

    sortable_dates = [item for item in dates if isinstance(item[0], (datetime, date))]
    if len(sortable_dates) == len(dates):
        sortable_dates.sort(key=lambda item: item[0])
        return f"Ngày: {sortable_dates[0][1]} - {sortable_dates[-1][1]}"
    return "Ngày: " + ", ".join(text for _, text in dates[:3])


def committed_norm_fill(value):
    key = text_key(value)
    if "chua co dinh muc" in key:
        return PatternFill("solid", fgColor="FFC7CE")
    if "chua duyet" in key:
        return PatternFill("solid", fgColor="FFF2CC")
    if "da duyet" in key:
        return PatternFill("solid", fgColor="C6EFCE")
    return None


def apply_a4_base_format(ws, last_row):
    thin_gray = Side(style="thin", color="808080")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    header_fill = PatternFill("solid", fgColor="D9E1F2")

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    widths = [12, 6.5, 9, 9.5, 18, 7, 9, 22, 7, 7.5, 10.5, 10, 10.5]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for index in range(A4_LAST_COL + 1, 27):
        ws.column_dimensions[get_column_letter(index)].hidden = True

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=A4_LAST_COL)
    date_cell = ws.cell(1, 1)
    date_cell.font = Font(name="Arial", size=9, bold=True)
    date_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    date_cell.border = border
    ws.row_dimensions[1].height = 22

    for cell in ws[A4_HEADER_ROW]:
        cell.fill = header_fill
        cell.font = Font(name="Arial", size=8, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[A4_HEADER_ROW].height = 32

    center_cols = {1, 2, 3, 4, 7, 10, 12}
    right_cols = {6, 9, 11, 13}
    for row in range(A4_DATA_START_ROW, last_row + 1):
        longest = max(len(str(ws.cell(row, col).value or "")) for col in (5, 8))
        if longest > 85:
            ws.row_dimensions[row].height = 42
        elif longest > 55:
            ws.row_dimensions[row].height = 32
        elif longest > 35:
            ws.row_dimensions[row].height = 24

        for col in range(1, A4_LAST_COL + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Arial", size=8)
            cell.border = border
            if col in right_cols:
                horizontal = "right"
            elif col in center_cols:
                horizontal = "center"
            else:
                horizontal = "left"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)

        ws.cell(row, 6).number_format = "0"
        for col in (9, 11, 13):
            value = ws.cell(row, col).value
            ws.cell(row, col).number_format = "0.###" if isinstance(value, float) and not value.is_integer() else "0"

        status_fill = committed_norm_fill(ws.cell(row, 7).value)
        if status_fill:
            ws.cell(row, 7).fill = status_fill


def same_a4_merge_key(rows, row_a, row_b, key_col, parent_cols):
    values_a = rows[row_a - A4_DATA_START_ROW]
    values_b = rows[row_b - A4_DATA_START_ROW]
    if text_key(values_a[key_col - 1]) == "":
        return False
    if text_key(values_a[key_col - 1]) != text_key(values_b[key_col - 1]):
        return False
    return all(text_key(values_a[parent - 1]) == text_key(values_b[parent - 1]) for parent in parent_cols)


def merge_a4_runs(ws, rows, sheet_col, key_col, parent_cols=()):
    if not rows:
        return
    start = A4_DATA_START_ROW
    last = len(rows) + A4_DATA_START_ROW - 1
    for row in range(A4_DATA_START_ROW + 1, last + 2):
        if row <= last and same_a4_merge_key(rows, row - 1, row, key_col, parent_cols):
            continue
        if row - 1 > start:
            ws.merge_cells(start_row=start, start_column=sheet_col, end_row=row - 1, end_column=sheet_col)
            ws.cell(start, sheet_col).alignment = Alignment(
                horizontal="left" if sheet_col == 5 else "center",
                vertical="center",
                wrap_text=True,
            )
        start = row


def configure_a4_print(ws, last_row):
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.print_title_rows = "1:2"
    ws.print_area = f"A1:{get_column_letter(A4_LAST_COL)}{last_row}"
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1


def process_sheet_format1_a4(ws_source):
    rows = collect_a4_records(ws_source)
    if not rows:
        raise ValueError("File nguồn không có dữ liệu để xuất Format 1.")

    workbook = Workbook()
    ws = workbook.active
    ws.title = "Mau A4"
    ws.append([a4_date_title(rows)])
    ws.append(A4_OUTPUT_HEADERS)
    for row in rows:
        ws.append(row[1:])

    last_row = len(rows) + A4_DATA_START_ROW - 1
    apply_a4_base_format(ws, last_row)
    merge_a4_runs(ws, rows, 1, 2, (1,))
    merge_a4_runs(ws, rows, 2, 3, (1, 2))
    merge_a4_runs(ws, rows, 3, 4, (1, 2, 3))
    merge_a4_runs(ws, rows, 4, 5, (1, 2, 3, 4))
    merge_a4_runs(ws, rows, 5, 6, (1, 2, 3, 4, 5))
    configure_a4_print(ws, last_row)
    return workbook


SHORT_NAMES = {
    "thịt heo nạc vai": "nạc",
    "thịt heo vai": "nạc",
    "thịt heo nạc dăm": "nạc dăm",
    "thịt xay": "xay",
    "thịt bằm": "xay",
    "trứng gà": "trứng",
    "trứng gà 300 quả/cây": "trứng",
    "cá biển chưa xác định": "cá biển",
    "thịt vai heo chưa xác định": "nạc",
    "mọc chưa xác định": "mọc",
    "cá sống chưa xác định": "cá sống",
    "tôm chưa xác định": "tôm",
    "cá basa chưa xác định": "cá basa",
    "đậu hũ chiên chưa xác định": "đậu",
    "đậu hũ trắng": "đậu",
    "thịt vịt chưa xác định": "vịt",
    "trứng cút": "cút",
    "chả cây": "chả",
    "chả lụa": "chả",
    "chả chay": "chả chay",
    "đậu hũ chiên": "đậu",
    "đậu hũ": "đậu",
    "đậu hũ da": "đậu da",
    "đậu cove": "ve",
    "cà rốt": "rốt",
    "nạm bò": "nạm",
    "bò nạm": "nạm",
    "giò heo": "giò",
    "bông cải": "bông cải",
    "sườn heo": "sườn heo",
    "dưa chua": "dưa chua",
    "mọc": "mọc",
    "tôm size 100": "tôm",
    "tôm size 50": "tôm",
    "tôm thẻ size 100 con": "tôm",
    "tôm thẻ size 60 con": "tôm",
    "tôm khô": "tôm khô",
    "măng chua": "măng",
    "trứng": "trứng",
    "giá": "giá",
    "rau dền": "rau dền",
    "rau ngót": "rau ngót",
    "cà tím": "cà tím",
    "hành tây": "hành tây",
    "cá thu": "cá thu",
    "má gà": "gà",
    "má gà tươi": "gà",
    "gà ta": "gà ta",
}

def to_numeric(val):
    if val is None:
        return None
    val_str = str(val).strip()
    if not val_str:
        return None
    try:
        if "." in val_str:
            return float(val_str)
        return int(val_str)
    except ValueError:
        try:
            return float(val_str.replace(",", "."))
        except ValueError:
            return val_str


def format_number_clean(value):
    value = to_numeric(value)
    if isinstance(value, bool) or value is None:
        return "" if value is None else str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.3f}".rstrip("0").rstrip(".")
    return str(value).strip()


def clean_ingredient_name(value):
    text = "" if value is None else str(value).strip()
    if not text:
        return ""

    result = []
    depth = 0
    for char in text:
        if char in "(（":
            depth += 1
            if result and not result[-1].isspace():
                result.append(" ")
            continue
        if char in ")）":
            if depth > 0:
                depth -= 1
            continue
        if depth == 0:
            result.append(char)

    return " ".join("".join(result).split())


def is_blank_cell_value(value):
    return value is None or str(value).strip() == ""


def find_approval_source_columns(ws):
    required_keys = {
        "id",
        "status",
        "dish_code",
        "dish_group",
        "dish",
        "ingredient",
        "ingredient_no",
        "unit",
        "purchase_unit",
    }

    for row in range(1, min(ws.max_row, 10) + 1):
        columns = {}
        for col in range(1, ws.max_column + 1):
            header_text = clean_header(ws.cell(row, col).value)
            if not header_text:
                continue
            if header_text == "id":
                columns.setdefault("id", col)
            elif header_text == "trang thai":
                columns.setdefault("status", col)
            elif header_text == "ngay duyet":
                columns.setdefault("approval_date", col)
            elif header_text == "nhom mon an theo nguyen lieu chinh":
                columns.setdefault("main_group", col)
            elif header_text == "nhom mon an theo nvl cap 1":
                columns.setdefault("ingredient_group", col)
            elif header_text == "ma mon an":
                columns.setdefault("dish_code", col)
            elif header_text == "nhom mon":
                columns.setdefault("dish_group", col)
            elif header_text == "ten mon an":
                columns.setdefault("dish", col)
            elif header_text in {"ten nguyen vat lieu", "nguyen vat lieu", "ten nvl"}:
                columns.setdefault("ingredient", col)
            elif header_text in {"stt nvl", "stt nguyen vat lieu"}:
                columns.setdefault("ingredient_no", col)
            elif header_text in {"don vi tinh di cho", "dvt di cho", "don vi tinh mua hang", "dvt mua hang"}:
                columns.setdefault("purchase_unit", col)
            elif header_text in {"don vi tinh", "dvt"}:
                columns.setdefault("unit", col)
            elif header_text == "version":
                columns.setdefault("version", col)
            elif header_text == "can duyet":
                columns.setdefault("needs_review", col)

        if not required_keys.issubset(columns):
            continue

        sentinel_cols = [
            col for col in (columns.get("version"), columns.get("needs_review"))
            if col and col > columns["purchase_unit"]
        ]
        customer_end = min(sentinel_cols) - 1 if sentinel_cols else ws.max_column
        customer_cols = []
        for col in range(columns["purchase_unit"] + 1, customer_end + 1):
            header_value = ws.cell(row, col).value
            header_text = clean_header(header_value)
            if header_text and header_text not in {"version", "can duyet"}:
                customer_cols.append((col, str(header_value).strip()))

        if customer_cols:
            return row, columns, customer_cols

    raise ValueError("Không tìm thấy đủ cột để xuất format Duyệt định mức.")


CUSTOMER_SHORT_NAMES = {
    "mon an sang om": "Ăn sáng OM",
    "scavi com khach 50k": "SCAVI 50K",
    "scavi com khach 70k": "SCAVI 70K",
    "sofa com quan ly 40k": "SOFA QL",
}


def short_customer_name(name):
    return CUSTOMER_SHORT_NAMES.get(text_key(name), name)


def concise_customer_label(names, all_names):
    missing = [name for name in all_names if name not in names]
    short_names = [short_customer_name(name) for name in names]
    short_missing = [short_customer_name(name) for name in missing]
    if len(names) == len(all_names):
        return "Tất cả KH"
    if len(names) >= 8:
        if missing and len(missing) <= 5:
            return f"{len(names)} KH (trừ {', '.join(short_missing)})"
        return f"{len(names)} KH ({', '.join(short_names[:5])}, ...)"
    return ", ".join(short_names)


def approval_norm_summary(ws, lookup, row, customer_cols):
    all_names = [name for _, name in customer_cols]
    grouped = {}
    order = []
    for col, customer_name in customer_cols:
        value = effective_cell_value(ws, lookup, row, col)
        if is_blank_cell_value(value):
            continue
        display_value = format_number_clean(value)
        if display_value not in grouped:
            grouped[display_value] = []
            order.append(display_value)
        grouped[display_value].append(customer_name)

    lines = []
    for display_value in order:
        customer_label = concise_customer_label(grouped[display_value], all_names)
        lines.append(f"{customer_label}: {display_value}")
    return "\n".join(lines)


def collect_approval_records(ws):
    header_row, columns, customer_cols = find_approval_source_columns(ws)
    lookup = merged_value_lookup(ws)
    rows = []

    for source_row in range(header_row + 1, ws.max_row + 1):
        raw = {
            key: effective_cell_value(ws, lookup, source_row, col)
            for key, col in columns.items()
        }
        if all(is_blank_cell_value(value) for value in raw.values()):
            continue

        approval_date = raw.get("approval_date")
        parsed_date = parse_date_value(approval_date)
        ingredient_group = raw.get("ingredient_group") or raw.get("main_group")
        values = [
            raw.get("id"),
            raw.get("status"),
            parsed_date or approval_date,
            raw.get("needs_review"),
            raw.get("dish_code"),
            raw.get("dish_group"),
            raw.get("dish"),
            ingredient_group,
            clean_ingredient_name(raw.get("ingredient")),
            parse_print_number(raw.get("ingredient_no")),
            raw.get("unit"),
            raw.get("purchase_unit"),
            approval_norm_summary(ws, lookup, source_row, customer_cols),
            raw.get("version"),
        ]
        if not a4_row_empty(values):
            rows.append(values)

    return rows


def approval_attention_fill(value, is_status=False):
    key = text_key(value)
    if not key:
        return None
    if "da duyet" in key or "chuan" in key:
        return PatternFill("solid", fgColor="C6EFCE")
    if "chua duyet" in key:
        return PatternFill("solid", fgColor="FFF2CC")
    if "can duyet" in key:
        return PatternFill("solid", fgColor="FCE4D6")
    if not is_status:
        return PatternFill("solid", fgColor="FCE4D6")
    return None


def estimate_wrapped_lines(value, chars_per_line):
    text = "" if value is None else str(value)
    if not text:
        return 1
    return sum(
        max(1, (len(line) + chars_per_line - 1) // chars_per_line)
        for line in text.splitlines()
    )


def apply_approval_base_format(ws, last_row):
    thin_gray = Side(style="thin", color="B7C0CC")
    medium_gray = Side(style="medium", color="6B7280")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    header_border = Border(left=thin_gray, right=thin_gray, top=medium_gray, bottom=medium_gray)
    title_fill = PatternFill("solid", fgColor="1F4E78")
    header_fill = PatternFill("solid", fgColor="DDEBF7")

    ws.title = "Duyệt định mức"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    widths = [8, 11, 10, 12, 10, 11, 24, 13, 25, 5.5, 6, 8, 38, 8]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for index in range(APPROVAL_LAST_COL + 1, 35):
        ws.column_dimensions[get_column_letter(index)].hidden = True

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=APPROVAL_LAST_COL)
    title_cell = ws.cell(1, 1)
    title_cell.font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    title_cell.border = header_border
    ws.row_dimensions[1].height = 28

    for col in range(1, APPROVAL_LAST_COL + 1):
        cell = ws.cell(APPROVAL_HEADER_ROW, col)
        cell.fill = header_fill
        cell.font = Font(name="Arial", size=8.5, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border
    ws.row_dimensions[APPROVAL_HEADER_ROW].height = 32

    center_cols = {1, 2, 3, 4, 5, 6, 10, 11, 12, 14}
    for row in range(APPROVAL_DATA_START_ROW, last_row + 1):
        line_count = max(
            estimate_wrapped_lines(ws.cell(row, 7).value, 24),
            estimate_wrapped_lines(ws.cell(row, 9).value, 25),
            estimate_wrapped_lines(ws.cell(row, 13).value, 38),
        )
        ws.row_dimensions[row].height = min(80, max(20, 13 * line_count + 6))

        for col in range(1, APPROVAL_LAST_COL + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Arial", size=8.5)
            cell.border = border
            horizontal = "center" if col in center_cols else "left"
            cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)

        ws.cell(row, 3).number_format = "dd/mm/yyyy"
        ws.cell(row, 10).number_format = "0"
        status_fill = approval_attention_fill(ws.cell(row, 2).value, is_status=True)
        if status_fill:
            ws.cell(row, 2).fill = status_fill
        review_fill = approval_attention_fill(ws.cell(row, 4).value)
        if review_fill:
            ws.cell(row, 4).fill = review_fill

    ws.auto_filter.ref = f"A{APPROVAL_HEADER_ROW}:{get_column_letter(APPROVAL_LAST_COL)}{last_row}"
    ws.print_area = f"A1:{get_column_letter(APPROVAL_LAST_COL)}{last_row}"
    ws.print_title_rows = "1:2"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3
    ws.page_margins.header = 0.12
    ws.page_margins.footer = 0.12


def approval_dish_key(values):
    if not values:
        return None
    code_key = text_key(values[4])
    dish_key = text_key(values[6])
    if not code_key and not dish_key:
        return None
    return code_key, dish_key


def add_approval_page_breaks(ws, rows, max_body_rows=34):
    ws.row_breaks.brk = []
    if not rows:
        return set()

    last_row = len(rows) + APPROVAL_DATA_START_ROW - 1
    boundaries = [
        APPROVAL_DATA_START_ROW + index
        for index in range(1, len(rows))
        if approval_dish_key(rows[index]) != approval_dish_key(rows[index - 1])
    ]
    break_rows = set()
    page_start = APPROVAL_DATA_START_ROW
    while page_start + max_body_rows <= last_row:
        latest_start = page_start + max_body_rows
        earliest_start = page_start + max(14, max_body_rows // 2)
        candidates = [
            row for row in boundaries
            if earliest_start <= row <= latest_start
        ]
        next_page = max(candidates or [latest_start])
        if next_page <= page_start:
            next_page = latest_start
        break_rows.add(next_page)
        ws.row_breaks.append(Break(id=next_page - 1))
        page_start = next_page
    return break_rows


def merge_approval_runs(ws, rows, page_break_rows=None):
    page_break_rows = set(page_break_rows or ())
    if not rows:
        return

    def can_merge_col(start, end, col):
        first = ws.cell(start, col).value
        return all(text_key(ws.cell(row, col).value) == text_key(first) for row in range(start + 1, end + 1))

    def merge_col(start, end, col):
        if end <= start or not can_merge_col(start, end, col):
            return
        ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
        horizontal = "left" if col == 7 else "center"
        ws.cell(start, col).alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)

    start = APPROVAL_DATA_START_ROW
    previous_key = approval_dish_key(rows[0])
    last = len(rows) + APPROVAL_DATA_START_ROW - 1
    for row in range(APPROVAL_DATA_START_ROW + 1, last + 2):
        current_key = approval_dish_key(rows[row - APPROVAL_DATA_START_ROW]) if row <= last else None
        should_close = row in page_break_rows or current_key != previous_key or previous_key is None
        if should_close:
            end = row - 1
            for col in (5, 6, 7):
                merge_col(start, end, col)
            start = row
            previous_key = current_key


def process_sheet_approval(ws_source):
    rows = collect_approval_records(ws_source)
    if not rows:
        raise ValueError("File nguồn không có dữ liệu để xuất format Duyệt định mức.")

    workbook = Workbook()
    ws = workbook.active
    ws.append(["DUYỆT ĐỊNH MỨC"])
    ws.append(APPROVAL_OUTPUT_HEADERS)
    for row in rows:
        ws.append(row)

    last_row = len(rows) + APPROVAL_DATA_START_ROW - 1
    apply_approval_base_format(ws, last_row)
    page_break_rows = add_approval_page_breaks(ws, rows)
    merge_approval_runs(ws, rows, page_break_rows)
    return workbook


def strip_accents(text):
    if not text:
        return ""
    text = text.replace("đ", "d").replace("Đ", "D")
    normalized = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in normalized if not unicodedata.combining(c))

def get_short_ingredient_repr(dish_name, nvl_name, dm_val, dm_unit):
    nvl_clean = clean_ingredient_name(nvl_name).lower()
    dish_clean = str(dish_name).strip().lower()
    
    nvl_no_accent = strip_accents(nvl_clean).replace(" ", "")
    dish_no_accent = strip_accents(dish_clean).replace(" ", "")
    
    for word in ["thitheo", "thit", "heo", "con"]:
        nvl_no_accent = nvl_no_accent.replace(word, "")
        
    qty_str = ""
    if dm_val is not None:
        if isinstance(dm_val, float):
            qty_str = f"{dm_val:.3f}".rstrip('0').rstrip('.')
        else:
            qty_str = str(dm_val)
            
    unit_clean = str(dm_unit).strip().lower()
    short_name = SHORT_NAMES.get(nvl_clean, nvl_clean)
    
    if short_name == "nạc" and "xào" in dish_clean:
        short_name = "nạc xắt"
        
    omit_name = False
    if nvl_no_accent and dish_no_accent:
        if nvl_no_accent in dish_no_accent or dish_no_accent in nvl_no_accent:
            omit_name = True
        elif strip_accents(short_name).replace(" ", "") in dish_no_accent:
            omit_name = True
            
    if not qty_str:
        if omit_name:
            return ""
        return short_name
        
    if omit_name:
        if unit_clean in ("gr", "g"):
            return f"{qty_str}g"
        elif unit_clean in ("cái", "c"):
            return f"{qty_str}c"
        else:
            return qty_str
            
    if unit_clean in ("gr", "g"):
        return f"{qty_str}g {short_name}"
    elif unit_clean in ("cái", "c"):
        return f"{qty_str}c {short_name}"
    elif unit_clean in ("quả", "q"):
        return f"{qty_str} {short_name}"
    elif unit_clean in ("miếng", "m"):
        return f"{qty_str} {short_name}"
    else:
        suffix = f" {unit_clean}" if unit_clean else ""
        return f"{qty_str}{suffix} {short_name}"

def normalize_comp(name):
    if not name:
        return ""
    s = str(name).lower()
    s = re.sub(r'[^a-z0-9]', '', s)
    if 'omdigital' in s or 'odsv' in s or 'osv' in s:
        return 'osv'
    if 'lixil' in s:
        return 'lixil'
    if 'smc' in s:
        return 'smc'
    if 'zamil' in s:
        return 'zamil'
    if 'wacoal' in s:
        return 'wacoal'
    if 'quadrille' in s:
        return 'quadrille'
    if 'watabe' in s:
        return 'watabe'
    if 'vikyno' in s:
        return 'vikyno'
    if 'shine' in s:
        return 'shine'
    if 'catthai' in s or 'cat-thai' in s or 'cat_thai' in s:
        return 'catthai'
    if 'maspro' in s:
        return 'maspro'
    if 'medic' in s:
        return 'medic'
    if 'scavi' in s:
        if '50k' in s or '50' in s:
            return 'scavicomkhach50k'
        if '70k' in s or '70' in s:
            return 'scavicomkhach70k'
        return 'scavi'
    if 'sofa' in s:
        if 'ql' in s or 'quanly' in s or 'comql' in s or '40k' in s:
            return 'sofacomql40k'
        return 'sofa'
    if 'dona' in s:
        return 'donaquebang'
    if 'saitexd13' in s:
        return 'saitexd13'
    if 'saitex4' in s:
        return 'saitex4'
    if 'saitex6' in s:
        return 'saitex6icd'
    if 'figla' in s:
        return 'figla'
    if 'shiseido' in s or 'shisedo' in s:
        return 'shisedo7a'
    if 'jfe' in s:
        return 'jfe'
    if 'briskheat' in s or 'brishheart' in s:
        return 'brishheart'
    if 'artwell' in s:
        return 'artwell'
    return s

def normalize_shift(shift_name):
    if not shift_name:
        return ""
    s = str(shift_name).lower()
    if 'ca 1' in s or 'ca1' in s or 'canteen 1' in s or 'c1' in s:
        return 'ca1'
    if 'ca 2' in s or 'ca2' in s or 'canteen 2' in s or 'c2' in s:
        return 'ca2'
    if 'ca 3' in s or 'ca3' in s or 'c3' in s:
        return 'ca3'
    if 'ăn sáng' in s or 'ansang' in s:
        if '2' in s:
            return 'ansang2'
        return 'ansang'
    return s

def get_dish_category(dish_name, nhom_mon=None):
    d_lower = str(dish_name).lower()
    nm_lower = str(nhom_mon).lower() if nhom_mon else ""
    if 'chay' in nm_lower or 'chay' in d_lower:
        return 'chay'
    if any(x in nm_lower for x in ['tráng miệng', 'trang mieng', 'trái cây', 'trai cay']) or        any(x in d_lower for x in ['trái cây', 'trai cay', 'chuối', 'nhãn', 'dưa hấu', 'sữa chua', 'bánh da lợn']):
        return 'trangmieng'
    if any(x in nm_lower for x in ['nước', 'nuoc', 'cháo', 'chao']):
        return 'nuoc'
    if any(d_lower.startswith(x) for x in ['bún ', 'mì ', 'hủ tíu', 'canh bún', 'bánh đa', 'nui ', 'phở ', 'cháo ']):
        return 'nuoc'
    return 'main'

def get_vietnamese_weekday(date_str):
    if " - " in date_str:
        date_str = date_str.split(" - ")[0]
    try:
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        weekday = dt.weekday()
        mapping = {
            0: "Thứ 2",
            1: "Thứ 3",
            2: "Thứ 4",
            3: "Thứ 5",
            4: "Thứ 6",
            5: "Thứ 7",
            6: "Chủ nhật"
        }
        return f"{mapping.get(weekday, '')} - "
    except Exception:
        return "Thứ … - "

def get_shift_order(sh_n):
    if sh_n == 'ca1': return 1
    if sh_n == 'ca2': return 2
    if sh_n == 'ca3': return 3
    if sh_n == 'ansang': return 4
    if sh_n == 'ansang2': return 5
    return 99

def get_cat_order(cat_n):
    if cat_n == 'main': return 1
    if cat_n == 'nuoc': return 2
    if cat_n == 'chay': return 3
    if cat_n == 'trangmieng': return 4
    return 99


def build_compact_shopping_workbook(
    parsed_rows,
    col_ma_kh_idx,
    col_ca_idx,
    col_mon_idx,
    col_qty_idx,
    col_nvl_idx,
    col_dinh_muc_idx,
    col_dinh_muc_unit_idx,
    col_nhom_mon_idx,
    col_site_an_idx,
    date_text,
    weekday_text,
    target_sheet_name,
):
    """Create a compact, two-panel A4 shopping sheet from BOM rows."""
    lixil_key = normalize_comp("LIXIL")
    company_names = {}
    groups = {}

    for row in parsed_rows:
        company = row[col_ma_kh_idx] if col_ma_kh_idx is not None else ""
        shift = row[col_ca_idx] if col_ca_idx is not None else ""
        dish = row[col_mon_idx] if col_mon_idx is not None else ""
        qty = row[col_qty_idx] if col_qty_idx is not None else ""
        category_value = row[col_nhom_mon_idx] if col_nhom_mon_idx is not None else ""
        site = row[col_site_an_idx] if col_site_an_idx is not None else ""
        ingredient = row[col_nvl_idx] if col_nvl_idx is not None else ""
        amount = row[col_dinh_muc_idx] if col_dinh_muc_idx is not None else ""
        amount_unit = row[col_dinh_muc_unit_idx] if col_dinh_muc_unit_idx is not None else ""

        if not company or not dish:
            continue

        company_key = normalize_comp(company)
        company_names.setdefault(company_key, str(company).strip())
        site_text = str(site).strip() if company_key == lixil_key else ""
        shift_key = normalize_shift(shift)
        category_key = get_dish_category(dish, category_value)
        group_key = (company_key, site_text, shift_key, category_key, str(dish).strip())

        if group_key not in groups:
            groups[group_key] = {
                "qty": to_numeric(qty) or 0,
                "ingredients": [],
                "seen_ingredients": set(),
            }

        ingredient_text = clean_ingredient_name(ingredient)
        ingredient_key = ingredient_text.casefold()
        if ingredient_text and ingredient_key not in groups[group_key]["seen_ingredients"]:
            groups[group_key]["seen_ingredients"].add(ingredient_key)
            groups[group_key]["ingredients"].append(
                (ingredient_text, to_numeric(amount), amount_unit)
            )

    if not groups:
        raise ValueError("File BOM không có dữ liệu món ăn để tạo giấy đi chợ.")

    block_rows = {}
    lixil_dishes = {}

    def site_label(site_text):
        normalized = str(site_text).strip().casefold().replace(" ", "").replace("-", "")
        if "fab1" in normalized:
            return "FAB1"
        if "fab2" in normalized:
            return "FAB2"
        if "vpc" in normalized:
            return "VPC"
        return str(site_text).strip().upper() or "SITE"

    def shift_label_for(shift_key, category_key):
        if category_key == "trangmieng":
            return "T.miệng"
        if category_key == "nuoc":
            return "M.nước"
        if category_key == "chay":
            return "Chay"
        return {
            "ca1": "Ca 1",
            "ca2": "Ca 2",
            "ca3": "Ca 3",
            "ansang": "Ăn sáng",
            "ansang2": "Ăn sáng 2",
        }.get(shift_key, str(shift_key).upper())

    for key, data in groups.items():
        company_key, site_text, shift_key, category_key, dish = key
        recipe_parts = []
        for ingredient, amount, amount_unit in data["ingredients"]:
            recipe = get_short_ingredient_repr(dish, ingredient, amount, amount_unit)
            if recipe:
                recipe_parts.append(recipe)
        recipe_text = "+".join(recipe_parts)
        sort_key = (get_shift_order(shift_key), get_cat_order(category_key), dish.casefold())

        if company_key == lixil_key:
            dish_key = (shift_key, category_key, dish)
            lixil_entry = lixil_dishes.setdefault(
                dish_key,
                {
                    "sort": sort_key,
                    "shift": shift_label_for(shift_key, category_key),
                    "dish": dish,
                    "sites": {},
                },
            )
            normalized_site = site_label(site_text)
            lixil_entry["sites"][normalized_site] = {
                "qty": data["qty"],
                "recipe": recipe_text,
                "order": get_lixil_site_order(site_text),
            }
            continue

        block_rows.setdefault((company_key, ""), []).append(
            {
                "sort": sort_key,
                "qty": data["qty"],
                "shift": shift_label_for(shift_key, category_key),
                "dish": dish,
                "recipe": recipe_text,
            }
        )

    for lixil_entry in lixil_dishes.values():
        ordered_sites = sorted(
            lixil_entry["sites"].items(),
            key=lambda item: (item[1]["order"], item[0].casefold()),
        )
        quantity_text = "+".join(
            format_number_clean(site_data["qty"])
            for _site, site_data in ordered_sites
        )
        recipe_text = "\n".join(
            f"{site} - {site_data['recipe']}"
            for site, site_data in ordered_sites
        )
        block_rows.setdefault((lixil_key, ""), []).append(
            {
                "sort": lixil_entry["sort"],
                "qty": quantity_text,
                "shift": lixil_entry["shift"],
                "dish": lixil_entry["dish"],
                "recipe": recipe_text,
            }
        )

    blocks = []
    for (company_key, site_text), rows in block_rows.items():
        rows.sort(key=lambda item: item["sort"])
        label = company_names.get(company_key, company_key.upper())
        blocks.append(
            {
                "company_key": company_key,
                "site": site_text,
                "label": label,
                "rows": rows,
            }
        )

    blocks.sort(
        key=lambda block: (
            block["company_key"],
            get_lixil_site_order(block["site"]) if block["company_key"] == lixil_key else 0,
            block["site"].casefold(),
        )
    )

    left_blocks = []
    right_blocks = []
    left_count = 0
    right_count = 0
    for block in blocks:
        if not left_blocks or left_count <= right_count:
            left_blocks.append(block)
            left_count += len(block["rows"])
        else:
            right_blocks.append(block)
            right_count += len(block["rows"])

    left_row_count = sum(len(block["rows"]) for block in left_blocks)
    right_row_count = sum(len(block["rows"]) for block in right_blocks)
    max_panel_rows = max(left_row_count, right_row_count)
    page_break_rows = set()
    if max_panel_rows > 14:
        page_break_rows.add(3 + max_panel_rows // 2)

    workbook = Workbook()
    worksheet = workbook.active
    sheet_title = "Bếp trung tâm" if target_sheet_name == "Bếp trung tâm 1" else target_sheet_name
    worksheet.title = sheet_title[:31]

    sheet_title_key = sheet_title.casefold()
    if "tại chỗ" in sheet_title_key:
        kitchen_label = "BẾP TẠI CHỖ"
    elif "trung tâm 2" in sheet_title_key:
        kitchen_label = "BẾP TRUNG TÂM 2"
    else:
        kitchen_label = "BẾP TRUNG TÂM"
    weekday_label = weekday_text.strip().rstrip("-").strip()
    title = f"GIẤY ĐI CHỢ - {kitchen_label}"
    if weekday_label or date_text:
        title += f" | {weekday_label} {date_text}".rstrip()

    worksheet.merge_cells("A1:J1")
    title_cell = worksheet["A1"]
    title_cell.value = title
    title_cell.font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor="404040")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 32

    headers = ["Công ty", "SL theo Site", "Ca", "Món ăn", "Khối lượng đi chợ"]
    header_fill = PatternFill("solid", fgColor="BFBFBF")
    thin = Side(style="thin", color="D9D9D9")
    group_side = Side(style="medium", color="A6A6A6")
    body_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for start_col in (1, 6):
        for offset, value in enumerate(headers):
            cell = worksheet.cell(2, start_col + offset)
            cell.value = value
            cell.font = Font(name="Calibri", size=14, bold=True, color="000000")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=group_side, bottom=group_side)
    worksheet.row_dimensions[2].height = 34

    company_fill = PatternFill("solid", fgColor="E7E6E6")

    def merge_segments(start_row, end_row):
        cuts = sorted(row for row in page_break_rows if start_row < row <= end_row)
        starts = [start_row] + cuts
        ends = [row - 1 for row in cuts] + [end_row]
        return list(zip(starts, ends))

    def write_panel(panel_blocks, start_col):
        def estimate_lines(value, chars_per_line):
            text = str(value or "")
            return sum(
                max(1, (len(line) + chars_per_line - 1) // chars_per_line)
                for line in text.splitlines() or [""]
            )

        current_row = 3
        for block in panel_blocks:
            block_start = current_row
            for item in block["rows"]:
                values = [block["label"], item["qty"], item["shift"], item["dish"], item["recipe"]]
                for offset, value in enumerate(values):
                    cell = worksheet.cell(current_row, start_col + offset)
                    cell.value = value
                    cell.font = Font(
                        name="Calibri",
                        size=14 if offset == 0 else 13,
                        bold=offset in (0, 2),
                    )
                    cell.border = body_border
                    cell.alignment = Alignment(
                        horizontal="center" if offset in (0, 1, 2) else "left",
                        vertical="center",
                        wrap_text=offset != 1,
                    )
                worksheet.cell(current_row, start_col + 1).number_format = "General"
                estimated_lines = max(
                    1,
                    estimate_lines(item["dish"], 18),
                    estimate_lines(item["recipe"], 26),
                )
                worksheet.row_dimensions[current_row].height = max(
                    worksheet.row_dimensions[current_row].height or 0,
                    min(190, max(26, 19 * estimated_lines + 4)),
                )
                current_row += 1

            block_end = current_row - 1
            for row_num in range(block_start, block_end + 1):
                company_cell = worksheet.cell(row_num, start_col)
                company_cell.fill = company_fill
                company_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            shift_col = start_col + 2
            shift_start = block_start
            previous_shift = worksheet.cell(block_start, shift_col).value
            for row_num in range(block_start + 1, block_end + 2):
                shift_value = worksheet.cell(row_num, shift_col).value if row_num <= block_end else None
                if row_num in page_break_rows or shift_value != previous_shift:
                    if row_num - 1 > shift_start:
                        worksheet.merge_cells(
                            start_row=shift_start,
                            start_column=shift_col,
                            end_row=row_num - 1,
                            end_column=shift_col,
                        )
                    shift_start = row_num
                    previous_shift = shift_value
        return current_row - 1

    left_end = write_panel(left_blocks, 1)
    right_end = write_panel(right_blocks, 6)
    last_row = max(left_end, right_end, 3)

    def merge_adjacent_same_company(start_col, start_row, end_row):
        if end_row < start_row:
            return

        segment_start = start_row
        previous_value = worksheet.cell(start_row, start_col).value

        for row_num in range(start_row + 1, end_row + 2):
            current_value = worksheet.cell(row_num, start_col).value if row_num <= end_row else None
            should_close = row_num in page_break_rows or current_value != previous_value
            if should_close:
                if previous_value and row_num - 1 > segment_start:
                    worksheet.merge_cells(
                        start_row=segment_start,
                        start_column=start_col,
                        end_row=row_num - 1,
                        end_column=start_col,
                    )
                company_cell = worksheet.cell(segment_start, start_col)
                company_cell.value = previous_value
                company_cell.fill = company_fill
                company_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                segment_start = row_num
                previous_value = current_value

    merge_adjacent_same_company(1, 3, left_end)
    merge_adjacent_same_company(6, 3, right_end)

    for row in range(3, last_row + 1):
        for col in range(1, 11):
            if worksheet.cell(row, col).value is None:
                worksheet.cell(row, col).border = body_border

    widths = {
        "A": 11, "B": 27, "C": 9, "D": 19, "E": 30,
        "F": 11, "G": 27, "H": 9, "I": 19, "J": 30,
    }
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width

    worksheet.freeze_panes = "A3"
    worksheet.sheet_view.showGridLines = False
    worksheet.print_area = f"A1:J{last_row}"
    worksheet.print_title_rows = "1:2"
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.page_setup.blackAndWhite = True
    worksheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    for break_row in sorted(page_break_rows):
        worksheet.row_breaks.append(Break(id=break_row - 1))
    worksheet.print_options.horizontalCentered = True
    worksheet.page_margins.left = 0.2
    worksheet.page_margins.right = 0.2
    worksheet.page_margins.top = 0.25
    worksheet.page_margins.bottom = 0.25
    worksheet.page_margins.header = 0.15
    worksheet.page_margins.footer = 0.15
    return workbook


def process_sheet_format2(ws_source, file_path, date_mode="auto", selected_kitchen=None):
    # 1. Map columns from source header
    col_map = {}
    header = [ws_source.cell(1, c).value for c in range(1, ws_source.max_column + 1)]
    for idx, col in enumerate(header):
        if col is None:
            continue
        c_lower = str(col).lower().strip()
        if any(x in c_lower for x in ["ngày", "ngay", "date"]):
            col_map["ngay"] = idx
        elif any(x in c_lower for x in ["khách hàng", "khach hang", "mã kh", "ma kh", "khach_hang", "mã khách hàng"]) and "site" not in c_lower:
            col_map["ma_kh"] = idx
        elif any(x in c_lower for x in ["ca", "shift"]):
            col_map["ca"] = idx
        elif any(x in c_lower for x in ["nhóm món", "nhom mon", "cơ cấu món ăn", "co cau mon an", "cơ cấu menu", "co cau menu"]):
            col_map["nhom_mon"] = idx
        # BUG FIX: Exclude 'cơ cấu' and 'co cau' from matching 'mon' (dish name)
        elif any(x in c_lower for x in ["món ăn", "mon an", "tên món", "ten mon"]) and "cơ cấu" not in c_lower and "co cau" not in c_lower:
            col_map["mon"] = idx
        elif any(x in c_lower for x in ["số lượng", "so luong", "s.lượng", "qty", "quantity"]):
            col_map["qty"] = idx
        elif any(x in c_lower for x in ["nguyên vật liệu", "nguyen vat lieu", "nvl", "tên nguyên vật liệu"]):
            col_map["nvl"] = idx
        elif any(x in c_lower for x in ["đvt", "đơn vị", "don vi", "dvt"]) and "đi chợ" not in c_lower and "di cho" not in c_lower and "mua hàng" not in c_lower and "mua hang" not in c_lower:
            col_map["dinh_muc_unit"] = idx
        elif any(x in c_lower for x in ["định mức", "dinh muc", "khối lượng", "định lượng"]) and "đi chợ" not in c_lower and "di cho" not in c_lower and "yêu cầu" not in c_lower and "yeu cau" not in c_lower and "mua hàng" not in c_lower and "mua hang" not in c_lower:
            col_map["dinh_muc"] = idx
        elif any(x in c_lower for x in ["no.", "stt"]):
            col_map["no"] = idx
        elif any(x in c_lower for x in ["site ăn", "site an", "site"]):
            col_map["site_an"] = idx

    col_ngay_idx = col_map.get("ngay")
    col_ma_kh_idx = col_map.get("ma_kh")
    col_ca_idx = col_map.get("ca")
    col_mon_idx = col_map.get("mon")
    col_qty_idx = col_map.get("qty")
    col_nvl_idx = col_map.get("nvl")
    col_dinh_muc_idx = col_map.get("dinh_muc")
    col_dinh_muc_unit_idx = col_map.get("dinh_muc_unit")
    col_no_idx = col_map.get("no")
    col_nhom_mon_idx = col_map.get("nhom_mon")
    col_site_an_idx = col_map.get("site_an")

    # Read rows from ws_source starting at row 2
    raw_rows = []
    for r in range(2, ws_source.max_row + 1):
        row_vals = [ws_source.cell(r, c).value for c in range(1, ws_source.max_column + 1)]
        if any(x is not None for x in row_vals):
            raw_rows.append(row_vals)

    # Forward fill logic
    prev_vals = {}
    parsed_rows = []
    for row in raw_rows:
        row_vals = list(row)
        for key, idx in [("ngay", col_ngay_idx), ("ma_kh", col_ma_kh_idx), ("ca", col_ca_idx), ("mon", col_mon_idx), ("qty", col_qty_idx), ("no", col_no_idx)]:
            if idx is not None and idx < len(row_vals):
                val = str(row_vals[idx]).strip() if row_vals[idx] is not None else ""
                if not val:
                    val = prev_vals.get(key, "")
                row_vals[idx] = val
                prev_vals[key] = val
        parsed_rows.append(row_vals)

    # Date extraction using our improved parser
    dates = []
    seen_dates = set()
    if col_ngay_idx is not None:
        for row in parsed_rows:
            val = row[col_ngay_idx]
            dt_val = parse_date_value(val)
            text = format_date(dt_val) if dt_val else str(val).strip()
            if text and text not in seen_dates:
                seen_dates.add(text)
                dates.append((dt_val or val, text))

    date_text = ""
    if dates:
        if len(dates) == 1:
            date_text = dates[0][1]
        else:
            sortable_dates = [x for x in dates if isinstance(x[0], (datetime, date))]
            if len(sortable_dates) == len(dates):
                sortable_dates.sort(key=lambda x: x[0])
                date_text = f"{sortable_dates[0][1]} - {sortable_dates[-1][1]}"
            else:
                date_text = ", ".join(t for _, t in dates[:3])
    else:
        # Fallback from filename
        year = datetime.now().year
        fn_date = extract_date_from_filename(file_path, year)
        if fn_date:
            date_text = fn_date

    weekday_text = get_vietnamese_weekday(date_text) if date_text else "Thứ … - "

    # Determine kitchen sheet based on filename or selected_kitchen
    if selected_kitchen:
        if selected_kitchen == "Bếp trung tâm":
            target_sheet_name = "Bếp trung tâm 1"
        else:
            target_sheet_name = selected_kitchen
    else:
        filename = file_path.name.casefold()
        if "bếp trung tâm 2" in filename or "bep trung tam 2" in filename:
            target_sheet_name = "Bếp trung tâm 2"
        elif "bếp tại chỗ" in filename or "bep tai cho" in filename:
            target_sheet_name = "Bếp tại chỗ"
        elif "bếp trung tâm" in filename or "bep trung tam" in filename:
            target_sheet_name = "Bếp trung tâm 1"
        else:
            target_sheet_name = "Bếp tại chỗ"

    return build_compact_shopping_workbook(
        parsed_rows,
        col_ma_kh_idx,
        col_ca_idx,
        col_mon_idx,
        col_qty_idx,
        col_nvl_idx,
        col_dinh_muc_idx,
        col_dinh_muc_unit_idx,
        col_nhom_mon_idx,
        col_site_an_idx,
        date_text,
        weekday_text,
        target_sheet_name,
    )

def format_workbook_bytes(file_bytes, filename, date_mode="auto", format_mode="format1", selected_kitchen=None):
    if not file_bytes:
        raise ValueError("File Excel đang trống.")

    safe_name = safe_filename(filename)
    suffix = Path(safe_name).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".csv"}:
        raise ValueError("Vui lòng chọn file Excel .xlsx, .xlsm hoặc .csv.")

    if suffix == ".csv":
        from io import StringIO
        workbook = Workbook()
        worksheet = workbook.active
        text = file_bytes.decode("utf-8-sig", errors="ignore")
        reader = csv.reader(StringIO(text))
        for row in reader:
            worksheet.append(row)
    else:
        workbook = load_workbook(BytesIO(file_bytes))
        worksheet = ensure_visible_worksheet(workbook, workbook.active)

    if format_mode == "format2":
        workbook = process_sheet_format2(worksheet, Path(safe_name), date_mode, selected_kitchen)
    elif format_mode == APPROVAL_FORMAT_MODE:
        workbook = process_sheet_approval(worksheet)
    else:
        workbook = process_sheet_format1_a4(worksheet)

    output = BytesIO()
    ensure_visible_worksheet(workbook)
    workbook.save(output)
    workbook.close()
    return output.getvalue()




def parse_multipart_form(headers, body):
    content_type = headers.get("Content-Type", "")
    marker = "boundary="
    if marker not in content_type:
        raise ValueError("Form upload không hợp lệ.")

    boundary = content_type.split(marker, 1)[1].split(";", 1)[0].strip().strip('"')
    if not boundary:
        raise ValueError("Không đọc được boundary của file upload.")

    delimiter = b"--" + boundary.encode("utf-8")
    fields = {}
    file_bytes = None
    filename = None

    for raw_part in body.split(delimiter):
        part = raw_part.strip()
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].strip()
        if b"\r\n\r\n" not in part:
            continue

        raw_headers, val_data = part.split(b"\r\n\r\n", 1)
        header_text = raw_headers.decode("utf-8", "replace")
        disposition_line = next(
            (
                line
                for line in header_text.splitlines()
                if line.lower().startswith("content-disposition:")
            ),
            "",
        )
        disposition_params = {}
        for segment in disposition_line.split(";")[1:]:
            if "=" not in segment:
                continue
            key, value = segment.split("=", 1)
            disposition_params[key.strip().lower()] = value.strip().strip('"')

        name = disposition_params.get("name")
        if not name:
            continue

        if val_data.endswith(b"\r\n"):
            val_data = val_data[:-2]

        if "filename" in disposition_params or "filename*" in disposition_params:
            if name == "excel_file":
                file_bytes = val_data
                fname = "file.xlsx"
                if "filename*" in disposition_params:
                    encoded_filename = disposition_params["filename*"]
                    if "''" in encoded_filename:
                        encoded_filename = encoded_filename.split("''", 1)[1]
                    fname = unquote(encoded_filename)
                elif "filename" in disposition_params:
                    fname = disposition_params["filename"]
                filename = safe_filename(fname)
        else:
            fields[name] = val_data.decode("utf-8", "ignore").strip()

    return fields, filename, file_bytes


index_html = """<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Định dạng BOM đi chợ</title>
  <style>
    :root {
      --bg: #f4f6f8;
      --panel: #ffffff;
      --text: #17202a;
      --muted: #5b6673;
      --line: #d8dee6;
      --accent: #145c52;
      --accent-strong: #0d4039;
      --soft: #e8f3f0;
      --danger: #9f2d2d;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: Arial, Helvetica, sans-serif;
      color: var(--text);
      background:
        linear-gradient(180deg, #ffffff 0, var(--bg) 260px),
        var(--bg);
    }
    main {
      width: min(920px, calc(100% - 32px));
      margin: 0 auto;
      padding: 40px 0;
    }
    .top {
      display: flex;
      align-items: flex-end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 22px;
    }
    h1 {
      margin: 0;
      font-size: 30px;
      line-height: 1.2;
      letter-spacing: 0;
    }
    .badge {
      border: 1px solid var(--line);
      background: var(--panel);
      color: var(--muted);
      padding: 9px 12px;
      border-radius: 6px;
      font-size: 14px;
      white-space: nowrap;
    }
    .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 16px 36px rgba(31, 42, 55, 0.08);
      overflow: hidden;
    }
    .form {
      padding: 26px;
      display: grid;
      gap: 20px;
    }
    .drop {
      display: grid;
      place-items: center;
      min-height: 260px;
      border: 2px dashed #aab5c2;
      border-radius: 8px;
      background: #fbfcfd;
      cursor: pointer;
      transition: border-color .18s ease, background .18s ease;
    }
    .drop:hover, .drop.is-dragover {
      border-color: var(--accent);
      background: var(--soft);
    }
    .drop input {
      position: absolute;
      opacity: 0;
      pointer-events: none;
    }
    .drop-inner {
      display: grid;
      justify-items: center;
      gap: 12px;
      text-align: center;
      padding: 22px;
    }
    .icon {
      width: 58px;
      height: 58px;
      border-radius: 8px;
      display: grid;
      place-items: center;
      background: var(--soft);
      color: var(--accent);
    }
    .icon svg {
      width: 30px;
      height: 30px;
    }
    .file-title {
      font-size: 20px;
      font-weight: 700;
    }
    .file-subtitle {
      color: var(--muted);
      font-size: 15px;
    }
    .actions {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      flex-wrap: wrap;
    }
    .status {
      min-height: 24px;
      color: var(--muted);
      font-size: 15px;
    }
    .status.error {
      color: var(--danger);
      font-weight: 700;
    }
    button {
      min-width: 190px;
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: #ffffff;
      padding: 14px 20px;
      font-size: 16px;
      font-weight: 700;
      cursor: pointer;
    }
    button:hover { background: var(--accent-strong); }
    button:disabled {
      background: #9aa7b2;
      cursor: not-allowed;
    }
    .steps {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      border-top: 1px solid var(--line);
      background: #f9fafb;
    }
    .step {
      padding: 18px 20px;
      border-right: 1px solid var(--line);
      color: var(--muted);
      font-size: 14px;
    }
    .step:last-child { border-right: 0; }
    .step strong {
      display: block;
      color: var(--text);
      font-size: 15px;
      margin-bottom: 4px;
    }
    .options-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      margin-bottom: 8px;
    }
    .control-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .control-group label {
      font-size: 12px;
      font-weight: 700;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .control-group select {
      padding: 12px;
      font-size: 14px;
      font-family: inherit;
      border: 1px solid var(--line);
      border-radius: 6px;
      background-color: var(--panel);
      color: var(--text);
      outline: none;
      cursor: pointer;
      transition: border-color .15s ease, box-shadow .15s ease;
    }
    .control-group select:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(20, 92, 82, 0.15);
    }
    @media (max-width: 720px) {
      .options-row { grid-template-columns: 1fr; }
      main { width: min(100% - 20px, 920px); padding: 24px 0; }
      .top { align-items: flex-start; flex-direction: column; }
      h1 { font-size: 25px; }
      .drop { min-height: 220px; }
      .actions { align-items: stretch; flex-direction: column; }
      button { width: 100%; }
      .steps { grid-template-columns: 1fr; }
      .step { border-right: 0; border-bottom: 1px solid var(--line); }
      .step:last-child { border-bottom: 0; }
    }
  </style>
</head>
<body>
  <main>
    <div class="top">
      <h1>Định dạng BOM đi chợ</h1>
      <div class="badge">A4 dọc/ngang · dễ in · gộp dữ liệu</div>
    </div>
    <section class="panel">
      <form class="form" action="/format" method="post" enctype="multipart/form-data">
        <label class="drop" id="drop">
          <input id="file" name="excel_file" type="file" accept=".xlsx,.xlsm" required>
          <span class="drop-inner">
            <span class="icon" aria-hidden="true">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <path d="M14 2v6h6"></path>
                <path d="M12 18v-6"></path>
                <path d="m9 15 3 3 3-3"></path>
              </svg>
            </span>
            <span class="file-title" id="fileTitle">Chọn file Excel</span>
            <span class="file-subtitle">Kéo thả file vào đây hoặc bấm để chọn</span>
          </span>
        </label>
        <div class="options-row">
          <div class="control-group">
            <label for="kitchen">Lựa chọn bếp</label>
            <select id="kitchen" name="kitchen">
              <option value="bep_tai_cho" selected>Bếp tại chỗ</option>
              <option value="bep_trung_tam">Bếp trung tâm</option>
              <option value="bep_trung_tam_2">Bếp trung tâm 2</option>
            </select>
          </div>
          <div class="control-group">
            <label for="format_mode">Bố cục xuất file</label>
            <select id="format_mode" name="format_mode">
              <option value="format1" selected>Format 1 (Bảng A4 dọc 13 cột + dòng ngày)</option>
              <option value="format2">Format 2 (Giấy đi chợ)</option>
              <option value="duyet_dinh_muc">Duyệt định mức (A4 ngang, gộp món liền kề)</option>
            </select>
          </div>
        </div>
        <div class="actions">
          <div class="status" id="status">Chưa chọn file</div>
          <button id="submit" type="submit" disabled>Xuất file đã định dạng</button>
        </div>
      </form>
      <div class="steps">
        <div class="step"><strong>1. Chọn file</strong>File BOM Excel cần in.</div>
        <div class="step"><strong>2. Xuất file</strong>Chọn Format 1, Format 2 hoặc Duyệt định mức rồi tải file.</div>
        <div class="step"><strong>3. In A4</strong>Mở file tải về rồi in theo thiết lập sẵn.</div>
      </div>
    </section>
  </main>
  <script>
    const drop = document.getElementById('drop');
    const file = document.getElementById('file');
    const title = document.getElementById('fileTitle');
    const status = document.getElementById('status');
    const submit = document.getElementById('submit');
    const form = document.querySelector('form');

    function updateFileName() {
      const selected = file.files && file.files[0];
      title.textContent = selected ? selected.name : 'Chọn file Excel';
      status.textContent = selected ? 'Sẵn sàng xuất file' : 'Chưa chọn file';
      status.classList.remove('error');
      submit.disabled = !selected;
    }

    file.addEventListener('change', updateFileName);
    drop.addEventListener('dragover', event => {
      event.preventDefault();
      drop.classList.add('is-dragover');
    });
    drop.addEventListener('dragleave', () => drop.classList.remove('is-dragover'));
    drop.addEventListener('drop', event => {
      event.preventDefault();
      drop.classList.remove('is-dragover');
      if (event.dataTransfer.files.length) {
        file.files = event.dataTransfer.files;
        updateFileName();
      }
    });
    form.addEventListener('submit', () => {
      submit.disabled = true;
      status.textContent = 'Đang xử lý file...';
      window.setTimeout(() => {
        submit.disabled = false;
        status.textContent = file.files[0] ? 'Có thể xuất lại file' : 'Chưa chọn file';
      }, 1800);
    });
  </script>
</body>
</html>
"""


def error_html(message):
    escaped_message = html.escape(message)
    return f"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Không xử lý được file</title>
  <style>
    body {{
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: Arial, Helvetica, sans-serif;
      background: #f4f6f8;
      color: #17202a;
    }}
    .box {{
      width: min(560px, calc(100% - 32px));
      background: #fff;
      border: 1px solid #d8dee6;
      border-radius: 8px;
      padding: 26px;
      box-shadow: 0 16px 36px rgba(31, 42, 55, 0.08);
    }}
    h1 {{ margin: 0 0 12px; font-size: 24px; }}
    p {{ margin: 0 0 18px; color: #5b6673; line-height: 1.5; }}
    a {{
      display: inline-block;
      background: #145c52;
      color: #fff;
      text-decoration: none;
      border-radius: 6px;
      padding: 12px 16px;
      font-weight: 700;
    }}
  </style>
</head>
<body>
  <div class="box">
    <h1>Không xử lý được file</h1>
    <p>{escaped_message}</p>
    <a href="/">Quay lại</a>
  </div>
</body>
</html>"""


class BomFormatterHandler(BaseHTTPRequestHandler):
    server_version = "BomFormatter/1.0"

    def send_text(self, status, content, content_type="text/html; charset=utf-8"):
        payload = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if self.path in {"/", "/index.html"}:
            self.send_text(200, index_html)
            return
        if self.path == "/health":
            self.send_text(200, "OK", "text/plain; charset=utf-8")
            return
        self.send_text(404, error_html("Không tìm thấy trang."))

    def do_POST(self):
        if self.path != "/format":
            self.send_text(404, error_html("Không tìm thấy trang."))
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length <= 0:
                raise ValueError("Vui lòng chọn file Excel trước khi xuất.")
            if content_length > max_upload_bytes:
                raise ValueError("File quá lớn. Giới hạn hiện tại là 50 MB.")

            body = self.rfile.read(content_length)
            fields, filename, file_bytes = parse_multipart_form(self.headers, body)

            if not file_bytes:
                raise ValueError("Vui lòng chọn file Excel trước khi xuất.")

            kitchen_opt = fields.get("kitchen", "bep_tai_cho")
            format_opt = fields.get("format_mode", "format1")

            if kitchen_opt == "bep_trung_tam":
                selected_kitchen = "Bếp trung tâm"
            elif kitchen_opt == "bep_trung_tam_2":
                selected_kitchen = "Bếp trung tâm 2"
            else:
                selected_kitchen = "Bếp tại chỗ"

            output_bytes = format_workbook_bytes(
                file_bytes, filename, date_mode="auto", format_mode=format_opt, selected_kitchen=selected_kitchen
            )
            download_name = output_filename_for_format(filename, format_opt)
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header(
                "Content-Disposition", content_disposition_filename(download_name)
            )
            self.send_header("Content-Length", str(len(output_bytes)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(output_bytes)
        except Exception as exc:
            self.send_text(400, error_html(str(exc)))

    def log_message(self, format, *args):
        print(f"{self.address_string()} - {format % args}")


def run_web_server(host="127.0.0.1", port=8100, open_browser=True):
    last_error = None
    for candidate_port in range(port, port + 20):
        try:
            server = ThreadingHTTPServer((host, candidate_port), BomFormatterHandler)
            break
        except OSError as exc:
            last_error = exc
    else:
        raise RuntimeError(f"Không mở được cổng web: {last_error}")

    url = f"http://{host}:{server.server_port}"
    browser_url = f"http://127.0.0.1:{server.server_port}" if host == "0.0.0.0" else url
    print(f"Web đã sẵn sàng: {url}")
    print("Nhấn Ctrl+C để dừng.")
    if open_browser:
        webbrowser.open(browser_url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nĐã dừng web.")
    finally:
        server.server_close()


def run_batch():
    output_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(input_dir.glob("*.xlsx"))
    files += sorted(input_dir.glob("*.xlsm"))
    files += sorted(input_dir.glob("*.csv"))
    files = [f for f in files if not f.name.startswith("~$")]

    if not files:
        print(f"Không tìm thấy file Excel/CSV trong thư mục: {input_dir}")
        return

    for file_path in files:
        file_bytes = file_path.read_bytes()
        outputs = (
            ("format1", f"{file_path.stem} - da dinh dang.xlsx", "Format 1"),
            ("format2", f"{file_path.stem} - giay di cho.xlsx", "Format 2"),
            (APPROVAL_FORMAT_MODE, f"{file_path.stem} - duyet dinh muc.xlsx", "Duyệt định mức"),
        )
        for format_mode, output_name, label in outputs:
            try:
                output_bytes = format_workbook_bytes(
                    file_bytes,
                    file_path.name,
                    date_mode="auto",
                    format_mode=format_mode,
                )
                out_path = output_dir / output_name
                with open(out_path, "wb") as f_out:
                    f_out.write(output_bytes)
                print(f"Đã gộp xong ({label}): {out_path}")
            except Exception as exc:
                print(f"Lỗi khi xử lý {label} cho {file_path.name}: {exc}")


def main():
    parser = argparse.ArgumentParser(description="Định dạng file BOM Excel để in đi chợ.")
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Xử lý tất cả file trong thư mục Excel và lưu vào thư mục Đã gộp.",
    )
    parser.add_argument("--host", default="127.0.0.1", help="Địa chỉ web local.")
    parser.add_argument("--port", type=int, default=8100, help="Cổng web local.")
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Không tự mở trình duyệt sau khi bật web.",
    )
    args = parser.parse_args()

    if args.batch:
        run_batch()
    else:
        run_web_server(args.host, args.port, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
