"""
Format Bizen Catering – chuyển đổi file PO Lưới từ BIZEN - CATERING sang bảng
dữ liệu gọn gàng với 9 cột:

  Mã hàng | Tên nhà cung cấp | Diễn giải | Đơn vị | Số lượng | Đơn giá | Thành tiền | Khách hàng | Ca ăn

File nguồn được nhận diện khi tên chứa chuỗi
"BIZEN - CATERING_Đặt hàng (PO)_Lưới".
"""

import re
from copy import copy
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, numbers
from openpyxl.utils import get_column_letter

from .common import merged_value_lookup, effective_cell_value


# ---------------------------------------------------------------------------
# Tên file nhận diện
# ---------------------------------------------------------------------------
BIZEN_IDENTIFIER = "BIZEN - CATERING_Đặt hàng (PO)_Lưới"
BIZEN_FORMAT_MODE = "bizen_po"

# ---------------------------------------------------------------------------
# Mapping cột nguồn (1-indexed) theo header row 1 của file Lưới
# ---------------------------------------------------------------------------
_SRC_MA_DAT_HANG = 1       # Mã đặt hàng
_SRC_TEN_NVL = 7            # Tên nguyên vật liệu
_SRC_KHOI_LUONG = 9         # Khối lượng đặt hàng  (= Số lượng)
_SRC_DON_VI = 10            # Đơn vị tính
_SRC_NHA_CUNG_CAP = 11      # Nhà cung cấp
_SRC_DON_GIA = 12           # Đơn giá
_SRC_THANH_TIEN = 15        # Thành tiền
_SRC_NOI_GIAO = 16          # Nơi giao hàng (Khách hàng)

# ---------------------------------------------------------------------------
# Header đầu ra
# ---------------------------------------------------------------------------
OUTPUT_HEADERS = [
    "Mã hàng",
    "Tên nhà cung cấp",
    "Diễn giải",
    "Đơn vị",
    "Số lượng",
    "Đơn giá",
    "Thành tiền",
    "Khách hàng",
    "Ca ăn",
]

# Column indices trong output (1-indexed)
_OUT_MA_HANG = 1
_OUT_NCC = 2
_OUT_DIEN_GIAI = 3
_OUT_DON_VI = 4
_OUT_SO_LUONG = 5
_OUT_DON_GIA = 6
_OUT_THANH_TIEN = 7
_OUT_KHACH_HANG = 8
_OUT_CA_AN = 9


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _parse_currency(value):
    """Loại bỏ ký hiệu tiền tệ (₫, đ, VND …) và chuyển sang số.

    Ví dụ:
        '₫65,000'   → 65000
        'đ65.000'   → 65000
        '65.000đ'   → 65000
        '₫1,100'    → 1100
        65000        → 65000   (giữ nguyên nếu đã là số)
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    if not text:
        return None

    # Loại bỏ các ký hiệu tiền tệ
    text = text.replace("₫", "").replace("đ", "").replace("VND", "").replace("vnd", "").strip()

    # Xử lý dấu phân cách:
    # Trường hợp '1,100' hoặc '65,000' (dấu phẩy là separator hàng nghìn)
    # Trường hợp '65.000' (dấu chấm là separator hàng nghìn kiểu VN)
    #
    # Heuristic: nếu có dấu phẩy, coi dấu phẩy là separator nghìn
    #            nếu chỉ có dấu chấm, kiểm tra xem có phải separator nghìn hay thập phân
    if "," in text and "." in text:
        # Both separators: e.g. '1,234.56' or '1.234,56'
        # Detect which is thousands separator
        last_comma = text.rfind(",")
        last_dot = text.rfind(".")
        if last_comma > last_dot:
            # European style: 1.234,56 → remove dots, comma→dot
            text = text.replace(".", "").replace(",", ".")
        else:
            # US style: 1,234.56 → remove commas
            text = text.replace(",", "")
    elif "," in text:
        # Only commas: '1,100' → thousands separator
        text = text.replace(",", "")
    elif "." in text:
        # Only dots: check if thousands separator
        # If there's exactly one dot and 3 digits after → thousands (VN style)
        parts = text.split(".")
        if len(parts) == 2 and len(parts[1]) == 3:
            text = text.replace(".", "")
        elif len(parts) > 2:
            # Multiple dots like '1.234.567' → thousands
            text = text.replace(".", "")
        # else: single dot with != 3 decimals → treat as decimal

    try:
        num = float(text)
        return int(num) if num == int(num) else num
    except (ValueError, OverflowError):
        return value  # trả về nguyên nếu không parse được


def _parse_number(value):
    """Parse giá trị số (Số lượng, Đơn giá dạng number thuần)."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if not text:
        return None
    # Thay dấu phẩy thành chấm cho an toàn
    text = text.replace(",", ".")
    try:
        num = float(text)
        return int(num) if num == int(num) else num
    except ValueError:
        return value


def _detect_source_columns(ws):
    """Tự động phát hiện cột dựa theo header row 1.

    Trả về dict mapping tên logic → col index (1-indexed).
    Nếu không tìm thấy đủ cột bắt buộc thì raise ValueError.
    """
    header_map = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val:
            header_map[str(val).strip()] = c

    mapping = {
        "ma_dat_hang": header_map.get("Mã đặt hàng"),
        "ten_nvl": header_map.get("Tên nguyên vật liệu"),
        "khoi_luong": header_map.get("Khối lượng đặt hàng"),
        "don_vi": header_map.get("Đơn vị tính"),
        "nha_cung_cap": header_map.get("Nhà cung cấp"),
        "don_gia": header_map.get("Đơn giá"),
        "thanh_tien": header_map.get("Thành tiền"),
        "noi_giao": header_map.get("Nơi giao hàng (Khách hàng)"),
        "ca_an": header_map.get("Ca ăn"),  # có thể None
    }

    required = ["ma_dat_hang", "ten_nvl", "don_vi", "nha_cung_cap", "don_gia", "thanh_tien"]
    missing = [k for k in required if mapping[k] is None]
    if missing:
        raise ValueError(
            "Không tìm thấy các cột bắt buộc trong file nguồn: "
            + ", ".join(missing)
        )
    return mapping


# ---------------------------------------------------------------------------
# Main processor
# ---------------------------------------------------------------------------

def process_sheet_bizen_po(ws):
    """Đọc worksheet nguồn (BIZEN PO Lưới) và trả về Workbook mới đã định dạng."""

    col_map = _detect_source_columns(ws)
    lookup = merged_value_lookup(ws)

    # Thu thập dữ liệu
    rows_data = []
    for r in range(2, ws.max_row + 1):
        ma_hang = effective_cell_value(ws, lookup, r, col_map["ma_dat_hang"])
        ten_nvl = effective_cell_value(ws, lookup, r, col_map["ten_nvl"])
        don_vi = effective_cell_value(ws, lookup, r, col_map["don_vi"])
        nha_cung_cap = effective_cell_value(ws, lookup, r, col_map["nha_cung_cap"])
        don_gia_raw = effective_cell_value(ws, lookup, r, col_map["don_gia"])
        thanh_tien_raw = effective_cell_value(ws, lookup, r, col_map["thanh_tien"])
        noi_giao = effective_cell_value(ws, lookup, r, col_map["noi_giao"])

        khoi_luong_raw = effective_cell_value(ws, lookup, r, col_map["khoi_luong"])
        ca_an_col = col_map.get("ca_an")
        ca_an = effective_cell_value(ws, lookup, r, ca_an_col) if ca_an_col else None

        # Bỏ qua dòng hoàn toàn trống
        if ma_hang is None and ten_nvl is None and don_gia_raw is None:
            continue

        so_luong = _parse_number(khoi_luong_raw)
        don_gia = _parse_currency(don_gia_raw)
        thanh_tien = _parse_currency(thanh_tien_raw)

        rows_data.append({
            "ma_hang": ma_hang,
            "nha_cung_cap": nha_cung_cap,
            "dien_giai": ten_nvl,
            "don_vi": don_vi,
            "so_luong": so_luong,
            "don_gia": don_gia,
            "thanh_tien": thanh_tien,
            "khach_hang": noi_giao,
            "ca_an": ca_an,
        })

    if not rows_data:
        raise ValueError("Không tìm thấy dữ liệu trong file nguồn.")

    # -----------------------------------------------------------------------
    # Tạo workbook đầu ra
    # -----------------------------------------------------------------------
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Đặt hàng"

    # ---- Styles ----------------------------------------------------------
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin_side = Side(style="thin", color="B0B0B0")
    header_border = Border(
        left=Side(style="thin", color="305496"),
        right=Side(style="thin", color="305496"),
        top=Side(style="thin", color="305496"),
        bottom=Side(style="thin", color="305496"),
    )
    body_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    number_format_thousands = '#,##0'
    number_format_decimal = '#,##0.00'

    # ---- Header row ------------------------------------------------------
    for col_idx, header_text in enumerate(OUTPUT_HEADERS, start=1):
        cell = out_ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = header_align

    # ---- Data rows -------------------------------------------------------
    for row_idx, data in enumerate(rows_data, start=2):
        values = [
            data["ma_hang"],
            data["nha_cung_cap"],
            data["dien_giai"],
            data["don_vi"],
            data["so_luong"],
            data["don_gia"],
            data["thanh_tien"],
            data["khach_hang"],
            data["ca_an"],
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = out_ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = body_font
            cell.border = body_border

            # Alignment theo yêu cầu
            if col_idx in (_OUT_MA_HANG, _OUT_DON_VI, _OUT_KHACH_HANG, _OUT_CA_AN):
                cell.alignment = align_center
            elif col_idx in (_OUT_SO_LUONG, _OUT_DON_GIA, _OUT_THANH_TIEN):
                cell.alignment = align_right
                # Number format
                if isinstance(val, (int, float)):
                    if isinstance(val, float) and not val.is_integer():
                        cell.number_format = number_format_decimal
                    else:
                        cell.number_format = number_format_thousands
            else:
                cell.alignment = align_left

    # ---- Auto-fit column widths ------------------------------------------
    for col_idx in range(1, len(OUTPUT_HEADERS) + 1):
        max_len = len(OUTPUT_HEADERS[col_idx - 1])
        for row_idx in range(2, len(rows_data) + 2):
            cell_val = out_ws.cell(row=row_idx, column=col_idx).value
            if cell_val is not None:
                # Cho số có format → ước lượng chiều rộng
                if isinstance(cell_val, (int, float)):
                    display_len = len("{:,.0f}".format(cell_val))
                else:
                    display_len = len(str(cell_val))
                if display_len > max_len:
                    max_len = display_len
        # Thêm padding
        adjusted_width = max_len + 4
        # Giới hạn tối thiểu và tối đa
        adjusted_width = max(10, min(adjusted_width, 45))
        out_ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    # ---- Auto filter trên header row ------------------------------------
    last_col_letter = get_column_letter(len(OUTPUT_HEADERS))
    last_row = len(rows_data) + 1
    out_ws.auto_filter.ref = "A1:{}{}".format(last_col_letter, last_row)

    # ---- Freeze header row -----------------------------------------------
    out_ws.freeze_panes = "A2"

    # ---- Alternating row colors (subtle) ---------------------------------
    even_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    for row_idx in range(2, last_row + 1):
        if row_idx % 2 == 0:
            for col_idx in range(1, len(OUTPUT_HEADERS) + 1):
                out_ws.cell(row=row_idx, column=col_idx).fill = even_fill

    return out_wb
