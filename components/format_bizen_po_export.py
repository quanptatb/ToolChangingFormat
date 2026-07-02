import os
from copy import copy
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .common import merged_value_lookup, effective_cell_value, input_dir, output_dir

BIZEN_EXPORT_IDENTIFIER = "BIZEN - CATERING_Xuất PO ra ex"
BIZEN_EXPORT_FORMAT_MODE = "bizen_po_export"

# Headers required for detection
_REQUIRED_HEADERS_ALT_1 = {
    "Ngày",
    "Mã PO",
    "Nhà cung cấp",
    "Mã nguyên vật liệu",
    "Tên nguyên vật liệu",
    "Đơn vị tính",
    "Tổng Khối lượng đặt hàng",
    "Tổng Đơn giá",
    "Tổng thành tiền",
    "Khách hàng",
    "Ca",
    "Nơi giao hàng (Khách hàng)",
    "Giờ giao hàng"
}

_REQUIRED_HEADERS_ALT_2 = {
    "Ngày",
    "Mã PO",
    "Nhà cung cấp",
    "Mã nguyên vật liệu",
    "Tên nguyên vật liệu",
    "Đơn vị tính",
    "Sum: Khối lượng đặt hàng",
    "Sum: Đơn giá",
    "Sum: Thành tiền",
    "Khách hàng",
    "Ca",
    "Nơi giao hàng (Khách hàng)",
    "Giờ giao hàng"
}

def can_process_bizen_po_export(ws):
    headers = set()
    for c in range(1, min(ws.max_column + 1, 50)):
        val = ws.cell(1, c).value
        if val:
            headers.add(str(val).strip())
    return _REQUIRED_HEADERS_ALT_1.issubset(headers) or _REQUIRED_HEADERS_ALT_2.issubset(headers)

def _parse_currency(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("₫", "").replace("đ", "").replace("VND", "").replace("vnd", "").strip()
    if "," in text and "." in text:
        last_comma = text.rfind(",")
        last_dot = text.rfind(".")
        if last_comma > last_dot:
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", "")
    elif "." in text:
        parts = text.split(".")
        if len(parts) == 2 and len(parts[1]) == 3:
            text = text.replace(".", "")
        elif len(parts) > 2:
            text = text.replace(".", "")
    try:
        num = float(text)
        return int(num) if num == int(num) else num
    except (ValueError, OverflowError):
        return value

def _parse_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        num = float(text)
        return int(num) if num == int(num) else num
    except ValueError:
        return value

def _detect_source_columns(ws):
    header_map = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(1, c).value
        if val:
            header_map[str(val).strip()] = c

    mapping = {
        "ngay": header_map.get("Ngày"),
        "ma_po": header_map.get("Mã PO"),
        "nha_cung_cap": header_map.get("Nhà cung cấp"),
        "ma_hang": header_map.get("Mã nguyên vật liệu"),
        "dien_giai": header_map.get("Tên nguyên vật liệu"),
        "khoi_luong": header_map.get("Tổng Khối lượng đặt hàng") or header_map.get("Sum: Khối lượng đặt hàng"),
        "don_vi": header_map.get("Đơn vị tính"),
        "don_gia": header_map.get("Tổng Đơn giá") or header_map.get("Sum: Đơn giá"),
        "thanh_tien": header_map.get("Tổng thành tiền") or header_map.get("Sum: Thành tiền"),
        "khach_hang": header_map.get("Khách hàng"),
        "ca_an": header_map.get("Ca"),
        "noi_giao": header_map.get("Nơi giao hàng (Khách hàng)"),
        "gio_giao": header_map.get("Giờ giao hàng"),
    }
    
    missing = [k for k, v in mapping.items() if v is None]
    if missing:
        raise ValueError("Không tìm thấy các cột bắt buộc: " + ", ".join(missing))
    return mapping

def process_sheet_bizen_po_export(ws):
    col_map = _detect_source_columns(ws)
    lookup = merged_value_lookup(ws)

    # Đọc dữ liệu và gom nhóm theo Mã PO
    grouped_data = defaultdict(list)
    
    for r in range(2, ws.max_row + 1):
        ma_po = effective_cell_value(ws, lookup, r, col_map["ma_po"])
        if not ma_po:
            continue
            
        ngay = effective_cell_value(ws, lookup, r, col_map["ngay"])
        nha_cung_cap = effective_cell_value(ws, lookup, r, col_map["nha_cung_cap"])
        ma_hang = effective_cell_value(ws, lookup, r, col_map["ma_hang"])
        dien_giai = effective_cell_value(ws, lookup, r, col_map["dien_giai"])
        khoi_luong_raw = effective_cell_value(ws, lookup, r, col_map["khoi_luong"])
        don_vi = effective_cell_value(ws, lookup, r, col_map["don_vi"])
        don_gia_raw = effective_cell_value(ws, lookup, r, col_map["don_gia"])
        thanh_tien_raw = effective_cell_value(ws, lookup, r, col_map["thanh_tien"])
        khach_hang = effective_cell_value(ws, lookup, r, col_map["khach_hang"])
        ca_an = effective_cell_value(ws, lookup, r, col_map["ca_an"])
        noi_giao = effective_cell_value(ws, lookup, r, col_map["noi_giao"])
        gio_giao = effective_cell_value(ws, lookup, r, col_map["gio_giao"])
        
        khoi_luong = _parse_number(khoi_luong_raw)
        don_gia = _parse_currency(don_gia_raw)
        thanh_tien = _parse_currency(thanh_tien_raw)
        
        grouped_data[ma_po].append({
            "ngay": ngay,
            "ma_po": ma_po,
            "nha_cung_cap": nha_cung_cap,
            "ma_hang": ma_hang,
            "dien_giai": dien_giai,
            "khoi_luong": khoi_luong,
            "don_vi": don_vi,
            "don_gia": don_gia,
            "thanh_tien": thanh_tien,
            "khach_hang": khach_hang,
            "ca_an": ca_an,
            "noi_giao": noi_giao,
            "gio_giao": gio_giao
        })
        
    if not grouped_data:
        raise ValueError("Không có dữ liệu PO.")

    template_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "file định dạng fotmat chuyển đổi chính.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Không tìm thấy file mẫu tại {template_path}")

    # Output workbook
    out_wb = Workbook()
    # Xoá sheet mặc định
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    thin_side = Side(style="thin", color="000000")
    body_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")
    
    number_format_thousands = '#,##0'
    number_format_decimal = '#,##0.00'
    
    _CENTER_COLS = {1, 4, 8} # Mã hàng, Đơn vị, Ca ăn
    _RIGHT_COLS = {3, 5, 6} # Khối lượng, Đơn giá, Thành tiền

    for ma_po, rows in grouped_data.items():
        # Sắp xếp theo khách hàng (A -> Z)
        rows.sort(key=lambda x: str(x["khach_hang"] or ""))
        
        # Load lại template cho mỗi sheet
        tpl_wb = load_workbook(template_path)
        tpl_ws = tpl_wb.active
        
        out_ws = out_wb.create_sheet(title=str(ma_po)[:31]) # Giới hạn tên sheet 31 ký tự
        
        # Copy từ tpl_ws sang out_ws
        # Copy cell values, styles, merged cells, row/col dimensions
        for row in tpl_ws.iter_rows():
            for cell in row:
                new_cell = out_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)
                    
                    # Giữ nguyên màu/viền cho header (dòng 1-5), còn từ dòng 6 trở đi bỏ trống để dùng style của Excel Table (Ctrl+T)
                    if cell.row < 6:
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
        
        for merged_cell_range in tpl_ws.merged_cells.ranges:
            out_ws.merge_cells(str(merged_cell_range))
            
        for col_letter, col_dim in tpl_ws.column_dimensions.items():
            out_ws.column_dimensions[col_letter] = copy(col_dim)
        for row_index, row_dim in tpl_ws.row_dimensions.items():
            out_ws.row_dimensions[row_index] = copy(row_dim)
            
        # Ghi các thông tin chung
        first_row = rows[0]
        
        # Format the date properly
        ngay_val = first_row["ngay"]
        cell_ngay = out_ws.cell(row=1, column=5, value=ngay_val)
        cell_ngay.font = Font(bold=True, size=12)
        import datetime
        if isinstance(ngay_val, datetime.datetime):
            cell_ngay.number_format = 'DD/MM/YYYY'
            
        cell_po = out_ws.cell(row=2, column=5, value=first_row["ma_po"])
        cell_po.font = Font(bold=True, size=12)
        
        cell_ncc = out_ws.cell(row=3, column=5, value=first_row["nha_cung_cap"])
        cell_ncc.font = Font(bold=True, size=12)
        
        # Đóng khung (add borders) cho phần thông tin chung
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        # Apply border to A1:E3
        for r in range(1, 4):
            for c in range(1, 6):
                c_obj = out_ws.cell(row=r, column=c)
                c_obj.border = thin_border
        # Ghi dữ liệu bảng
        start_row = 6
        for i, r_data in enumerate(rows):
            r_idx = start_row + i
            values = [
                r_data["ma_hang"],
                r_data["dien_giai"],
                r_data["khoi_luong"],
                r_data["don_vi"],
                r_data["don_gia"],
                r_data["thanh_tien"],
                r_data["khach_hang"],
                r_data["ca_an"],
                r_data["noi_giao"],
                r_data["gio_giao"]
            ]
            for col_idx, val in enumerate(values, start=1):
                cell = out_ws.cell(row=r_idx, column=col_idx, value=val)
                cell.border = body_border
                cell.alignment = align_center
                if isinstance(val, (int, float)):
                    if isinstance(val, float) and not val.is_integer():
                        cell.number_format = number_format_decimal
                    else:
                        cell.number_format = number_format_thousands
        
        # Auto-fit column widths (starting row 1 to include general info headers)
        for col_idx in range(1, 11):
            max_len = 0
            for row_idx in range(1, start_row + len(rows)):
                cell_val = out_ws.cell(row=row_idx, column=col_idx).value
                if cell_val is not None:
                    if isinstance(cell_val, (int, float)):
                        display_len = len("{:,.0f}".format(cell_val))
                    else:
                        display_len = len(str(cell_val))
                    if display_len > max_len:
                        max_len = display_len
            
            # Cột B, I cần rộng hơn
            adjusted_width = max_len + 2
            if col_idx in [2, 9]:
                adjusted_width = max(20, min(adjusted_width, 60))
            else:
                adjusted_width = max(10, min(adjusted_width, 40))
            
            # Chỉ set nếu nó lớn hơn width hiện tại
            current_width = out_ws.column_dimensions[get_column_letter(col_idx)].width
            if current_width is None or adjusted_width > current_width:
                out_ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

        # Auto filter & Excel Table
        last_col_letter = get_column_letter(10)
        last_row = start_row + len(rows) - 1
        
        # Add Excel Table (Ctrl+T equivalent)
        from openpyxl.worksheet.table import Table, TableStyleInfo
        safe_name = str(ma_po)[:10].replace("-", "_").replace(" ", "")
        tab = Table(displayName=f"PO_{safe_name}", ref=f"A5:{last_col_letter}{last_row}")
        
        # Add a default table style with row stripes
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        out_ws.add_table(tab)
        
        # Freeze panes
        out_ws.freeze_panes = "A6"
        
    return out_wb
