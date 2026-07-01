from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from pathlib import Path

from .common import (
    normalize_comp, to_numeric, clean_ingredient_name,
    normalize_shift, get_dish_category, get_shift_order,
    get_cat_order, get_lixil_site_order, format_number_clean,
    get_vietnamese_weekday, extract_date_from_filename,
    parse_date_value, format_date, strip_accents,
    SHORT_NAMES
)


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
        elif any(x in c_lower for x in ["nhóm món", "nhom mon", "cơ cấu món ăn", "co cau mon an", "cơ cấu menu", "co cau menu"]):
            col_map["nhom_mon"] = idx
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
            if idx is not None and (idx - 1) < len(row_vals):
                val = str(row_vals[idx - 1]).strip() if row_vals[idx - 1] is not None else ""
                if not val:
                    val = prev_vals.get(key, "")
                row_vals[idx - 1] = val
                prev_vals[key] = val
        parsed_rows.append(row_vals)

    # Date extraction using our improved parser
    dates = []
    seen_dates = set()
    if col_ngay_idx is not None:
        for row in parsed_rows:
            if (col_ngay_idx - 1) < len(row):
                val = row[col_ngay_idx - 1]
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

    # Subtract 1 from indices for 0-indexed list access (since they are 1-indexed Excel column numbers)
    return build_compact_shopping_workbook(
        parsed_rows,
        col_ma_kh_idx - 1 if col_ma_kh_idx is not None else None,
        col_ca_idx - 1 if col_ca_idx is not None else None,
        col_mon_idx - 1 if col_mon_idx is not None else None,
        col_qty_idx - 1 if col_qty_idx is not None else None,
        col_nvl_idx - 1 if col_nvl_idx is not None else None,
        col_dinh_muc_idx - 1 if col_dinh_muc_idx is not None else None,
        col_dinh_muc_unit_idx - 1 if col_dinh_muc_unit_idx is not None else None,
        col_nhom_mon_idx - 1 if col_nhom_mon_idx is not None else None,
        col_site_an_idx - 1 if col_site_an_idx is not None else None,
        date_text,
        weekday_text,
        target_sheet_name,
    )
