import datetime
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils import normalize_category_name, category_key, is_same_category


def get_unique_sheet_name(wb: openpyxl.Workbook, base_name: str = "สรุปยอด") -> str:
    """
    Checks if base_name exists in workbook.
    If it exists, appends timestamp (YYYYMMDD_HHMMSS) to ensure uniqueness.
    """
    if base_name not in wb.sheetnames:
        return base_name
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_({timestamp})"


def is_orange_fill(cell) -> bool:
    """Returns True if the cell has Orange/Amber background fill."""
    if cell and cell.fill and cell.fill.fill_type and cell.fill.start_color and cell.fill.start_color.rgb:
        rgb_str = str(cell.fill.start_color.rgb).upper()
        if "FFC000" in rgb_str or "FFD700" in rgb_str or "ORANGE" in rgb_str:
            return True
    return False


def get_skipped_orange_rows(ws_sales: openpyxl.worksheet.worksheet.Worksheet, max_r: int) -> set:
    """
    Scans sheet 'ขาย' for rows with Orange fill.
    For each orange row r, checks row r-1 and row r+1.
    If Date(r) == Date(adj) AND Qty(r) == Qty(adj), marks row r to be skipped.
    """
    skipped_rows = set()
    
    for r in range(2, max_r + 1):
        cell_b = ws_sales.cell(row=r, column=2)
        cell_d = ws_sales.cell(row=r, column=4)
        cell_e = ws_sales.cell(row=r, column=5)
        
        is_orange = any(is_orange_fill(c) for c in (cell_b, cell_d, cell_e))
        
        if is_orange:
            d_curr = ws_sales.cell(row=r, column=2).value
            q_curr = ws_sales.cell(row=r, column=4).value
            
            # Compare with prev row (r-1)
            match_prev = False
            if r > 2:
                d_prev = ws_sales.cell(row=r - 1, column=2).value
                q_prev = ws_sales.cell(row=r - 1, column=4).value
                if d_curr == d_prev and q_curr == q_prev and q_curr is not None:
                    match_prev = True
            
            # Compare with next row (r+1)
            match_next = False
            if r < max_r:
                d_next = ws_sales.cell(row=r + 1, column=2).value
                q_next = ws_sales.cell(row=r + 1, column=4).value
                if d_curr == d_next and q_curr == q_next and q_curr is not None:
                    match_next = True
            
            if match_prev or match_next:
                skipped_rows.add(r)
                
    return skipped_rows


def detect_weeks_and_categories(ws_sales: openpyxl.worksheet.worksheet.Worksheet) -> tuple:
    """
    Detects week blocks (WK.xx) and date ranges in sheet 'ขาย'.
    Also extracts all distinct product categories across the sheet.
    
    Returns:
        (weeks: list of dicts, categories: list of dicts)
    """
    max_r = ws_sales.max_row
    weeks = []
    start_r = 2

    # Step 1: Scan for week separators (Column B starting with 'WK.')
    for r in range(2, max_r + 1):
        cell_b = ws_sales.cell(row=r, column=2).value
        cell_c = ws_sales.cell(row=r, column=3).value

        if cell_b and str(cell_b).strip().startswith("WK."):
            d_str = str(cell_c).strip() if cell_c else ""
            if d_str.startswith("วันที่ "):
                d_str = d_str.replace("วันที่ ", "")

            weeks.append({
                "name": str(cell_b).strip(),
                "date_range": d_str,
                "start_row": start_r,
                "end_row": r
            })
            start_r = r + 1

    # If remaining rows exist after the last WK separator
    if start_r <= max_r:
        dates = []
        for r in range(start_r, max_r + 1):
            v = ws_sales.cell(row=r, column=2).value
            if hasattr(v, "strftime"):
                y = v.year if v.year < 2500 else v.year - 543
                dates.append(f"{v.day}/{v.month}/{str(y)[-2:]}")
        
        d_range = f"{dates[0]} - {dates[-1]}" if len(dates) > 1 else (dates[0] if dates else "ท้ายเดือน")
        
        has_content = any(ws_sales.cell(row=r, column=5).value for r in range(start_r, max_r + 1))
        if has_content:
            wk_num = len(weeks) + 32 if weeks else 1
            weeks.append({
                "name": f"WK.{wk_num}" if weeks else "สรุปทั้งเดือน",
                "date_range": d_range,
                "start_row": start_r,
                "end_row": max_r
            })

    if not weeks:
        weeks = [{
            "name": "สรุปทั้งเดือน",
            "date_range": "ทั้งหมด",
            "start_row": 2,
            "end_row": max_r
        }]

    # Step 2: Calculate date ranges for weeks if empty
    for wk in weeks:
        if not wk["date_range"]:
            dates = []
            for r in range(wk["start_row"], wk["end_row"] + 1):
                v = ws_sales.cell(row=r, column=2).value
                if hasattr(v, "strftime"):
                    y = v.year if v.year < 2500 else v.year - 543
                    dates.append(f"{v.day}/{v.month}/{str(y)[-2:]}")
            if dates:
                wk["date_range"] = f"{dates[0]} - {dates[-1]}" if len(dates) > 1 else dates[0]
            else:
                wk["date_range"] = "-"

    # Step 3: Extract distinct categories from Column E (Col 5)
    categories = []
    seen_keys = set()

    for r in range(2, max_r + 1):
        c_val = ws_sales.cell(row=r, column=5).value # Column E: ประเภท
        clean_name = normalize_category_name(c_val)
        c_key = category_key(clean_name)

        if clean_name and c_key not in seen_keys and c_key not in ("รวม", "ประเภท", "วันที่", "ชื่อลูกค้า"):
            seen_keys.add(c_key)
            categories.append({
                "name": clean_name,
                "key": c_key
            })

    return weeks, categories


def add_summary_sheet(file_path: str) -> dict:
    """
    Processes the specified Excel file, creates a styled multi-column weekly matrix 'สรุปยอด' sheet,
    linking formulas dynamically per week block and date range.
    
    Includes Orange Fill Duplicate Exclusion Rule:
    Skips rows filled with Orange if adjacent row (prev or next) has the same Date and Quantity.
    """
    if not os.path.exists(file_path):
        return {
            "success": False,
            "error": f"ไม่พบไฟล์: {file_path}",
            "sheet_name": None,
        }

    try:
        wb = openpyxl.load_workbook(file_path)

        if "ขาย" not in wb.sheetnames:
            return {
                "success": False,
                "error": "ไม่พบชีท 'ขาย' ในไฟล์ Excel ที่ระบุ",
                "sheet_name": None,
            }

        ws_sales = wb["ขาย"]
        max_r = ws_sales.max_row

        # Detect skipped orange duplicate rows
        skipped_orange_rows = get_skipped_orange_rows(ws_sales, max_r)

        # Detect weeks and categories dynamically
        weeks, categories = detect_weeks_and_categories(ws_sales)

        if not categories:
            return {
                "success": False,
                "error": "ไม่พบประเภทสินค้าในคอลัมน์ E ของชีท 'ขาย'",
                "sheet_name": None
            }

        # Determine target sheet name
        target_sheet_name = get_unique_sheet_name(wb, "สรุปยอด")
        ws_summary = wb.create_sheet(title=target_sheet_name)

        # Styles
        font_header_1 = Font(name="Calibri", size=11, bold=True, color="000000")
        font_header_2 = Font(name="Calibri", size=10, bold=False, italic=True, color="333333")
        font_data = Font(name="Calibri", size=11, bold=False, color="000000")
        font_total = Font(name="Calibri", size=11, bold=True, color="000000")

        fill_header_1 = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Light Blue Accent
        fill_header_2 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # Light Gray Accent

        thin_side = Side(border_style="thin", color="D9D9D9")
        thick_side = Side(border_style="medium", color="000000")
        double_side = Side(border_style="double", color="000000")

        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        num_format = "#,##0"

        # --- Row 1: Header Row 1 (Week Names) ---
        c_item_1 = ws_summary.cell(row=1, column=1, value="รายการ")
        c_item_1.font = font_header_1
        c_item_1.fill = fill_header_1
        c_item_1.alignment = align_center
        c_item_1.border = Border(left=thick_side, right=thin_side, top=thick_side, bottom=thin_side)

        col_idx = 2
        for wk in weeks:
            cell = ws_summary.cell(row=1, column=col_idx, value=wk["name"])
            cell.font = font_header_1
            cell.fill = fill_header_1
            cell.alignment = align_center
            cell.border = Border(left=thin_side, right=thin_side, top=thick_side, bottom=thin_side)
            col_idx += 1

        cell_total_hdr1 = ws_summary.cell(row=1, column=col_idx, value="รวมทั้งเดือน")
        cell_total_hdr1.font = font_header_1
        cell_total_hdr1.fill = fill_header_1
        cell_total_hdr1.alignment = align_center
        cell_total_hdr1.border = Border(left=thin_side, right=thick_side, top=thick_side, bottom=thin_side)

        # --- Row 2: Header Row 2 (Date Ranges) ---
        c_item_2 = ws_summary.cell(row=2, column=1, value="ช่วงวันที่")
        c_item_2.font = font_header_2
        c_item_2.fill = fill_header_2
        c_item_2.alignment = align_center
        c_item_2.border = Border(left=thick_side, right=thin_side, top=thin_side, bottom=thick_side)

        col_idx = 2
        for wk in weeks:
            cell = ws_summary.cell(row=2, column=col_idx, value=wk["date_range"])
            cell.font = font_header_2
            cell.fill = fill_header_2
            cell.alignment = align_center
            cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_side)
            col_idx += 1

        cell_total_hdr2 = ws_summary.cell(row=2, column=col_idx, value="")
        cell_total_hdr2.font = font_header_2
        cell_total_hdr2.fill = fill_header_2
        cell_total_hdr2.alignment = align_center
        cell_total_hdr2.border = Border(left=thin_side, right=thick_side, top=thin_side, bottom=thick_side)

        max_name_len = 10

        # --- Data Rows (Rows 3..M+2) ---
        current_row = 3
        for cat in categories:
            c_name = ws_summary.cell(row=current_row, column=1, value=cat["name"])
            c_name.font = font_data
            c_name.alignment = align_left
            c_name.border = border_cell

            max_name_len = max(max_name_len, len(cat["name"]))

            col_idx = 2
            for wk in weeks:
                start_r = wk["start_row"]
                end_r = wk["end_row"]
                
                # Collect valid transaction row references for this category and week block
                valid_cell_refs = []
                for r in range(start_r, end_r + 1):
                    if r in skipped_orange_rows:
                        continue
                    c_val = ws_sales.cell(row=r, column=5).value
                    clean_c = normalize_category_name(c_val)
                    if is_same_category(clean_c, cat["name"]):
                        valid_cell_refs.append(f"'ขาย'!D{r}")

                c_val = ws_summary.cell(row=current_row, column=col_idx)
                c_val.font = font_data
                c_val.alignment = align_right
                c_val.number_format = num_format
                c_val.border = border_cell

                if not valid_cell_refs:
                    c_val.value = 0
                elif len(valid_cell_refs) == 1:
                    c_val.value = f"={valid_cell_refs[0]}"
                else:
                    c_val.value = f"=SUM({', '.join(valid_cell_refs)})"

                col_idx += 1

            # Total column for row (Sum of week columns B..last_week_col)
            last_wk_col_letter = get_column_letter(col_idx - 1)
            c_row_total = ws_summary.cell(row=current_row, column=col_idx)
            c_row_total.font = font_total
            c_row_total.alignment = align_right
            c_row_total.number_format = num_format
            c_row_total.border = border_cell
            c_row_total.value = f"=SUM(B{current_row}:{last_wk_col_letter}{current_row})"

            current_row += 1

        # --- Bottom Total Row (รวม) ---
        c_tot_name = ws_summary.cell(row=current_row, column=1, value="รวม")
        c_tot_name.font = font_total
        c_tot_name.alignment = align_center
        c_tot_name.border = border_total

        last_data_row = current_row - 1
        col_idx = 2
        for wk in weeks:
            col_letter = get_column_letter(col_idx)
            c_col_tot = ws_summary.cell(row=current_row, column=col_idx)
            c_col_tot.font = font_total
            c_col_tot.alignment = align_right
            c_col_tot.number_format = num_format
            c_col_tot.border = border_total
            c_col_tot.value = f"=SUM({col_letter}3:{col_letter}{last_data_row})"
            col_idx += 1

        # Total of all totals (bottom-right cell)
        total_col_letter = get_column_letter(col_idx)
        c_grand_tot = ws_summary.cell(row=current_row, column=col_idx)
        c_grand_tot.font = font_total
        c_grand_tot.alignment = align_right
        c_grand_tot.number_format = num_format
        c_grand_tot.border = border_total
        c_grand_tot.value = f"=SUM({total_col_letter}3:{total_col_letter}{last_data_row})"

        # Column Widths
        ws_summary.column_dimensions['A'].width = max(24, max_name_len + 6)
        for c in range(2, col_idx + 1):
            c_let = get_column_letter(c)
            ws_summary.column_dimensions[c_let].width = 18

        # Save workbook
        wb.save(file_path)

        return {
            "success": True,
            "sheet_name": target_sheet_name,
            "file_path": file_path,
            "error": None
        }

    except Exception as e:
        return {
            "success": False,
            "error": f"เกิดข้อผิดพลาดในการสร้างชีทสรุปยอด: {str(e)}",
            "sheet_name": None
        }
