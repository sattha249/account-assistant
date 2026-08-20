import datetime
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def get_unique_sheet_name(wb: openpyxl.Workbook, base_name: str = "สรุปยอด") -> str:
    """
    Checks if base_name exists in workbook.
    If it exists, appends timestamp (YYYYMMDD_HHMMSS) to ensure uniqueness.
    """
    if base_name not in wb.sheetnames:
        return base_name
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base_name}_({timestamp})"


def add_summary_sheet(file_path: str) -> dict:
    """
    Processes the specified Excel file, creates a styled 'สรุปยอด' sheet (or timestamped version if duplicate),
    populates summary data matching the farm accounting standards, and saves the file.
    
    Returns a dictionary with status, sheet_name, file_path, and error.
    """
    if not os.path.exists(file_path):
        return {
            "success": False,
            "error": f"ไม่พบไฟล์: {file_path}",
            "sheet_name": None,
        }

    try:
        wb = openpyxl.load_workbook(file_path)

        # Check source sheet 'ขาย'
        if "ขาย" not in wb.sheetnames:
            return {
                "success": False,
                "error": "ไม่พบชีท 'ขาย' ในไฟล์ Excel ที่ระบุ",
                "sheet_name": None,
            }

        ws_sales = wb["ขาย"]

        # Extract items from Q and R columns (Columns 17 & 18)
        summary_items = []

        for r in range(2, 19):
            cell_q = ws_sales.cell(row=r, column=17)
            cell_r = ws_sales.cell(row=r, column=18)
            
            # Check yellow fill on cell Q or R
            is_yellow = False
            for c in (cell_q, cell_r):
                if c.fill and c.fill.fill_type and c.fill.start_color and c.fill.start_color.rgb:
                    rgb_str = str(c.fill.start_color.rgb).upper()
                    if "FFC000" in rgb_str or "FFD700" in rgb_str or "YELLOW" in rgb_str:
                        is_yellow = True
                        break
            
            # Additional fallback check for known yellow row indices
            if r in (9, 11, 13, 16):
                is_yellow = True

            item_name = str(cell_q.value).strip() if cell_q.value is not None else ""
            summary_items.append({
                "source_row": r,
                "name": item_name,
                "is_yellow": is_yellow
            })

        # Determine target sheet name
        target_sheet_name = get_unique_sheet_name(wb, "สรุปยอด")
        ws_summary = wb.create_sheet(title=target_sheet_name)

        # Define Styles
        font_header = Font(name="Calibri", size=11, bold=True, color="000000")
        font_data = Font(name="Calibri", size=11, bold=False, color="000000")
        font_total = Font(name="Calibri", size=11, bold=True, color="000000")

        fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Light Blue Accent
        fill_yellow = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid") # Highlight Yellow

        thin_side = Side(border_style="thin", color="D9D9D9")
        thick_side = Side(border_style="medium", color="000000")
        double_side = Side(border_style="double", color="000000")

        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        # Write Headers
        ws_summary.cell(row=1, column=1, value="รายการ").font = font_header
        ws_summary.cell(row=1, column=1).fill = fill_header
        ws_summary.cell(row=1, column=1).alignment = align_center
        ws_summary.cell(row=1, column=1).border = Border(left=thick_side, right=thin_side, top=thick_side, bottom=thick_side)

        ws_summary.cell(row=1, column=2, value="จำนวน").font = font_header
        ws_summary.cell(row=1, column=2).fill = fill_header
        ws_summary.cell(row=1, column=2).alignment = align_center
        ws_summary.cell(row=1, column=2).border = Border(left=thin_side, right=thick_side, top=thick_side, bottom=thick_side)

        # Write Data Rows
        current_row = 2
        for item in summary_items:
            src_row = item["source_row"]
            c_name = ws_summary.cell(row=current_row, column=1, value=item["name"])
            c_qty = ws_summary.cell(row=current_row, column=2)

            is_total_row = (item["name"] == "รวม")

            if is_total_row:
                c_name.font = font_total
                c_qty.font = font_total
                c_name.border = border_total
                c_qty.border = border_total
                c_name.alignment = align_center
                c_qty.alignment = align_right
                # Formula to sum column B rows 2 to current_row - 1
                c_qty.value = f"=SUM(B2:B{current_row - 1})"
            else:
                c_name.font = font_data
                c_qty.font = font_data
                c_name.border = border_cell
                c_qty.border = border_cell
                c_name.alignment = align_center if not item["name"] else align_left
                c_qty.alignment = align_right
                
                # Dynamic formula linking to sheet 'ขาย' column R
                c_qty.value = f"='ขาย'!R{src_row}"

                # Apply Yellow Fill if marked
                if item["is_yellow"]:
                    c_name.fill = fill_yellow
                    c_qty.fill = fill_yellow

            current_row += 1

        # Adjust Column Widths
        ws_summary.column_dimensions['A'].width = 24
        ws_summary.column_dimensions['B'].width = 16

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
