# main.py
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os
from typing import List, Optional, Callable

from GC335 import GC335Client
from GC337 import GC337Client


def safe_save_workbook(wb, filename):
    try:
        wb.save(filename)
        print(f"\nSaved successfully: {filename}")
    except PermissionError:
        base, ext = os.path.splitext(filename)
        out_name = f"{base}_out{ext}"
        wb.save(out_name)
        print(f"\n'{filename}' is locked (probably open in Excel).")
        print(f"Saved to: {out_name}")


# ----------------------------
# NEW: GUI-friendly helpers
# ----------------------------
def read_column_values(excel_path: str, sheet_name: str, col_letter: str, start_row: int = 2) -> List[str]:
    """
    Read values from a column (A-Z...) in a given sheet.
    Returns a list of strings (plates), skipping blanks.
    """
    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]
    col_letter = col_letter.strip().upper()

    values = []
    for cell in ws[col_letter][start_row - 1 :]:  # openpyxl is 0-indexed for slicing
        if cell.value is None:
            continue
        s = str(cell.value).strip()
        if s:
            values.append(s)
    return values


# ----------------------------
# Your existing robust helpers
# ----------------------------
def find_row_by_value(ws, value, start_row=2):
    target = str(value).strip()
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if cell.value is None:
                continue
            if str(cell.value).strip() == target:
                return cell.row
    return None


def first_empty_row_any(ws, start_row=2):
    r = start_row
    check_cols = max(ws.max_column, 10)
    while True:
        any_value = False
        for c in range(1, check_cols + 1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                any_value = True
                break
        if not any_value:
            return r
        r += 1


def first_empty_col_in_row(ws, row, start_col=1):
    for c in range(start_col, ws.max_column + 1):
        if ws.cell(row=row, column=c).value in (None, ""):
            return c
    return ws.max_column + 1


# ----------------------------
# ✅ STEP 2: callable entrypoint for your GUI
# ----------------------------
def run_job(
    excel_path: str,
    sheet_name: str,
    plate_col: str,
    vehicle_type: str,           # "335" or "337"
    headless: bool = True,
    output_path: Optional[str] = None,  # if None -> overwrite same file (safe_save_workbook handles locked file)
    progress_cb: Optional[Callable[[int, int, str], None]] = None,  # <-- NEW
):
    """
    GUI calls this.
    - Reads plates from (excel_path, sheet_name, plate_col)
    - Queries GC335/GC337
    - Writes results into the SAME sheet, into the next empty columns on each plate's row
    - Saves to output_path if provided, otherwise saves back to excel_path

    progress_cb(current, total, plate)
      - current starts at 1
    """
    excel_path = excel_path.strip().strip('"').strip("'")
    if not excel_path.lower().endswith(".xlsx"):
        raise ValueError("Please select a .xlsx file")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File not found: {excel_path}")

    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")
    ws = wb[sheet_name]

    # Read plate list from the selected column
    plates = read_column_values(excel_path, sheet_name, plate_col, start_row=2)
    total = len(plates)

    # If no plates, still report progress as 0/0 and just save/exit
    if progress_cb:
        progress_cb(0, total, "")

    vehicle_type = str(vehicle_type).strip()
    if vehicle_type == "335":
        client = GC335Client(headless=headless)
        keys_order = ["receiveDate", "dealNote", "taxreFund", "custCd", "caseNo", "msg", "dealDate", "rptDate"]
    elif vehicle_type == "337":
        client = GC337Client(headless=headless)
        keys_order = ["transId", "crtDate", "status", "refundDt"]
    else:
        raise ValueError("vehicle_type must be '335' or '337'")

    try:
        for idx, plate in enumerate(plates, start=1):
            if progress_cb:
                progress_cb(idx, total, plate)

            # find existing row that already has this plate anywhere
            row = find_row_by_value(ws, plate, start_row=2)

            # if not found, append new empty row and write plate in col A
            if row is None:
                row = first_empty_row_any(ws, start_row=2)
                ws.cell(row=row, column=1, value=plate)

            json_data = client.query_plate_first_row(plate)
            print(f"\nResult for {plate}: {json_data}")

            no_result = (
                (not isinstance(json_data, dict))
                or (not json_data)
                or (json_data.get("ok") is False)
            )

            start_col = first_empty_col_in_row(ws, row, start_col=1)

            if no_result:
                ws.cell(row=row, column=start_col, value="無結果")
                reason = json_data.get("error", "") if isinstance(json_data, dict) else ""
                ws.cell(row=row, column=start_col + 1, value=reason)
                continue

            for i, key in enumerate(keys_order):
                ws.cell(row=row, column=start_col + i, value=json_data.get(key, ""))

    finally:
        try:
            client.close()
        except Exception:
            pass

    # Save
    if output_path:
        safe_save_workbook(wb, output_path)
    else:
        safe_save_workbook(wb, excel_path)


# ----------------------------
# CLI mode still works (optional)
# ----------------------------
def main():
    # Keep your old interactive workflow if you still want it.
    excel_path = input("Enter the Excel file name (with .xlsx): ").strip().strip('"').strip("'")

    wb = load_workbook(excel_path)
    print("\nAvailable sheets:")
    for name in wb.sheetnames:
        print(f"- {name}")
    sheet_name = input("\nEnter the sheet name to write data into: ").strip()

    plate_col = input("Enter the plate column letter (A-Z): ").strip().upper()

    vehicle = input("Type (electric / gasoline): ").strip().lower()
    vehicle_type = "337" if vehicle in ("electric", "e") else "335"

    run_job(
        excel_path=excel_path,
        sheet_name=sheet_name,
        plate_col=plate_col,
        vehicle_type=vehicle_type,
        headless=True,
        output_path=None,
        progress_cb=None,
    )


if __name__ == "__main__":
    main()