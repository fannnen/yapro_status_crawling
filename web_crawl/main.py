# main.py
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os

import excel_read
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


def prompt_excel_file() -> tuple:
    while True:
        filename = input("Enter the Excel file name (with .xlsx): ").strip().strip('"').strip("'")

        if not filename:
            print("File name cannot be empty.\n")
            continue
        if not filename.lower().endswith(".xlsx"):
            print("Please enter a .xlsx file.\n")
            continue
        if not os.path.exists(filename):
            print("File not found. Check the name/path and try again.\n")
            continue

        try:
            wb = load_workbook(filename)
            return wb, filename
        except InvalidFileException:
            print("Invalid Excel file. Make sure it is a real .xlsx file.\n")
        except PermissionError:
            print("Permission denied. Close the file in Excel and try again.\n")
        except Exception as e:
            print(f"Could not open file: {e}\n")


def select_sheet(workbook):
    print("\nAvailable sheets:")
    for name in workbook.sheetnames:
        print(f"- {name}")

    while True:
        sheet_name = input("\nEnter the sheet name to write data into: ").strip()
        if not sheet_name:
            print("Sheet name cannot be empty.")
            continue
        if sheet_name in workbook.sheetnames:
            return workbook[sheet_name]
        print("Sheet not found. Please choose from the list above.")


def get_vehicle_type():
    while True:
        user_choice = input("What type of car would you like to search (electric / gasoline)? ")
        if not user_choice:
            print("Please enter something.")
            continue

        choice = user_choice.strip().lower()
        if choice in ("electric", "e"):
            return "337"
        elif choice in ("gasoline", "g", "gas", "petrol"):
            return "335"
        else:
            print("Invalid input. Please enter 'electric' (e) or 'gasoline' (g).")


# ----------------------------
# NEW: robust sheet helpers
# ----------------------------

def find_row_by_value(ws, value, start_row=2):
    """
    Find a row that contains the exact plate value anywhere in the sheet (from start_row).
    Returns row index if found, else None.
    """
    target = str(value).strip()
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if cell.value is None:
                continue
            if str(cell.value).strip() == target:
                return cell.row
    return None


def first_empty_row_any(ws, start_row=2):
    """
    Returns the first row that is completely empty (across columns 1..ws.max_column).
    If the sheet is narrow/empty, this still works and will append cleanly.
    """
    r = start_row
    # Ensure we check at least a few columns even when ws.max_column is 1
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
    """
    Find the first empty column in a specific row, scanning from start_col.
    If no empty column exists within current max_column, returns max_column + 1.
    """
    for c in range(start_col, ws.max_column + 1):
        if ws.cell(row=row, column=c).value in (None, ""):
            return c
    return ws.max_column + 1


def main():
    data = excel_read.read_data()

    wb, filename = prompt_excel_file()
    ws = select_sheet(wb)

    vehicle_type = get_vehicle_type()
    if vehicle_type == "335":
        client = GC335Client(headless=False)
        keys_order = [
            "receiveDate", "dealNote", "taxreFund", "custCd",
            "caseNo", "msg", "dealDate", "rptDate"
        ]
    else:
        client = GC337Client(headless=False)
        keys_order = ["transId", "crtDate", "status", "refundDt"]

    try:
        for item in data:
            if item is None or str(item).strip() == "":
                continue

            plate = str(item).strip()

            # ✅ Row = the row that already contains this plate (any column)
            row = find_row_by_value(ws, plate, start_row=2)

            # ✅ If plate not found anywhere, append a new empty row
            if row is None:
                row = first_empty_row_any(ws, start_row=2)
                # Put plate somewhere predictable so it can be found next run
                ws.cell(row=row, column=1, value=plate)

            json_data = client.query_plate_first_row(plate)
            print(f"\nResult for {plate}: {json_data}")

            no_result = (
                (not isinstance(json_data, dict)) or
                (not json_data) or
                (json_data.get("ok") is False)
            )

            # ✅ Start writing at the next empty column in THIS row
            start_col = first_empty_col_in_row(ws, row, start_col=1)

            if no_result:
                ws.cell(row=row, column=start_col, value="無結果")
                reason = ""
                if isinstance(json_data, dict):
                    reason = json_data.get("error", "")
                ws.cell(row=row, column=start_col + 1, value=reason)
                continue

            for i, key in enumerate(keys_order):
                ws.cell(row=row, column=start_col + i, value=json_data.get(key, ""))

    finally:
        client.close()

    safe_save_workbook(wb, filename)


if __name__ == "__main__":
    main()
