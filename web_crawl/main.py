from openpyxl import load_workbook
import os
import shutil
from typing import List, Optional, Callable

from GC335 import GC335Client
from GC337 import GC337Client


PLATE_COL = "D"


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


def create_backup_copy(file_path: str) -> str:
    base, ext = os.path.splitext(file_path)
    copy_path = f"{base}_copy{ext}"

    counter = 1
    while os.path.exists(copy_path):
        copy_path = f"{base}_copy{counter}{ext}"
        counter += 1

    shutil.copy2(file_path, copy_path)
    print(f"Backup created: {copy_path}")
    return copy_path


def read_column_values(excel_path: str, sheet_name: str, start_row: int = 2) -> List[str]:
    wb = load_workbook(excel_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]

    values = []
    for cell in ws[PLATE_COL][start_row - 1:]:
        if cell.value is None:
            continue

        s = str(cell.value).strip()
        if s:
            values.append(s)

    return values


def detect_vehicle_type(ws) -> str:
    """
    Detect vehicle type from D2.

    E / RE = electric = GC337
    Others = gasoline = GC335
    """
    first_plate = ws[f"{PLATE_COL}2"].value

    if first_plate is None or str(first_plate).strip() == "":
        raise ValueError(f"No license plate found at {PLATE_COL}2.")

    plate = str(first_plate).strip().upper()

    if plate.startswith("E") or plate.startswith("RE"):
        return "337"

    return "335"


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


def run_job(
    excel_path: str,
    sheet_name: str,
    headless: bool = True,
    output_path: Optional[str] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
):
    excel_path = excel_path.strip().strip('"').strip("'")

    if not excel_path.lower().endswith(".xlsx"):
        raise ValueError("Please select a .xlsx file")

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"File not found: {excel_path}")

    if output_path:
        output_path = output_path.strip().strip('"').strip("'")

        if not output_path.lower().endswith(".xlsx"):
            raise ValueError("Output file must be a .xlsx file")

        if not os.path.exists(output_path):
            raise FileNotFoundError(f"Output file not found: {output_path}")

        target_file_to_backup = output_path
    else:
        target_file_to_backup = excel_path

    create_backup_copy(target_file_to_backup)

    wb = load_workbook(excel_path)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]

    plates = read_column_values(excel_path, sheet_name, start_row=2)
    total = len(plates)

    if total == 0:
        raise ValueError(f"No license plates found in column {PLATE_COL}.")

    if progress_cb:
        progress_cb(0, total, "")

    vehicle_type = detect_vehicle_type(ws)
    print(f"Detected vehicle type: {vehicle_type}")

    if vehicle_type == "337":
        client = GC337Client(headless=headless)
        keys_order = ["transId", "crtDate", "status", "refundDt"]
    else:
        client = GC335Client(headless=headless)
        keys_order = [
            "receiveDate",
            "dealNote",
            "taxreFund",
            "custCd",
            "caseNo",
            "msg",
            "dealDate",
            "rptDate",
        ]

    try:
        for idx, plate in enumerate(plates, start=1):
            if progress_cb:
                progress_cb(idx, total, plate)

            row = find_row_by_value(ws, plate, start_row=2)

            if row is None:
                row = first_empty_row_any(ws, start_row=2)
                ws.cell(row=row, column=4, value=plate)

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

    if output_path:
        safe_save_workbook(wb, output_path)
    else:
        safe_save_workbook(wb, excel_path)


def main():
    excel_path = input("Enter the Excel file name (with .xlsx): ").strip().strip('"').strip("'")

    wb = load_workbook(excel_path)

    print("\nAvailable sheets:")
    for name in wb.sheetnames:
        print(f"- {name}")

    sheet_name = input("\nEnter the sheet name to write data into: ").strip()

    run_job(
        excel_path=excel_path,
        sheet_name=sheet_name,
        headless=True,
        output_path=None,
        progress_cb=None,
    )


if __name__ == "__main__":
    main()