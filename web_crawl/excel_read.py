from openpyxl.utils.exceptions import InvalidFileException
from openpyxl import load_workbook

def read_data():
    while True:
        file_name = input("Enter the Excel file name (with .xlsx): ")
        try:
            workbook = load_workbook(file_name)
            break
        except FileNotFoundError:
            print("File not found. Please try again.")
        except InvalidFileException:
            print("Invalid file. Please enter a valid Excel file.")

    ws = workbook.active
    print("\nAvailable sheets:")

    for name in workbook.sheetnames:
        print(f"- {name}")

    while True:
        sheet_name = input("\nEnter the sheet name you want to access: ")
        if sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            break
        else:
            print("Sheet not found. Please choose from the list above.")

    print(f"\nYou are now accessing the sheet: {sheet_name}")

    while True:
        col_ref = input("Enter the column you want to read (A-Z): ").upper()
        if col_ref.isalpha():
            col_values = [cell.value for cell in ws[col_ref][1:] if cell.value is not None]
            return col_values
        else:
            print("Please enter a valid column letter (A-ZZ).")

def read_data_from(file_path: str, sheet_name: str, col_ref: str):
    """
    Non-interactive: read values from a specific sheet+column.
    Returns a list of non-empty cell values from row 2 onward.
    """
    wb = load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {wb.sheetnames}")

    ws = wb[sheet_name]
    col_ref = col_ref.strip().upper()

    values = []
    for cell in ws[col_ref][1:]:
        if cell.value is None:
            continue
        s = str(cell.value).strip()
        if s:
            values.append(cell.value)
    return values