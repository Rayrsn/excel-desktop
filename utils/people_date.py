from openpyxl import load_workbook
from pprint import pprint
def last_row(wb, sheet):
    """
    check first column if data doesn't exist it return last row number
    """
    # NOTE: row number 17 is header
    last_row_with_data = 17
    last_row = sheet.max_row
    for row in range(17, last_row):
        selected_row = list(sheet[row])
        # check first header data is Exist or no
        if selected_row[0].value == None:
            last_row_with_data = row
            break
    return last_row_with_data

if __name__ == "__main__":
    excel_file = "../../docs/Law Clients Excel Sheet Shared_MainV3.xlsx"
    # Load the workbook and select the first sheet
    wb = load_workbook(excel_file)
    # This selects the first sheet
    sheet = wb.worksheets[0]

    print(f"last row is {last_row(wb, sheet)}\n")
