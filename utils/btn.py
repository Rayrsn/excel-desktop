import openpyxl
from datetime import datetime


def create_upcoming_month_sheet(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)

    # Access the "Opening File" worksheet
    ws_source = wb["Opening File"]

    # Check if "Upcoming Cases this Month" sheet exists, if not create it
    if "Upcoming Cases this Month" not in wb.sheetnames:
        wb.create_sheet("Upcoming Cases this Month")
    ws_target = wb["Upcoming Cases this Month"]

    target_row = 2  # Start writing to the second row of the target sheet
    for row in ws_source.iter_rows(min_row=2, values_only=True):
        # Apply filtering condition - for example, checking a date or a specific text
        # This is where you'd customize based on your actual filter condition
        if your_filter_condition(row):
            for col, value in enumerate(row, start=17):
                ws_target.cell(row=target_row, column=col).value = value
            target_row += 1

    # Check if the sheet has any data to clear
    if ws_target.max_row > 0:
        # ws.delete_rows() can be used to delete rows in the given range
        ws_target.delete_rows(0, ws_target.max_row)

    # Save the workbook with the new data
    wb.save("edit_month.xlsx")


def your_filter_condition(row):
    # Example filter condition: True if the row meets the criteria, False otherwise
    # Adjust this function based on your actual filtering criteria
    return True  # Placeholder - replace with your actual condition


def clear_worksheet(workbook_path, sheet_name):
    wb = openpyxl.load_workbook(workbook_path)
    ws = wb[sheet_name]

    # Check if the sheet has any data to clear
    if ws.max_row > 0:
        # ws.delete_rows() can be used to delete rows in the given range
        ws.delete_rows(1, ws.max_row)

    # Save the workbook after clearing the sheet
    wb.save(workbook_path)


# Example usage
# clear_worksheet("path_to_your_excel_file.xlsx", "SheetNameToClear")

# Replace 'path_to_your_excel_file.xlsx' with the actual path to your workbook
excel_file = "../../docs/Law Clients Excel Sheet Shared_MainV3.xlsx"
create_upcoming_month_sheet(excel_file)
