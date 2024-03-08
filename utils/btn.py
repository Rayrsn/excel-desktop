import openpyxl
from datetime import datetime, timedelta


def create_upcoming_month_sheet(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)

    # Access the "Opening File" worksheet
    ws_source = wb["Opening File"]

    # Check if "Upcoming Cases this Month" sheet exists, if not create it
    if "Upcoming Cases this Month" not in wb.sheetnames:
        wb.create_sheet("Upcoming Cases this Month")
    ws_target = wb["Upcoming Cases this Month"]

    # Assuming the first row is the header and we're filtering based on a condition in column 1
    # Copy header row to target sheet
    for col in range(1, ws_source.max_column + 1):
        ws_target.cell(row=1, column=col).value = ws_source.cell(
            row=1, column=col
        ).value

    target_row = 1  # Start writing to the second row of the target sheet
    # data of first sheet will start from 17 row
    for index, row in enumerate(ws_source.iter_rows(min_row=17, values_only=True)):
        # Apply filtering condition - for example, checking a date or a specific text
        # This is where you'd customize based on your actual filter condition
        if index == 0:
            print("yes it is ")
            for col, value in enumerate(row, start=1):
                ws_target.cell(row=target_row, column=col).value = value
            target_row += 1
        elif date_open_filter(row) and row:
            for col, value in enumerate(row, start=1):
                ws_target.cell(row=target_row, column=col).value = value
            target_row += 1

    # Save the workbook with the new data
    wb.save("upcommit_month.xlsx")


def create_upcoming_week_sheet(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)

    # Access the "Opening File" worksheet
    ws_source = wb["Opening File"]

    # Check if "Upcoming Cases this Month" sheet exists, if not create it
    if "Upcoming Cases this Month" not in wb.sheetnames:
        wb.create_sheet("Upcoming Cases this Month")
    ws_target = wb["Upcoming Cases this Month"]

    # Assuming the first row is the header and we're filtering based on a condition in column 1
    # Copy header row to target sheet
    for col in range(1, ws_source.max_column + 1):
        ws_target.cell(row=1, column=col).value = ws_source.cell(
            row=1, column=col
        ).value

    target_row = 1  # Start writing to the second row of the target sheet
    # data of first sheet will start from 17 row
    for index, row in enumerate(ws_source.iter_rows(min_row=17, values_only=True)):
        # Apply filtering condition - for example, checking a date or a specific text
        # This is where you'd customize based on your actual filter condition
        if index == 0:
            print("yes it is ")
            for col, value in enumerate(row, start=1):
                ws_target.cell(row=target_row, column=col).value = value
            target_row += 1
        elif date_open_filter(row) and row:
            for col, value in enumerate(row, start=1):
                ws_target.cell(row=target_row, column=col).value = value
            target_row += 1

    # Save the workbook with the new data
    wb.save("upcommit_month.xlsx")


def create_legal_aid_report_sheet(workbook_path):
    def calc_report_vars(last_column_values):
        vals = {
            "Submitted": 0,
            "Refused": 0,
            "Date Stamped": 0,
            "Approved": 0,
            "Appealed": 0,
        }
        for i in last_column_values:
            if i and i in vals:
                vals[i] += 1
        return vals

    # Load the Excel file
    wb = openpyxl.load_workbook(workbook_path)

    # Select the specific sheet
    sheet = wb["Opening File"]  # Replace 'Sheet1' with the name of your sheet

    # Find the last column with data
    last_column = sheet.max_column

    # Create an empty list to store values from the last column
    last_column_values = [
        sheet.cell(row=row, column=last_column).value
        for row in range(18, sheet.max_row + 1)
    ]
    print(calc_report_vars(last_column_values))

    # Close the Excel file
    wb.close()


def is_current_week(date_str):
    # Assuming the date is in 'YYYY-MM-DD' format; adjust the format as necessary
    date_opened = datetime.strptime(date_str, "%Y-%m-%d")

    today = datetime.today()
    start_week = today - timedelta(days=today.isoweekday() - 1)  # Monday
    end_week = start_week + timedelta(days=6)  # Sunday

    return start_week <= date_opened <= end_week


def date_open_filter(row):
    # Assume the "date opened" is in the first column (index 0) of the row
    date_opened_str = row[7]
    print(type(date_opened_str))
    print(date_opened_str)

    # Assuming the date is in 'YYYY-MM-DD' format; adjust the format as necessary
    # date_opened = datetime.strptime(date_opened_str, "%d.%m.%Y")
    date_opened = date_opened_str

    # Get the first and last day of the current month
    today = datetime.today()
    first_day_of_month = datetime(today.year, today.month, 1)
    last_day_of_month = datetime(today.year, today.month + 1, 1) - timedelta(days=1)

    # Check if the date_opened falls within the current month
    return first_day_of_month <= date_opened <= last_day_of_month


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
# create_upcoming_month_sheet(excel_file)
create_legal_aid_report_sheet(excel_file)
