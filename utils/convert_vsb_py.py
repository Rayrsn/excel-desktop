import openpyxl
from datetime import datetime, timedelta


def generate_monthly_cases_report(filepath):
    """Generates a monthly cases report from an Excel workbook."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    data_worksheet = workbook["Opening File"]
    report_worksheet_name = "Upcoming Cases this Month"
    workbook.create_sheet("Upcoming Cases this Month")
    report_worksheet = workbook[report_worksheet_name]
    # report_worksheet = workbook.get_sheet_by_name(report_worksheet_name)
    # if not report_worksheet:
    #     report_worksheet = workbook.create_sheet(report_worksheet_name)
    # else:
    for row in report_worksheet.iter_rows():
        for cell in row:
            cell.value = None  # Clear existing data

    # Get references to tables
    master_data_table = data_worksheet["MasterData"]

    # Filter data for the current month
    header_row = master_data_table.min_row
    current_month_start = datetime.now().replace(day=1)
    current_month_end = (current_month_start + relativedelta(months=+1)) - timedelta(
        days=1
    )
    filtered_data = [
        row
        for row in master_data_table.iter_rows(min_row=header_row + 1)
        if current_month_start <= row[0].value <= current_month_end
    ]

    # Copy filtered data to report worksheet
    report_worksheet.append(master_data_table[header_row])  # Paste headers
    for row in filtered_data:
        report_worksheet.append([cell.value for cell in row])

    # Convert pasted data to Excel table
    report_table = openpyxl.worksheet.table.Table(
        displayName="Upcoming_Cases_Table", ref=report_worksheet.tables.tables[0].ref
    )
    report_worksheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Monthly cases report generated successfully!")


def generate_weekly_cases_report(filepath):
    """Generates a weekly cases report from an Excel workbook."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    data_worksheet = workbook["Opening File"]
    report_worksheet_name = "Upcoming Cases this Week"
    workbook.create_sheet(report_worksheet_name)
    report_worksheet = workbook[report_worksheet_name]
    # report_worksheet = workbook.get_sheet_by_name(report_worksheet_name)
    # if not report_worksheet:
    #     report_worksheet = workbook.create_sheet(report_worksheet_name)
    # else:
    for row in report_worksheet.iter_rows():
        for cell in row:
            cell.value = None  # Clear existing data

    # Get references to tables
    master_data_table = data_worksheet["MasterData"]

    # Calculate start and end dates for the current week
    today = date.today()
    weekday = today.weekday()  # 0 for Monday, 6 for Sunday
    start_date = today - timedelta(days=weekday)  # Monday of the current week
    end_date = start_date + timedelta(days=6)  # Sunday of the current week

    # Filter data for the current week
    header_row = master_data_table.min_row
    filtered_data = [
        row
        for row in master_data_table.iter_rows(min_row=header_row + 1)
        if start_date <= row[0].value <= end_date
    ]

    # Copy filtered data to report worksheet
    report_worksheet.append(master_data_table[header_row])  # Paste headers
    for row in filtered_data:
        report_worksheet.append([cell.value for cell in row])

    # Convert pasted data to Excel table
    report_table = openpyxl.worksheet.table.Table(
        displayName="Upcoming_Cases_Table", ref=report_worksheet.tables.tables[0].ref
    )
    report_worksheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Weekly cases report generated successfully!")


def generate_legal_aid_report(filepath):
    """Generates a legal aid report from an Excel workbook."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    data_worksheet = workbook["Opening File"]
    report_worksheet_name = "Legal Aid Report"
    report_worksheet = workbook.get_sheet_by_name(report_worksheet_name)
    if not report_worksheet:
        report_worksheet = workbook.create_sheet(report_worksheet_name)
    else:
        for row in report_worksheet.iter_rows():
            for cell in row:
                cell.value = None  # Clear existing data

    # Get references to tables
    master_data_table = data_worksheet["MasterData"]

    # Find the legal aid column index
    legal_aid_col = None
    for col in master_data_table.iter_cols(min_row=1):
        if col[0].value == "Legal Aid":
            legal_aid_col = col[0].column

    # Check if legal aid column exists
    if not legal_aid_col:
        print("Error: Legal Aid column not found in MasterData table!")
        return

    # Find the last row of data
    last_row = master_data_table.max_row

    # Initialize counters
    submitted_count = 0
    refused_count = 0
    date_stamped_count = 0
    approved_count = 0
    appealed_count = 0

    # Loop through the data and count clients in each category
    for row in master_data_table.iter_rows(min_row=2):
        if row[legal_aid_col - 1].value:  # Check if cell is not empty
            legal_aid_status = row[legal_aid_col - 1].value
            if legal_aid_status == "Submitted":
                submitted_count += 1
            elif legal_aid_status == "Refused":
                refused_count += 1
            elif legal_aid_status == "Date Stamped":
                date_stamped_count += 1
            elif legal_aid_status == "Approved":
                approved_count += 1
            elif legal_aid_status == "Appealed":
                appealed_count += 1

    # Write data to report worksheet
    report_worksheet["A1"] = "Legal Aid Category"
    report_worksheet["B1"] = "Number of Clients"
    report_worksheet["A2"] = "Submitted"
    report_worksheet["B2"] = submitted_count
    report_worksheet["A3"] = "Refused"
    report_worksheet["B3"] = refused_count
    report_worksheet["A4"] = "Date Stamped"
    report_worksheet["B4"] = date_stamped_count
    report_worksheet["A5"] = "Approved"
    report_worksheet["B5"] = approved_count
    report_worksheet["A6"] = "Appealed"
    report_worksheet["B6"] = appealed_count

    # Calculate and write total client count
    client_count = (
        submitted_count
        + refused_count
        + date_stamped_count
        + approved_count
        + appealed_count
    )
    report_worksheet["A8"] = "Total Number of Clients"
    report_worksheet["B8"] = client_count

    # Create a table from the report data
    report_table = openpyxl.worksheet.table.Table(
        displayName="Legal Aid Report", ref="A1:B8"
    )
    report_table.table_style = "TableStyleLight9"  # Optional table style
    report_worksheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Legal aid report generated successfully!")


def generate_bail_refused_report(filepath):
    """Generates a bail refused report from an Excel workbook."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    magistrates_sheet = workbook["Magistrates Merge"]
    crown_court_sheet = workbook["Crown Court Merge"]
    report_sheet_name = "Clients in Prison Report"
    report_sheet = workbook.get_sheet_by_name(report_sheet_name)
    if not report_sheet:
        report_sheet = workbook.create_sheet(report_sheet_name)
    else:
        for row in report_sheet.iter_rows():
            for cell in row:
                cell.value = None  # Clear existing data

    # Define column indices (assuming headers are in the first row)
    bail_col = get_column_index(magistrates_sheet, 1, "Bail")
    prison_number_col = get_column_index(magistrates_sheet, 1, "Prison Number")
    file_no_col = get_column_index(magistrates_sheet, 1, "File No.")

    # Copy headers from Magistrates Merge sheet to report sheet
    for i in range(1, 18):  # Assuming 17 headers (A-Q)
        report_sheet.cell(row=1, column=i).value = magistrates_sheet.cell(
            row=1, column=i
        ).value

    # Initialize variables
    report_row = 2  # Start from row 2 for data
    used_file_numbers = set()  # Use a set to store unique file numbers

    # Extract data from Magistrates Merge sheet
    extract_data_for_bail_refused(
        magistrates_sheet.iter_rows(min_row=2),
        report_sheet,
        report_row,
        bail_col,
        prison_number_col,
        file_no_col,
        used_file_numbers,
    )

    # Extract data from Crown Court Merge sheet
    extract_data_for_bail_refused(
        crown_court_sheet.iter_rows(min_row=2),
        report_sheet,
        report_row,
        bail_col,
        prison_number_col,
        file_no_col,
        used_file_numbers,
    )

    # Create a table from the report data
    report_table = openpyxl.worksheet.table.Table(
        displayName="Clients in Prison Report",
        ref="A1:" + report_sheet.cells(report_sheet.max_row, 1).end(xlUp).row,
    )
    report_table.table_style = "TableStyleLight9"  # Optional table style
    report_sheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Bail refused report generated successfully!")


def extract_data_for_bail_refused(
    data_rows,
    report_sheet,
    report_row,
    bail_col,
    prison_number_col,
    file_no_col,
    used_file_numbers,
):
    """Extracts data for bail refused cases and writes to report sheet."""

    bail_criteria = ("Bail Refused", "JC Bail Refused")

    for row in data_rows:
        cells = [cell.value for cell in row]
        if (
            cells[bail_col - 1] in bail_criteria
            and cells[prison_number_col - 1]
            and cells[file_no_col - 1] not in used_file_numbers
        ):
            used_file_numbers.add(cells[file_no_col - 1])
            for i in range(len(cells)):
                report_sheet.cell(row=report_row, column=i + 1).value = cells[i]
            report_row += 1


def get_column_index(worksheet, header_row, header_text):
    """Gets the column index of a header in a worksheet."""

    for cell in worksheet.iter_rows(min_row=header_row):
        for col, value in enumerate(cell, 1):
            if value == header_text:
                return col
    return 0  # Return 0 if header not found


# def get_column_index(worksheet, header):
#     """Gets the column index of a header in a worksheet."""

#     for col, value in enumerate(worksheet.iter_rows(min_row=1), 1):
#         if header in value:
#             return col
#     return 0


def generate_empty_counsel_report(filepath):
    """Generates a report of rows with empty counsel names in the Crown Court sheet."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    crown_court_sheet = workbook["Crown Court Merge"]
    report_sheet_name = "Empty Counsel Report"
    report_sheet = workbook.get_sheet_by_name(report_sheet_name)
    if not report_sheet:
        report_sheet = workbook.create_sheet(report_sheet_name)
    else:
        for row in report_sheet.iter_rows():
            for cell in row:
                cell.value = None  # Clear existing data

    # Define counsel column index (assuming headers are in the first row)
    counsel_col = get_column_index(crown_court_sheet, 1, "Name of Counsel")

    # Copy headers from Crown Court sheet to report sheet
    for i in range(1, crown_court_sheet.max_column + 1):
        report_sheet.cell(row=1, column=i).value = crown_court_sheet.cell(
            row=1, column=i
        ).value

    # Initialize report row
    report_row = 2

    # Extract data from Crown Court sheet with empty counsel names
    extract_data_for_empty_counsel(
        crown_court_sheet.iter_rows(min_row=2), report_sheet, report_row, counsel_col
    )

    # Create a table from the report data
    report_table = openpyxl.worksheet.table.Table(
        displayName="Empty Counsel Report",
        ref="A1:" + report_sheet.cells(report_sheet.max_row, 1).end(xlUp).row,
    )
    report_table.table_style = "TableStyleLight9"  # Optional table style
    report_sheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Empty counsel report generated successfully!")


def extract_data_for_empty_counsel(data_rows, report_sheet, report_row, counsel_col):
    """Extracts rows with empty counsel names and writes them to the report sheet."""

    for row in data_rows:
        cells = [cell.value for cell in row]
        if not cells[counsel_col - 1]:
            for i in range(len(cells)):
                report_sheet.cell(row=report_row, column=i + 1).value = cells[i]
            report_row += 1


def generate_non_zero_balance_report(filepath):
    """Generates a report of rows with non-zero balance in the Road Traffic sheet."""

    workbook = openpyxl.load_workbook(filepath)

    # Get references to worksheets
    road_traffic_sheet = workbook["Road Traffic"]
    report_sheet_name = "Non-Zero Balance Report"
    report_sheet = workbook.get_sheet_by_name(report_sheet_name)
    if not report_sheet:
        report_sheet = workbook.create_sheet(report_sheet_name)
    else:
        for row in report_sheet.iter_rows():
            for cell in row:
                cell.value = None  # Clear existing data

    # Define balance column index (assuming headers are in the first row)
    balance_col = get_column_index(road_traffic_sheet, 1, "Balance")

    # Copy headers from Road Traffic sheet to report sheet
    for i in range(1, road_traffic_sheet.max_column + 1):
        report_sheet.cell(row=1, column=i).value = road_traffic_sheet.cell(
            row=1, column=i
        ).value

    # Initialize report row
    report_row = 2

    # Extract data from Road Traffic sheet with non-zero balance
    extract_data_for_non_zero_balance(
        road_traffic_sheet.iter_rows(min_row=2), report_sheet, report_row, balance_col
    )

    # Create a table from the report data
    report_table = openpyxl.worksheet.table.Table(
        displayName="Non-Zero Balance Report",
        ref="A1:" + report_sheet.cells(report_sheet.max_row, 1).end(xlUp).row,
    )
    report_table.table_style = "TableStyleLight9"  # Optional table style
    report_sheet.add_table(report_table)

    # Save the workbook
    workbook.save(filepath)

    print("Non-zero balance report generated successfully!")


def extract_data_for_non_zero_balance(data_rows, report_sheet, report_row, balance_col):
    """Extracts rows with non-zero balance and writes them to the report sheet."""

    for row in data_rows:
        cells = [cell.value for cell in row]
        if cells[balance_col - 1] is not None and cells[balance_col - 1] != 0:
            for i in range(len(cells)):
                report_sheet.cell(row=report_row, column=i + 1).value = cells[i]
            report_row += 1


def generate_stage_reports(filepath):
    """Generates stage reports for the current month from the Crown Court Merge sheet."""

    workbook = openpyxl.load_workbook(filepath)
    crown_court_sheet = workbook["Crown Court Merge"]
    current_month = datetime.date.today().month

    # Generate reports for each stage
    stages = ["Stage 1", "Stage 2", "Stage 3", "Stage 4"]
    for stage in stages:
        generate_stage_report(workbook, crown_court_sheet, stage, current_month)

    workbook.save(filepath)
    print("Stage reports generated successfully!")


def generate_stage_report(workbook, ws, stage, current_month):
    """Generates a report for a specific stage and month."""

    report_sheet_name = stage + " Report"
    report_sheet = workbook.get_sheet_by_name(report_sheet_name)
    if not report_sheet:
        report_sheet = workbook.create_sheet(report_sheet_name)
    else:
        for row in report_sheet.iter_rows():
            for cell in row:
                cell.value = None  # Clear existing data

    stage_col = get_column_index(ws, stage)

    # Copy headers
    for i in range(1, ws.max_column + 1):
        report_sheet.cell(row=1, column=i).value = ws.cell(row=1, column=i).value

    report_row = 2

    # Filter and extract data
    for cell in ws.iter_cols(min_row=2, min_col=stage_col, max_col=stage_col):
        for row_cell in cell:
            if (
                isinstance(row_cell.value, datetime.date)
                and row_cell.value.month == current_month
            ):
                extract_data_for_stage_report(
                    row_cell.row, ws, report_sheet, report_row
                )

    # Create a table
    report_table = openpyxl.worksheet.table.Table(
        displayName=stage,
        ref="A1:" + report_sheet.cells(report_sheet.max_row, 1).end("up").row,
    )
    report_table.table_style = "TableStyleLight9"  # Optional table style
    report_sheet.add_table(report_table)

    if report_row == 2:
        print(f"No data found for {stage}")


def extract_data_for_stage_report(row_num, ws, report_sheet, report_row):
    """Copies a row of data to the report sheet."""

    for i in range(1, ws.max_column + 1):
        report_sheet.cell(row=report_row, column=i).value = ws.cell(
            row=row_num, column=i
        ).value
    report_row += 1


if __name__ == "__main__":
    filepath = "../Law Clients.xlsm"
    # generate_stage_reports(filepath)

    # generate_non_zero_balance_report(filepath)
    # generate_empty_counsel_report(filepath)
    # generate_bail_refused_report(filepath)
    # generate_legal_aid_report(filepath)
    # generate_weekly_cases_report(filepath)
    generate_monthly_cases_report(filepath)
