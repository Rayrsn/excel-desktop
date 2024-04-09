changed loadExcelData to loadJsonData (uses the same logic, but modified to work with json)

added addSheetToTabs (use cases: adding a report tab in the program)


## GET REQUESTS

- get_data
  - arguments: url
  - returns: json_data

- get_sheets
    - arguments: json_data
    - returns: list

- get_headers
    - arguments: json_data, sheet
    - returns: list

- get_data_from_column
    - arguments: json_data, sheet, column
    - returns: list

- get_data_from_row
    - arguments: json_data, sheet, row_idx
    - returns: dict

- get_data_from_cell
    - arguments: json_data, sheet, row_idx, column
    - returns: data

- get_sheet_data
    - arguments: json_data, sheet
    - returns: list

- get_row_count
    - arguments: json_data, sheet
    - returns: int

- get_column_count
    - arguments: json_data, sheet
    - returns: int

## POST REQUESTS

- post_data
    - arguments: url, data
    - returns: response.json()

## OPERATIONS

- add_sheet
    - arguments: json_data, sheet_name
    - returns: None

- gen_report
    - arguments: url, rep_name
    - returns: None