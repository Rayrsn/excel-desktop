changed loadExcelData to loadJsonData (uses the same logic, but modified to work with json)


## GET REQUESTS

- get_data
  - arguments: url
  - returns: json_data

- get_sheets
    - arguments: json_data
    - returns: sheet_names

- get_headers
    - arguments: json_data, sheet_name
    - returns: headers

- get_data_from_column
    - arguments: json_data, sheet_name, column_name
    - returns: data

- get_data_from_row
    - arguments: json_data, sheet_name, row_number
    - returns: data

- get_data_from_cell
    - arguments: json_data, sheet_name, column_name, row_number
    - returns: data

## POST REQUESTS

- post_data
    - arguments: url, data
    - returns: json_data

## OPERATIONS

- add_sheet
    - arguments: json_data, sheet_name, data
    - returns: json_data

- gen_report
    - arguments: url, report_type
    - returns: sheet_data