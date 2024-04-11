import requests

## GET REQUESTS ##

def get_data(url):
    try:
        response = requests.get(url, timeout=5)
    except Exception as e:
        print(e)
        return None
    return response.json()

def get_sheets(json_data):
    # Returns a list of all the sheet names
    return list(json_data['data'].keys())

def get_headers(json_data, sheet):
    # Returns a list of all the headers in the sheet
    # the headers are in ['data']['headers']
    return json_data['headers'][sheet]

def get_data_from_column(json_data, sheet, column):
    # Returns a list of all the data in a column
    return [row[column] for row in json_data['data'][sheet]]

def get_data_from_row(json_data, sheet, row_idx):
    # Returns a dictionary representing the data in a row
    return json_data['data'][sheet][row_idx]

def get_data_from_cell(json_data, sheet, row_idx, column):
    # Returns the data in a specific cell
    if column not in json_data['headers'][sheet]:
        return None
    elif row_idx >= len(json_data['data'][sheet]):
        return None
    else:
        return json_data['data'][sheet][row_idx][column]

def get_sheet_data(json_data, sheet):
    # Returns all the data in a sheet
    return json_data['data'][sheet]

def get_row_count(json_data, sheet):
    # Returns the number of rows in a sheet
    return len(json_data['data'][sheet])

def get_column_count(json_data, sheet):
    # Returns the number of columns in a sheet
    return len(json_data['headers'][sheet])

## POST REQUESTS ##

def post_data(url, data):
    # Sends a POST request to the server with the provided data
    response = requests.post(url, json=data)
    print(response.text)
    return response


## OPERATIONS ##

def add_sheet(json_data, sheet_name):
    # Adds a new sheet to the JSON data
    if sheet_name not in json_data['data']:
        json_data['data'][sheet_name] = []

def gen_report(url, rep_name):
    # Generates a report by sending a GET request to the server
    data = get_data(url)
    with open(rep_name, 'w') as f:
        f.write(str(data))
