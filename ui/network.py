import requests

""" Example JSON response from the server
{
    "sheet1": {
        "data": {
            "column A": [
                "row 1": "value 1",
                "row 2": "value 2",
                "row 3": "value 3"
            ],
            "column B": [
                "row 1": "value 4",
                "row 2": "value 5",
                "row 3": "value 6"
            ],
        }
    }
}
"""

def get_data(url):
    response = requests.get(url)
    return response.json()

def get_sheets(json_data):
    return json_data.keys()

def get_headers(json_data, sheet):
    return json_data[sheet]["data"].keys()

def get_data_from_column(json_data, sheet, column):
    return json_data[sheet]["data"][column]

def get_data_from_row(json_data, sheet, row):
    data = {}
    for column in json_data[sheet]["data"]:
        data[column] = json_data[sheet]["data"][column][row]
    return data

def get_data_from_cell(json_data, sheet, column, row):
    return json_data[sheet]["data"][column][row]
