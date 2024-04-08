import requests

# Example JSON response from the server
json =""" 
{
    "data": {
        "Opening_File": [
            {
                "id": 1,
                "Sr_No": "13",
                "Office": "joe",
                "Matter_Type": "joe",
                "Email": "joe@joe.com",
                "Previous_Number": "joe",
                "CRIME": "joe",
                "File_No": "joe",
                "Date_Opened": null,
                "Fee_Earner": null,
                "Clients_Surname": null,
                "Clients_Forename": null,
                "Clients_Title": null,
                "Marital_Status": null,
                "Letters_to_Home_Address": null,
                "Address": null,
                "City": null,
                "Postcode": null,
                "Postal_Address_if_Different": null,
                "Postal_Address_Postcode": null,
                "Home_Telephone": null,
                "Work_Telephone": null,
                "Mobile_Number": null,
                "Occupation": null,
                "Date_of_Birth": null,
                "Ethnicity": null,
                "HMP": null,
                "Prison_Number": null,
                "National_Insurance_Number": null,
                "_3rd_Party": null,
                "Initial": null,
                "Conflict": null,
                "Date": null,
                "Costs_Information": null,
                "Cost_Estimate": null,
                "Charge_Basis": null,
                "Court": null,
                "Legal_Aid": null
            }
        ],
        "Police_Station": [],
        "List_of_Letters": [],
        "Bail": [],
        "Magistrates_Merge": [],
        "Magistrates": [],
        "Crown_Court_Merge": [],
        "Crown_Court": [],
        "Road_Traffic": [],
        "Appeals": [],
        "Appeals_Magistrates": [],
        "Appeals_Crown_Court": [],
        "Appeals_Road_Traffic": [],
        "PoliceStation_to_Magistrates": [],
        "Magistrates_to_Crown_Court_1": [],
        "Magistrates_to_Crown_Court_2": []
    }
}
"""

## GET REQUESTS ##

def get_data(url):
    response = requests.get(url)
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
    return response.json()


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
