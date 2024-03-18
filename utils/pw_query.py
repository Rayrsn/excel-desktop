import pandas as pd
import numpy as np
from query_list import queries


# def remove_column_from_table(df, table_name, column_to_remove):
#     """Removes a column from a DataFrame representing a table."""
#     return df.drop(
#         columns=[column_to_remove], inplace=False
#     )  # Avoid modifying original DataFrame


# def select_columns_from_table(df, table_name, columns_to_select):
#     """Selects specific columns from a DataFrame representing a table."""
#     return df[
#         columns_to_select
#     ]  # Return a new DataFrame with only the specified columns


# Assuming the sheet name is 'Magistrates'

# ... Set types for other columns


# def q_g1_magistrates(excel_file):
#     df = pd.read_excel(excel_file, sheet_name="Magistrates")
#     # Set data types (adjust types as needed)
#     df["Sr No."] = df["Sr No."].astype(object)
#     df["Matter Type"] = df["Matter Type"].astype(str)

#     # Remove columns
#     columns_to_remove = ["Previous Number", "CRIME", ...]  # Adjust list
#     df = df.drop(columns_to_remove, axis=1)

#     # Filter rows
#     filtered_df = df[df["Court"] == "Magistrates"]

#     # Optionally remove "Court" column after processing (if needed)
#     filtered_df = filtered_df.drop("Court", axis=1)

#     # Use the filtered_df for further analysis or processing
#     print(filtered_df)  # Print the filtered DataFrame


def process_excel_queries(filepath, sheet_name, queries):
    """Processes a series of Excel query-like operations on a sheet.

    Args:
        filepath (str): The path to the Excel workbook as a string.
        sheet_name (str): The name of the sheet containing the data as a string.
        queries (list): A list of dictionaries representing query operations.
            Each dictionary should have keys:
                - "operation" (str): The operation to perform (e.g., "read", "change_type", "remove_columns", "filter_rows").
                - "arguments" (list, optional): A list of arguments specific to the operation.

    Returns:
        pandas.DataFrame: A pandas DataFrame containing the processed data, or None if errors occur.
    """
    # NOTE : maybe for date column must date datetiem field

    # Read data from the sheet
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    try:
        # Process queries sequentially
        for query in queries:
            operation = query["operation"].lower()
            arguments = query.get("arguments", [])

            if operation == "read":
                # Read data if not already read
                sh_name = arguments["sheet_name"]
                # if df is None:
                if sh_name == "Opening File":
                    df = pd.read_excel(
                        filepath,
                        sheet_name=sh_name,
                        nrows=None,
                        skiprows=16,
                    )
                else:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
            elif operation == "change_type":
                # Set data types for specific columns
                for col_name, col_type in arguments:
                    df[col_name] = df[col_name].astype(col_type)
            elif operation == "remove_columns":
                df = df.drop(arguments, axis=1)

            # NOTE: it must complete
            elif operation == "combine_sheets":
                pass
            elif operation == "filter_rows":
                # Can handle multiple filter conditions (adjust logic as needed)
                filter_condition = arguments[0]
                df = df.query(filter_condition)
            else:
                print(f"Warning: Unsupported operation '{operation}'.")

        # return df

        df.to_excel(filepath, sheet_name="test_sh", index=False)
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
        return None
    except Exception as e:
        print(f"Error processing queries: {e}")
        return None


# Example usage:
# filepath = "../Law_v3.xlsm"
filepath = "../Law_Clients_v4.xlsm"


sheet_name = "Magistrates"
process_excel_queries(filepath, sheet_name, queries)


# processed_data = process_excel_queries(filepath, sheet_name, queries)

# if processed_data is not None:
#     print(processed_data)  # Print the processed DataFrame
# else:
#     print("Error: Processing failed. See error messages for details.")


if __name__ == "__main__":
    pass
    # +---------------------------+
    # |  get data of first sheet  |
    # +---------------------------+
    # df = pd.read_excel(filepath, nrows=None, skiprows=16)

    # excel_file = "../Law_v3.xlsm"
    # q_g1_magistrates(excel_file)

    # create DataFrame
    # df = pd.read_excel(excel_file, sheet_name="Crown Court")
    # df = pd.read_excel(excel_file, sheet_name="Road Traffic")
    # df = pd.read_excel(excel_file, sheet_name="Police Station")
    # z = len(df._info_axis)
    # print(z)
    # x = remove_column_from_table(df, "Filtered Rows2", "Court")
    # print(x)

    # x = df.worksheet.tables.items()
    # print(x)

    # Example usage:
    # modified_df = remove_column_from_table(df, "Filtered Rows", "Court")

    # Example usage:
    # selected_df = select_columns_from_table(
    #     df,
    #     "Filtered Rows",
    #     [
    #         "Office",
    #         "Matter Type",
    #         "Email",
    #         "File No.",
    #         "Fee Earner",
    #         "Client's Surname",
    #         "Client's Forename(s)",
    #         "Marital Status",
    #         "Address",
    #         "City",
    #         "Postcode",
    #         "Mobile Number",
    #         "Occupation",
    #         "Date of Birth",
    #         "Ethnicity",
    #         "HMP",
    #         "Prison Number",
    #         "National Insurance Number",
    #     ],
    # )
    # print(selected_df)

    # df.to_excel(excel_file, index=False)
