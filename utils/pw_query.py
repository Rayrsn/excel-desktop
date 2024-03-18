import pandas as pd
import numpy as np


def get_sheet(filepath, sheet_name):
    if sheet_name == "Opening File":
        return pd.read_excel(
            filepath,
            engine="openpyxl",
            sheet_name=sheet_name,
            index_col=0,
            nrows=None,
            skiprows=16,
        )
    return pd.read_excel(
        filepath, index_col=0, nrows=None, sheet_name=sheet_name, engine="openpyxl"
    )


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
    df = pd.read_excel(filepath, sheet_name=sheet_name, engine="openpyxl")
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
                        engine="openpyxl",
                    )
                else:
                    df = pd.read_excel(
                        filepath, sheet_name=sheet_name, engine="openpyxl"
                    )
            elif operation == "change_type":
                # Set data types for specific columns
                for col_name, col_type in arguments:
                    df[col_name] = df[col_name].astype(col_type)
            elif operation == "remove_columns":
                df = df.drop(arguments, axis=1)

            elif operation == "combine_sheets":
                data_frames = [get_sheet(filepath, i) for i in arguments]
                df = pd.concat(data_frames, ignore_index=True)

            # NOTE: must check all conditions
            # BUG: check have yes conditions and Variable that have space in it
            elif operation == "filter_rows":
                # Can handle multiple filter conditions (adjust logic as needed)
                # # Filter rows
                # filtered_df = df[df["Court"] == "Magistrates"]
                filter_condition = arguments[0]
                df = df.query(filter_condition)

            elif operation == "select_columns":
                df = df[arguments]
            else:
                print(f"Warning: Unsupported operation '{operation}'.")

        return df

    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
        return None
    except Exception as e:
        print(f"Error processing queries: {e}")
        return None


def main(filepath):
    # couter = 0
    for item in queries:
        # couter += 1
        # if couter == 1 or couter == 2:  # or couter == 3:
        # continue
        sheet_name = item["item_name"]
        queries_list = []
        for query in item["quires"]:
            queries_list.append(query)

        x = get_sheet(filepath, sheet_name)
        print(x)
        df = process_excel_queries(filepath, sheet_name, queries_list)
        print(df)
        # print("--------------------")
        # x = get_sheet(filepath, sheet_name)
        try:
            # BUG: df dataframe is not currently write into sheet
            writer = pd.ExcelWriter(
                filepath, engine="openpyxl", mode="a", if_sheet_exists="overlay"
            )  # Use 'openpyxl' for append mode
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            writer.book.save(filepath)
            # x = get_sheet(filepath, sheet_name)
            # print(x)
            # df.to_excel(filepath, sheet_name=sheet_name, index=False)
            print(f"sheet {sheet_name} was saved")
        except Exception as e:
            print(e)
        # break


if __name__ == "__main__":
    from query_list import queries

    # filepath = "../Law_v3.xlsm"
    filepath = "../Law Clients.xlsm"

    main(filepath)
else:
    from utils.query_list import queries
