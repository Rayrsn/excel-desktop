import pandas as pd
import numpy as np
import openpyxl


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
    operation = ""
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
            # NOTE: chage header if neaded
            elif operation == "change_type":
                # if haeder and arguments of change type are note same it will add arguments into header
                if list(df.columns) != [i for i, _ in arguments]:
                    # Update the header row
                    update_header_list = []
                    update_header_list.extend(df.columns)
                    for i, _ in arguments:
                        if i not in df.columns:
                            update_header_list.append(i)

                    workbook = openpyxl.load_workbook(filepath)
                    write_header(
                        workbook=workbook,
                        target_sh_name=sheet_name,
                        source_sh="",
                        start_header_row=1,
                        header_list=update_header_list,
                    )
                    workbook.save(filepath)

                    df = pd.read_excel(
                        filepath, sheet_name=sheet_name, engine="openpyxl"
                    )

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
                filter_condition = arguments[0]
                if filter_condition == "'Type of Offence' == 'Either Way'":
                    df = df[df["Type of Offence"] == "Either Way"]
                elif filter_condition == "'Type of Offence' == 'Indictable'":
                    df = df[df["Type of Offence"] == "Either Way"]
                else:
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
        print(f"in sheet '{sheet_name}' and operation '{operation}' ")
        return None


def main(filepath):
    for item in queries:
        sheet_name = item["item_name"]
        queries_list = []
        for query in item["quires"]:
            queries_list.append(query)

        df = process_excel_queries(filepath, sheet_name, queries_list)
        try:
            writer = pd.ExcelWriter(
                filepath, engine="openpyxl", mode="a", if_sheet_exists="overlay"
            )  # Use 'openpyxl' for append mode
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            writer.book.save(filepath)
        except Exception as e:
            print(e)


if __name__ == "__main__":
    from query_list import queries
    from btn import write_header

    filepath = "../Law Clients.xlsm"

    main(filepath)
else:
    from utils.query_list import queries
    from utils.btn import write_header
