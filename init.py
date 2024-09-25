import os
import pandas as pd
import warnings
from concurrent.futures import ThreadPoolExecutor
import cProfile
import openpyxl

# To avoid errors
warnings.simplefilter("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None


def parse_excel(path, sheets):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File '{path}' not found.")
    dataframes = {}
    # excel_data = pd.ExcelFile(path)
    with pd.ExcelFile(path, engine="openpyxl") as excel_data:
        for i in sheets:
            if i in excel_data.sheet_names:
                dataframes[i] = pd.read_excel(excel_data, sheet_name=i)
            else:
                warnings.warn(f"Warning: Sheet {i} not found in the file.")
    return dataframes


def parse_sheet(excel_data, sheet_name):
    """Helper function to read a single sheet."""
    if sheet_name in excel_data.sheet_names:
        return pd.read_excel(excel_data, sheet_name=sheet_name)
    else:
        warnings.warn(f"Warning: Sheet {sheet_name} not found in the file.")
        return None


def parse_excel2(path, sheets):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File '{path}' not found.")

    dataframes = {}

    with pd.ExcelFile(path, engine="openpyxl") as excel_data:
        # Use ThreadPoolExecutor to parallelize sheet reading
        with ThreadPoolExecutor() as executor:
            futures = {executor.submit(parse_sheet, excel_data, sheet): sheet for sheet in sheets}

            # Collect results as they complete
            for future in futures:
                sheet_name = futures[future]
                df = future.result()
                if df is not None:
                    dataframes[sheet_name] = df

    return dataframes


def load_sheet_as_df(file_path, sheet_name):
    """Read a specific sheet using openpyxl and convert it into a DataFrame."""
    wb = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(ws)
        data = ws.values
        columns = next(data)  # First row is assumed to be the header
        return pd.DataFrame(data, columns=columns)
    else:
        return None


def parse_excel3(path, sheets):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File '{path}' not found.")

    dataframes = {}

    for sheet in sheets:
        df = load_sheet_as_df(path, sheet)
        if df is not None:
            dataframes[sheet] = df
        else:
            print(f"Warning: Sheet {sheet} not found in the file.")

    return dataframes


path = r'C:\Users\Mano-BRCGE\Desktop\NeWare Data\20240923\240910-C04-C-NCMA90-EC-DEC-0.01091-RPF4.xlsx'
sheets = ['cycle', 'step', 'record']
print('Started..')

wf = parse_excel(path, sheets)
for sheet_name, df in wf.items():  # Converting each sheet into a separate dataframes
    globals()[sheet_name] = df  # Creates a global variable with the name of the sheet

print(f'Workframe is a {type(wf)}')
print(f'Sheet cycle is now a {type(cycle)}')
print(f'Sheet step is now a {type(step)}')
print(f'Sheet record is now a {type(record)}')
print('..Finished')
