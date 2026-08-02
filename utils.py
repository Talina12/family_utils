import datetime
import glob
import os
import re
import sys

import pandas as pd
from openpyxl.utils import get_column_letter


def resolve_files(pattern, default_directory="C:/Users/talin/Downloads"):
    file_directory = sys.argv[1] if len(sys.argv) > 1 else default_directory
    print(f"Using file directory: {file_directory}")
    files = glob.glob(os.path.join(file_directory, pattern))
    return file_directory, files


def next_available_path(path):
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    n = 1
    while os.path.exists(f"{root}({n}){ext}"):
        n += 1
    return f"{root}({n}){ext}"


def find_header_row(filepath, header_cell_value, usecols=None, search_rows=20):
    """Find the row index (0-based) containing header_cell_value in any cell,
    so it can be passed as the `header` argument to pd.read_excel."""
    preview_df = pd.read_excel(
        filepath, usecols=usecols, header=None, nrows=search_rows
    )
    matches = preview_df[
        preview_df.apply(
            lambda row: (
                row.astype(str).str.strip().eq(header_cell_value).any()
            ),
            axis=1,
        )
    ]
    if matches.empty:
        raise ValueError(
            f"Could not find header row containing '{header_cell_value}' "
            f"in the first {search_rows} rows of {filepath}"
        )
    return matches.index[0]


def write_dated_excel(result_data_frame, output_path, date_col):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_data_frame.to_excel(writer, index=False)
        date_col_index = result_data_frame.columns.get_loc(date_col) + 1
        worksheet = writer.sheets["Sheet1"]
        date_col_letter = get_column_letter(date_col_index)
        worksheet.column_dimensions[
            date_col_letter].number_format = "dd/mm/yyyy"
    print(f"Wrote {len(result_data_frame)} rows to {output_path}")


def date_filter_converter(value):
    if pd.isna(value) or value is None:
        return None  # This will become NaN

    # Excel already parses real date cells into datetime/Timestamp objects;
    # footer/disclaimer text in the same column comes back as plain str.
    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return value

    # Fallback for dates that arrive as strings, e.g. "07/04/2026"
    str_val = str(value).strip()
    date_patterns = [
        r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$',  # MM/DD/YYYY or DD/MM/YYYY
        r'^\d{4}[/-]\d{1,2}[/-]\d{1,2}$',    # YYYY/MM/DD
        r'^\d{1,2}-\w{3}-\d{2,4}$',          # DD-MMM-YYYY
        r'^\d{1,2}\.\d{1,2}\.\d{2,4}$',      # DD.MM.YYYY or DD.MM.YY
    ]

    if any(re.match(pattern, str_val) for pattern in date_patterns):
        return value
    else:
        return None  # This will become NaN and can be filtered out


def filter_out_non_date_rows(df, date_column):
    """Filter out rows where the date column is not a valid date."""
    df[date_column] = df[date_column].apply(date_filter_converter)
    df = df.dropna(subset=[date_column]).reset_index(drop=True)
    df[date_column] = pd.to_datetime(df[date_column], dayfirst=True).dt.date
    return df
