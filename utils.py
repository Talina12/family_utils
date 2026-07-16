import datetime
import glob
import os
import re
import sys

import pandas as pd


def resolve_files(pattern, default_directory="C:/Users/talin/Downloads"):
    file_directory = sys.argv[1] if len(sys.argv) > 1 else default_directory
    print(f"Using file directory: {file_directory}")
    files = glob.glob(os.path.join(file_directory, pattern))
    return file_directory, files


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
    ]

    if any(re.match(pattern, str_val) for pattern in date_patterns):
        return value
    else:
        return None  # This will become NaN and can be filtered out


def filter_out_non_date_rows(df, date_column):
    """Filter out rows where the date column is not a valid date."""
    df[date_column] = df[date_column].apply(date_filter_converter)
    df = df.dropna(subset=[date_column]).reset_index(drop=True)
    df[date_column] = pd.to_datetime(df[date_column]).dt.date
    return df

