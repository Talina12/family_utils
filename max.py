import pandas as pd
from utils import (filter_out_non_date_rows, next_available_path,
                   resolve_files, write_dated_excel)
import sys
import os


# next task same as visa

def read_sheet(sheet_name_to_read, file_path_to_read,
               result_data_frame_to_append):
    try:
        file_df = pd.read_excel(
                file_path_to_read, usecols="A:N", header=3,
                sheet_name=sheet_name_to_read)
    except ValueError:
        print(f"Sheet '{sheet_name_to_read}' not found in "
              f"{os.path.basename(file_path_to_read)}, skipping.")
        return result_data_frame_to_append
    date_col = file_df.columns[0]
    file_df = filter_out_non_date_rows(file_df, date_col)
    file_df["Source file"] = os.path.basename(file_path_to_read)
    result_data_frame_to_append = pd.concat([result_data_frame_to_append,
                                             file_df], ignore_index=True)
    return result_data_frame_to_append


def main():
    pattern = '*transaction-details_export*.xlsx'
    file_directory, files = resolve_files(pattern)
    if not files:
        print(f"No files with pattern '{pattern}' "
              f"found in the directory: {file_directory}")
        return

    result_data_frame = pd.DataFrame()
    for filepath in files:
        result_data_frame = read_sheet("עסקאות במועד החיוב", filepath,
                                       result_data_frame)
        result_data_frame = read_sheet('עסקאות חו"ל ומט"ח', filepath,
                                       result_data_frame)
    result_date_col = result_data_frame.columns[0]
    result_data_frame = result_data_frame.sort_values(
        by=result_date_col).reset_index(drop=True)

    billing_col_index = result_data_frame.columns.get_loc("תאריך חיוב") + 1
    for offset, column_name in enumerate(
        ["категория", "контрагент"]
    ):
        result_data_frame.insert(
            billing_col_index + offset, column_name, None
        )

    output_path = next_available_path(
        f"{file_directory}/combined_max_charges.xlsx")

    write_dated_excel(result_data_frame, output_path, result_date_col)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
