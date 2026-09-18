import pandas as pd
import os
import sys
from categories import get_category_store
from utils import (
    filter_out_non_date_rows,
    find_header_row,
    next_available_path,
    resolve_files,
    write_dated_excel,
)

# next task: recognize multiple payments and refer to billing day of the
#           document. remove ענף column from result


def main():
    pattern = '[0-9][0-9][0-9][0-9]_[0-9][0-9]_[0-9][0-9][0-9][0-9].xlsx'
    file_directory, files = resolve_files(pattern)
    if not files:
        print(f"No files with pattern '{pattern}' "
              f"found in the directory: {file_directory}")
        return

    result_data_frame = pd.DataFrame()
    for filepath in files:
        # The sections above "עסקאות למועד חיוב" vary between statements
        # (e.g. a narrower "עסקאות שטרם נקלטו" table), so locate the
        # first header row that has the billing columns instead of assuming
        # a fixed row number.
        header_row = find_header_row(filepath, "סכום חיוב", usecols="A:H")
        file_df = pd.read_excel(
            filepath, usecols="A:H", header=header_row,
            # Voucher numbers are IDs, not amounts: keep the leading zeros.
            dtype={"מס' שובר": str},
        )
        date_col = file_df.columns[0]
        file_df = filter_out_non_date_rows(file_df, date_col)
        file_df["Source file"] = os.path.basename(filepath)
        # print(file_df.info())
        result_data_frame = pd.concat(
            [result_data_frame, file_df], ignore_index=True
        )
    # print(result_data_frame.index)
    result_date_col = result_data_frame.columns[0]
    result_data_frame = result_data_frame.sort_values(
        by=result_date_col).reset_index(drop=True)

    billing_col_index = result_data_frame.columns.get_loc("מטבע חיוב") + 1
    category_store = get_category_store()
    categories = result_data_frame["שם בית עסק"].apply(
        category_store.categorize)
    contragents = result_data_frame["שם בית עסק"].apply(
        category_store.contragent)
    result_data_frame.insert(billing_col_index, "категория", categories)
    result_data_frame.insert(billing_col_index + 1, "контрагент", contragents)

    output_path = next_available_path(
        f"{file_directory}/combined_isracard_charges.xlsx")

    write_dated_excel(result_data_frame, output_path, result_date_col)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
