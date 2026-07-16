import pandas as pd
import os
import sys
from utils import (
    filter_out_non_date_rows,
    next_available_path,
    resolve_files,
    write_dated_excel,
)

# next task: add list of existing categories and contragents.
#           recognize multiple payments and refer to billing day of the
#           document. remove ענף column from result


def main():
    pattern = '*פירוט חיובים לכרטיס דיינרס מסטרקארד*.xlsx'
    file_directory, files = resolve_files(pattern)
    if not files:
        print(f"No files with pattern '{pattern}' "
              f"found in the directory: {file_directory}")
        return

    result_data_frame = pd.DataFrame()
    for filepath in files:
        file_df = pd.read_excel(
            filepath, usecols="A:G", header=3
        )
        date_col = file_df.columns[0]
        file_df = filter_out_non_date_rows(file_df, date_col)
        file_df["Source file"] = os.path.basename(filepath)
        result_data_frame = pd.concat(
            [result_data_frame, file_df], ignore_index=True
        )
    result_date_col = result_data_frame.columns[0]
    result_data_frame = result_data_frame.sort_values(
        by=result_date_col).reset_index(drop=True)

    billing_col_index = result_data_frame.columns.get_loc("סכום\nחיוב") + 1
    for offset, column_name in enumerate(["категория", "контрагент"]):
        result_data_frame.insert(
            billing_col_index + offset, column_name, None
        )

    output_path = next_available_path(
        f"{file_directory}/combined_visa_charges.xlsx")

    write_dated_excel(result_data_frame, output_path, result_date_col)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
