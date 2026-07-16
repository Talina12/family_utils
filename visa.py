import pandas as pd
import os
import sys
from openpyxl.utils import get_column_letter
from utils import filter_out_non_date_rows, resolve_files

# next task: add list of existing categories and contragents.
#           recognize multiple payments and refer to billing day of the
#           document. remove ענף column from result


def next_available_path(path):
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    n = 1
    while os.path.exists(f"{root}({n}){ext}"):
        n += 1
    return f"{root}({n}){ext}"


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
        date_col = result_data_frame.columns[0]
        file_df = filter_out_non_date_rows(file_df, date_col)
        file_df["Source file"] = os.path.basename(filepath)
        result_data_frame = pd.concat(
            [result_data_frame, file_df], ignore_index=True
        )

    result_data_frame = result_data_frame.sort_values(by=date_col).reset_index(
        drop=True
    )

    billing_col_index = result_data_frame.columns.get_loc("סכום\nחיוב") + 1
    for offset, column_name in enumerate(
        ["категория", "контрагент", "примечания"]
    ):
        result_data_frame.insert(
            billing_col_index + offset, column_name, None
        )

    output_path = next_available_path(
        f"{file_directory}/combined_visa_charges.xlsx")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_data_frame.to_excel(writer, index=False)
        date_col_index = result_data_frame.columns.get_loc(date_col) + 1
        worksheet = writer.sheets["Sheet1"]
        date_col_letter = get_column_letter(date_col_index)
        worksheet.column_dimensions[
            date_col_letter].number_format = "dd/mm/yyyy"
    print(f"Wrote {len(result_data_frame)} rows to {output_path}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
