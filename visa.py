import pandas as pd
import glob
import os
import sys
import datetime
import re
from openpyxl.utils import get_column_letter
# next task: add list of existing categories and contragents.
#           recognize multiple payments and refer to billing day of the document
#           remove ענף column from result


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


def next_available_path(path):
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    n = 1
    while os.path.exists(f"{root}({n}){ext}"):
        n += 1
    return f"{root}({n}){ext}"


def main():
    default_directory = "C:/Users/talin/Downloads"
    file_directory = sys.argv[1] if len(sys.argv) > 1 else default_directory
    print(f"Using file directory: {file_directory}")
    files = glob.glob((file_directory + '/*פירוט חיובים לכרטיס דיינרס מסטרקארד*.xlsx'))
    if not files:
        print(f"No files found in the directory: {file_directory}")
        return

    result_data_frame = pd.DataFrame()
    for filepath in files:
        file_df = pd.read_excel(
            filepath, usecols="A:G", header=3,
            # names=["Date", "Business name", "Transaction amount",
            #        "Billing amount", "Transaction type", "Industry",
            #        "Description"],
            dtype={"Transaction amount": float, "Billing amount": float}
        )
        # print(file_df.info)
        # print(file_df.dtypes)
        # amount_col = file_df.columns[2]
        # temp = file_df[amount_col]
        date_col = file_df.columns[0]
        file_df[date_col] = file_df[date_col].apply(date_filter_converter)
        file_df = file_df.dropna(subset=[date_col]).reset_index(drop=True)
        file_df[date_col] = pd.to_datetime(
            file_df[date_col]
        ).dt.date
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

    output_path = next_available_path(f"{file_directory}/combined_visa_charges.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        result_data_frame.to_excel(writer, index=False)
        date_col_index = result_data_frame.columns.get_loc(date_col) + 1
        worksheet = writer.sheets["Sheet1"]
        date_col_letter = get_column_letter(date_col_index)
        worksheet.column_dimensions[date_col_letter].number_format = "dd/mm/yyyy"
    print(f"Wrote {len(result_data_frame)} rows to {output_path}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
