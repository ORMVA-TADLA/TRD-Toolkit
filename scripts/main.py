import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import TableColumn
import argparse
import sys


def process_excel_table(
    filepath, target_table_name, group_col, val_col, target_col, output_col
):
    print(f"\nOpening '{filepath}' and searching for table '{target_table_name}'...")

    try:
        wb = openpyxl.load_workbook(filepath)
    except FileNotFoundError:
        print(f"\n[ERROR] The file '{filepath}' could not be found.")
        sys.exit(1)

    # 1. Search all sheets for the specific Table name
    target_ws = None
    target_table = None
    for ws in wb.worksheets:
        if target_table_name in ws.tables:
            target_ws = ws
            target_table = ws.tables[target_table_name]
            break

    if not target_table:
        print(
            f"\n[ERROR] Could not find a table named '{target_table_name}' in this workbook."
        )
        print("Please check the configuration and ensure the table name is correct.")
        sys.exit(1)

    print(f"Found Table: '{target_table.name}'. Reading specified columns...")

    # 2. Get table boundaries and read headers
    min_col, min_row, max_col, max_row = range_boundaries(target_table.ref)
    headers = [
        target_ws.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)
    ]

    # 3. Find the column indexes based on the fixed names
    try:
        group_idx = headers.index(group_col)
        val_idx = headers.index(val_col)
        target_idx = headers.index(target_col)
    except ValueError as e:
        print(
            f"\n[ERROR] Could not find one of the specified columns in the table headers."
        )
        print(f"Expected: '{group_col}', '{val_col}', or '{target_col}'")
        print(f"Headers found in Excel: {headers}")
        sys.exit(1)

    # 4. Extract the exact rows of data into a dictionary
    data = []
    for r in range(min_row + 1, max_row + 1):
        data.append(
            {
                "excel_row": r,  # Save the row number so we know exactly where to write back
                "group": target_ws.cell(row=r, column=min_col + group_idx).value,
                "val": target_ws.cell(row=r, column=min_col + val_idx).value,
                "target": target_ws.cell(row=r, column=min_col + target_idx).value,
            }
        )

    df = pd.DataFrame(data)

    # 5. Math logic for each group
    def apply_math(group_df):
        # 1. Force the target sum to be a clean decimal (float)
        target_sum = float(group_df["target"].iloc[0])

        # 2. Force the values column to be numbers.
        # (errors='coerce' turns text/blanks into NaN, fillna(0) turns NaN into 0)
        clean_vals = pd.to_numeric(group_df["val"], errors="coerce").fillna(0)
        orig_nums = clean_vals.values

        orig_sum = np.sum(orig_nums)

        if orig_sum == 0:
            group_df["final"] = 0
            return group_df

        exact_vals = orig_nums * (target_sum / orig_sum)
        base_ints = np.floor(exact_vals).astype(int)
        remainders = exact_vals - base_ints
        shortfall = int(target_sum - np.sum(base_ints))

        if shortfall > 0:
            largest_indices = np.argsort(remainders)[-shortfall:]
            for idx in largest_indices:
                base_ints[idx] += 1

        group_df["final"] = base_ints
        return group_df

    # Run the math
    result_df = df.groupby("group", group_keys=False).apply(apply_math)

    # 6. Write the data directly back into the Excel file
    print("Calculating and writing data back to the file...")
    new_col_idx = max_col + 1

    # Write the new Header
    target_ws.cell(row=min_row, column=new_col_idx).value = output_col

    # Write the final integers down the new column
    for index, row in result_df.iterrows():
        target_ws.cell(row=row["excel_row"], column=new_col_idx).value = row["final"]

    # 7. Update the Table Object settings to include the new column
    new_col_letter = get_column_letter(new_col_idx)
    start_cell = target_table.ref.split(":")[0]

    # Expand the table boundaries
    target_table.ref = f"{start_cell}:{new_col_letter}{max_row}"

    # Tell Excel a new column exists so it doesn't throw a "corrupt table" error
    new_table_column = TableColumn(
        id=len(target_table.tableColumns) + 1, name=output_col
    )
    target_table.tableColumns.append(new_table_column)

    # Save and overwrite the original file
    wb.save(filepath)
    print(
        f"\n[SUCCESS] Added column '{output_col}' to table '{target_table.name}' in {filepath}"
    )


if __name__ == "__main__":
    # ==========================================
    # CONFIGURATION: Set your fixed table and column names here
    # ==========================================
    TABLE_NAME = "PC"  # Replace with your exact Excel Table name
    GROUP_COL_NAME = "Ref"  # Replace with your exact header name
    VAL_COL_NAME = "Hours static"  # Replace with your exact header name
    TARGET_COL_NAME = "Target Sum"  # Replace with your exact header name
    OUTPUT_COL_NAME = "Hours final"  # The name for the new column
    # ==========================================

    help_text = f"""
Reads an Excel Table Object, finds specific fixed columns by name, 
scales the numbers, and appends a newly created column to the table.

Fixed Configuration Expected:
  Table Name    : '{TABLE_NAME}'
  Group Column  : '{GROUP_COL_NAME}'
  Values Column : '{VAL_COL_NAME}'
  Target Column : '{TARGET_COL_NAME}'
  Output Column : '{OUTPUT_COL_NAME}'
    """

    parser = argparse.ArgumentParser(
        description=help_text,
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="Example usage: python scale_table.py my_data.xlsx",
    )

    # Using nargs="?" makes the filepath argument optional from the command line
    parser.add_argument(
        "filepath",
        nargs="?",
        help="Path to the .xlsx file (if omitted, you will be prompted)",
    )

    args = parser.parse_args()

    # Check if the user provided the filepath argument. If not, prompt them for it.
    target_filepath = args.filepath
    if not target_filepath:
        target_filepath = input("Please enter the path to your Excel file: ").strip()

        # If the user just presses Enter without typing anything, exit cleanly
        if not target_filepath:
            print("\n[ERROR] No file path provided. Exiting.")
            sys.exit(1)

    process_excel_table(
        filepath=target_filepath,
        target_table_name=TABLE_NAME,
        group_col=GROUP_COL_NAME,
        val_col=VAL_COL_NAME,
        target_col=TARGET_COL_NAME,
        output_col=OUTPUT_COL_NAME,
    )
