import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows


def format_and_split_dataframe(df, output_filename):
    # 1. Initialize a new workbook
    wb = Workbook()
    # Remove the default sheet
    del wb["Sheet"]

    # 2. Define the colors
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # 3. Get column indices from DataFrame (0-based)
    # We use .get_loc to find positions safely
    try:
        m_idx = df.columns.get_loc("_merge")
        v_idx = df.columns.get_loc("VARIATION")
        # Handle flexible names for debit/credit
        d_a_idx = next(
            (i for i, col in enumerate(df.columns) if col in ["DEBIT_a", "DEBIT"]), None
        )
        c_b_idx = next(
            (i for i, col in enumerate(df.columns) if col in ["CREDIT_b", "CREDIT"]),
            None,
        )
        sheet_col_idx = df.columns.get_loc("SHEET")
    except KeyError as e:
        print(f"Error: Missing required column {e}")
        return

    # 4. Group data by the 'SHEET' column
    for sheet_name, group in df.groupby("SHEET"):
        # Clean sheet name for Excel
        clean_name = "".join(c for c in str(sheet_name) if c not in r"\*?:/[]")[:31]
        ws = wb.create_sheet(title=clean_name)

        # 5. Write DataFrame to the sheet
        # index=False, header=True
        for r_idx, row_data in enumerate(
            dataframe_to_rows(group, index=False, header=True), 1
        ):
            ws.append(row_data)

            # Skip the header row (index 1) for coloring logic
            if r_idx == 1:
                continue

            # 6. Apply your Audit Logic
            # Note: dataframe_to_rows yields row data as a list/tuple
            # row_data corresponds to the values in the current row
            m_val = row_data[m_idx] if m_idx is not None else None
            variation = row_data[v_idx] if v_idx is not None else 0
            debit_a = row_data[d_a_idx] if d_a_idx is not None else 0
            credit_b = row_data[c_b_idx] if c_b_idx is not None else 0

            row_fill = None
            if m_val == "both":
                row_fill = green if variation == 0 else yellow
            elif m_val == "left_only" and debit_a and debit_a != 0:
                row_fill = red
            elif m_val == "right_only" and credit_b and credit_b != 0:
                row_fill = red

            if row_fill:
                for cell in ws[ws.max_row]:  # Apply to the row just appended
                    cell.fill = row_fill

        # 7. Final Cleanup for this sheet
        # Delete columns in reverse order to maintain index integrity
        cols_to_delete = sorted(
            [i + 1 for i in [m_idx, sheet_col_idx] if i is not None], reverse=True
        )
        for col_idx in cols_to_delete:
            ws.delete_cols(col_idx)

        ws.freeze_panes = "A2"

    wb.save(output_filename)
    print(f"DataFrame successfully split and saved to {output_filename}")


# Usage:
# format_and_split_dataframe(your_dataframe, "Final_Audit.xlsx")
