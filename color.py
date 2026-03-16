from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def format_audit_excel(filename):
    wb = load_workbook(filename)
    ws = wb.active

    # 1. Define colors
    # Standard Red for mismatches/missing
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    # Yellow for rows that match but have a value difference
    yellow_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )

    # 2. Find column indices dynamically
    merge_col = None
    var_col = None

    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col).value
        if header == "_merge":
            merge_col = col
        elif header == "VARIATION":
            var_col = col

    # 3. Apply Conditional Formatting Logic
    if merge_col:
        for row_idx in range(2, ws.max_row + 1):
            merge_status = ws.cell(row=row_idx, column=merge_col).value

            # Initialize fill as None (Default/No color)
            fill_to_apply = None

            # Logic A: If it's only in one sheet, mark Red
            if merge_status in ["left_only", "right_only"]:
                fill_to_apply = red_fill

            # Logic B: If it's in both, check the variation
            elif merge_status == "both" and var_col:
                variation_val = ws.cell(row=row_idx, column=var_col).value
                # Check if variation is not zero (handling None/NaN as well)
                if variation_val is not None and variation_val != 0:
                    fill_to_apply = yellow_fill
                else:
                    # Match with 0 variation -> stays default color
                    fill_to_apply = None

            # 4. Apply the fill to the row
            if fill_to_apply:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill_to_apply

    # 5. Settings and Save
    ws.freeze_panes = "A2"
    wb.save(filename)
    print(f"File successfully formatted and saved as {filename}")
