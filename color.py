from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def format_audit_excel(filename):
    # 1. Load the workbook and select the active sheet
    wb = load_workbook(filename)
    ws = wb.active

    # 2. Define the colors (using hex codes)
    # You can customize these hex colors as needed
    yellow_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )  # Internal Only
    red_fill = PatternFill(
        start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
    )  # Supplier Only
    green_fill = PatternFill(
        start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
    )  # Perfect Match

    # 3. Find which column contains the '_merge' indicator dynamically
    merge_col_index = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "_merge":
            merge_col_index = col
            break

    # If we found the _merge column, proceed with coloring
    if merge_col_index:
        # 4. Loop through all data rows (starting at row 2 to skip headers)
        for row in range(2, ws.max_row + 1):
            # Read the merge status for this specific row
            merge_status = ws.cell(row=row, column=merge_col_index).value

            # Determine which color to apply based on the status
            fill_to_apply = None
            if merge_status == "left_only":
                fill_to_apply = yellow_fill
            elif merge_status == "right_only":
                fill_to_apply = red_fill
            elif merge_status == "both":
                fill_to_apply = green_fill

            # 5. Apply the selected color to every cell in that entire row
            if fill_to_apply:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = fill_to_apply

    # 6. Save the formatted file
    wb.save(filename)
    ws.freeze_panes = 'A2'
    print(f"File successfully formatted and saved as {filename}")
