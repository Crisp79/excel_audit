import openpyxl
from openpyxl.styles import PatternFill


def format_audit_excel(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Define the colors
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Map headers to column indices
    headers = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value}

    # Helper function to find columns with flexible naming
    def find_col(possible_names):
        for name in possible_names:
            if name in headers:
                return headers[name]
        return None

    # Identify indices (Supports standard names and merge suffixes)
    m_idx = find_col(["_merge"])
    v_idx = find_col(["VARIATION"])

    # Left-side columns (a)
    d_a_idx = find_col(["DEBIT_a", "DEBIT"])

    # Right-side columns (b)
    c_b_idx = find_col(["CREDIT_b", "CREDIT"])

    for row in ws.iter_rows(min_row=2):
        # Extract values (using 0 if column is missing to avoid errors)
        m_val = row[m_idx - 1].value if m_idx else None
        variation = row[v_idx - 1].value if v_idx else 0

        debit_a = row[d_a_idx - 1].value if d_a_idx else 0
        credit_b = row[c_b_idx - 1].value if c_b_idx else 0

        row_fill = None

        # Logic based on merge indicators
        if m_val == "both":
            if variation == 0:
                row_fill = green
            else:
                row_fill = yellow

        elif m_val == "left_only":
            # Check both debit and credit for the left side
            if debit_a and debit_a != 0:
                row_fill = red

        elif m_val == "right_only":
            # Check both debit and credit for the right side
            if credit_b and credit_b != 0:
                row_fill = red

        # Apply the color to the row
        if row_fill:
            for cell in row:
                cell.fill = row_fill

    # Save to a new file to preserve the original
    ws.freeze_panes = "A2"
    ws.delete_cols(m_idx)
    wb.save(filename)
    print(f"File successfully formatted and saved as {filename}")
