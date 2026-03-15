import pandas as pd

from edit import clean_transaction_data as ctd

clean_df_a = ctd(file_path="./sheets/CARE.xlsx", header_row_index=11, anchor_column="Vch No.")

# File B: 23 rows of top metadata. The anchor is 'Document Number'.
# This instantly removes 'Period Total', 'Cumulative Total', and fragmented 'Accounted' rows.
clean_df_b = ctd(
    file_path="./sheets/jiva.xlsx",
    header_row_index=23,
    anchor_column="Document Number",
)


# Mapping for your internal ledger (Care Banjara)
company_mapping = {
    "Vch No.": "Invoice_Number",
    "Date": "Transaction_Date",
    "Particulars": "Description",
}

# Mapping for the supplier statement
supplier_mapping = {
    "Document Number": "Invoice_Number",
    "Document Date": "Transaction_Date",
    # 'Description' is already named correctly in the supplier file,
    # but we map it anyway just to be safe if it changes.
    "Description": "Description",
}

# This converts all column names to uppercase and replaces spaces with underscores
# Example: "Invoice Number " becomes "INVOICE_NUMBER"
clean_df_a.columns = clean_df_a.columns.str.strip().str.upper().str.replace(" ", "_")
clean_df_b.columns = clean_df_b.columns.str.strip().str.upper().str.replace(" ", "_")

# Print the structural information of your dataframe
print(clean_df_a.info())

# Print the structural information of your dataframe
print(clean_df_b.info())
