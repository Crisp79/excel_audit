import pandas as pd


def clean_transaction_data(file_path, header_row_index, anchor_column):
    # 1. Load data starting from the actual header row
    all_sheets = pd.read_excel(file_path, skiprows=header_row_index, sheet_name=None)
    # TODO : skip empty sheets or say to empty sheets
    cleaned_sheets = []

    for name, df in all_sheets.items():
        # Perform cleaning on each individual dataframe
        df = df.dropna(how="all")
        # ... (rest of your cleaning logic) ...
        cleaned_sheets.append(df)

    # Combine them all into one big DataFrame

    # 2. Drop completely empty rows
    df = df.dropna(how="all")

    # 3. Remove repeated column headers
    # This filters out rows where the anchor column contains its own name
    df = df[df[anchor_column] != anchor_column]

    # 4. The Anchor Strategy: Drop all breaks, sub-totals, and partial rows
    # Any row missing a valid ID in the anchor column is dropped automatically
    df = df.dropna(subset=[anchor_column])

    # 5. Clean up any whitespace in text columns that might interfere with comparison
    # Only target 'object' (mixed/text) columns
    for col in df.select_dtypes(include=["object"]).columns:
        # Check each cell: strip if it's a string, otherwise leave it alone
        df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
    # 6. Apply specific business filters (Example: Removing cancelled transactions)
    if "Status" in df.columns:
        df = df[df["Status"] != "Cancelled"]

    # Reset index for a clean, sequential dataframe
    return df.reset_index(drop=True)
