import pandas as pd
from pandas.tseries.frequencies import key

from color import format_audit_excel as fae
from edit import clean_transaction_data as ctd

filepath1 = "./sheets/CARE.xlsx"
head_row1 = 12
key_1 = "Vch No."

filepath2 = "./sheets/jiva.xlsx"
head_row2 = 24
key_2 = "Document Number"

clean_df_a = ctd(
    file_path=filepath1, header_row_index=head_row1 - 1, anchor_column=key_1
)
clean_df_b = ctd(
    file_path=filepath2, header_row_index=head_row2 - 1, anchor_column=key_2
)

key_common = "DOCUMENT_NUMBER"
key_1 = key_1.strip().upper().replace(" ", "_")
key_2 = key_2.strip().upper().replace(" ", "_")

clean_df_a.columns = clean_df_a.columns.str.strip().str.upper().str.replace(" ", "_")
clean_df_b.columns = clean_df_b.columns.str.strip().str.upper().str.replace(" ", "_")

clean_df_a[key_common] = clean_df_a[key_1]
clean_df_b[key_common] = clean_df_b[key_2]

clean_df_a["DOC"] = "doc1"
clean_df_b["DOC"] = "doc2"

clean_df_a[key_common] = clean_df_a[key_common].astype(str)
clean_df_b[key_common] = clean_df_b[key_common].astype(str)

clean_df_a["DEBIT"] = clean_df_a["DEBIT"].fillna(0)
clean_df_a["CREDIT"] = clean_df_a["CREDIT"].fillna(0)

print(clean_df_a.info())
print(clean_df_b.info())


merged_df = pd.merge(
    clean_df_a,
    clean_df_b,
    on=key_common,
    how="outer",
    suffixes=("_A", "_B"),
    indicator=True,
)
print(merged_df.info())

merged_df_left = merged_df[merged_df["_merge"] == "left_only"]
merged_df_right = merged_df[merged_df["_merge"] == "right_only"]
merged_df_inner = merged_df[merged_df["_merge"] == "both"]

final_df_a = merged_df[merged_df["DOC_A"] == "doc1"]
final_df_b = merged_df[merged_df["DOC_B"] == "doc2"]

merged_df_inner["VARIATION"] = (
    merged_df_inner["DEBIT_A"]
    + merged_df_inner["DEBIT_B"]
    - merged_df_inner["CREDIT_A"]
    - merged_df_inner["CREDIT_B"]
)

merged_df_inner.to_excel("output/output_inner.xlsx")
merged_df_left.to_excel("output/output_left.xlsx")
merged_df_right.to_excel("output/output_right.xlsx")
merged_df.to_excel("output/output_outer.xlsx")
final_df_a.to_excel("output/output_a.xlsx")
final_df_b.to_excel("output/output_b.xlsx")

# Run the formatting function
fae("output/output_outer.xlsx")
fae("output/output_inner.xlsx")
fae("output/output_left.xlsx")
fae("output/output_right.xlsx")
fae("output/output_a.xlsx")
fae("output/output_b.xlsx")
