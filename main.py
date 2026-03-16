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

key_1 = key_1.strip().upper().replace(" ", "_")
key_2 = key_2.strip().upper().replace(" ", "_")

clean_df_a.columns = clean_df_a.columns.str.strip().str.upper().str.replace(" ", "_")
clean_df_b.columns = clean_df_b.columns.str.strip().str.upper().str.replace(" ", "_")

clean_df_a["DOC"] = "doc1"
clean_df_b["DOC"] = "doc2"

clean_df_a[key_1] = clean_df_a[key_1].astype(str)
clean_df_b[key_2] = clean_df_b[key_2].astype(str)

clean_df_a["DEBIT"] = clean_df_a["DEBIT"].fillna(0)
clean_df_a["CREDIT"] = clean_df_a["CREDIT"].fillna(0)
clean_df_b["DEBIT"] = clean_df_b["DEBIT"].fillna(0)
clean_df_b["CREDIT"] = clean_df_b["CREDIT"].fillna(0)

clean_df_a = clean_df_a.add_suffix("_A")
clean_df_b = clean_df_b.add_suffix("_B")
key_1 = key_1 + "_A"
key_2 = key_2 + "_B"

print(clean_df_a.info())
print(clean_df_b.info())


merged_df = pd.merge(
    clean_df_a,
    clean_df_b,
    left_on=key_1,
    right_on=key_2,
    how="outer",
    indicator=True,
)
print(merged_df.info())

merged_df_left = merged_df[merged_df["_merge"] == "left_only"]
merged_df_right = merged_df[merged_df["_merge"] == "right_only"]
merged_df_inner = merged_df[merged_df["_merge"] == "both"]

merged_df["VARIATION"] = (
    merged_df["DEBIT_A"]
    + merged_df["DEBIT_B"]
    - merged_df["CREDIT_A"]
    - merged_df["CREDIT_B"]
)

final_df_a = merged_df[merged_df["DOC_A"] == "doc1"]
final_df_b = merged_df[merged_df["DOC_B"] == "doc2"]

final_df_a = final_df_a.drop(
    columns=[col for col in final_df_a.columns if col.endswith("_B")]
)
final_df_b = final_df_b.drop(
    columns=[col for col in final_df_b.columns if col.endswith("_A")]
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
