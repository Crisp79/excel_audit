# Copyright (c) 2026 Akshith Vuppala
# Licensed under the MIT License.
# See LICENSE file in the project root for full license information.

import pandas as pd

from color import format_and_split_dataframe as fsd
from edit import clean_transaction_data as ctd


def audit_excel(filepath1, head_row1, key_1, filepath2, head_row2, key_2, out_put):
    print("Starting Process")
    clean_df_a = ctd(
        file_path=filepath1, header_row_index=head_row1 - 1, anchor_column=key_1
    )
    clean_df_b = ctd(
        file_path=filepath2, header_row_index=head_row2 - 1, anchor_column=key_2
    )

    key_1 = key_1.strip().upper().replace(" ", "_")
    key_2 = key_2.strip().upper().replace(" ", "_")

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
    print(clean_df_a.head())
    print(clean_df_b.info())
    print(clean_df_b.head())

    merged_df = pd.merge(
        clean_df_a,
        clean_df_b,
        left_on=key_1,
        right_on=key_2,
        how="outer",
        indicator=True,
    )
    print(merged_df.info())

    merged_df["VARIATION"] = (
        merged_df["DEBIT_A"].abs()
        + merged_df["DEBIT_B"].abs()
        - merged_df["CREDIT_A"].abs()
        - merged_df["CREDIT_B"].abs()
    ).round(2)
    final_df_a = merged_df[merged_df["_merge"].isin(["left_only", "both"])]
    final_df_b = merged_df[merged_df["_merge"].isin(["right_only", "both"])]

    final_df_a = final_df_a.drop(
        columns=[col for col in final_df_a.columns if col.endswith("_B")]
    )
    final_df_b = final_df_b.drop(
        columns=[col for col in final_df_b.columns if col.endswith("_A")]
    )

    final_df_a.columns = final_df_a.columns.str.removesuffix("_A")
    final_df_b.columns = final_df_b.columns.str.removesuffix("_B")
    output_a1 = out_put.removesuffix(".xlsx") + "_buyer.xlsx"
    output_b1 = out_put.removesuffix(".xlsx") + "_seller.xlsx"
    output_outer = out_put.removesuffix(".xlsx") + "_outer.xlsx"
    fsd(filepath1, final_df_a, output_a1)
    fsd(filepath2, final_df_b, output_b1)

    merged_df.to_excel(output_outer)
    print("processing done")


if __name__ == "__main__":
    filepath1 = "./sheets/CARE.xlsx"
    head_row1 = 12
    key_1 = "Vch No."

    filepath2 = "./sheets/jiva.xlsx"
    head_row2 = 24
    key_2 = "Document Number"

    out_put = "output/output.xlsx"
    audit_excel(filepath1, head_row1, key_1, filepath2, head_row2, key_2, out_put)
