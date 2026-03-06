import pandas as pd
import os
from pathlib import Path

def reconcile_all(input_path: str, output_path: str = None):
    """
        Run FO-BO-GL reconciliations.
        input_path: path to the Excel file containing sheets FO Dump, BO Dump, GL Dump
        output_path: optional path to write the output Excel. If None, placed next to input file.
        Returns the path to the written output file (str).
        """
    file_path = Path(input_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    # If no output_path provided, write next to input file
    if output_path:
        out_path = Path(output_path)
    else:
        out_path = file_path.parent / "FOBO_Reconciliation_Output.xlsx"
    # Load the Excel file and read all three sheets
    fo_df = pd.read_excel(file_path, sheet_name="FO Dump", engine="openpyxl")
    bo_df = pd.read_excel(file_path, sheet_name="BO Dump", engine="openpyxl")
    gl_df = pd.read_excel(file_path, sheet_name="GL Dump", engine="openpyxl")

    # Set 'Link Ref.' as index for FO and BO
    fo_df.set_index("Link Ref.", inplace=True)
    bo_df.set_index("Link Ref.", inplace=True)
    gl_df.set_index("Link Ref.", inplace=True)

    # ------------------ FO vs BO Reconciliation ------------------
    # Get all unique Link Ref. values from both sheets
    all_links = fo_df.index.union(bo_df.index)

    # Reindex both dataframes to include all Link Ref. values
    fo_all = fo_df.reindex(all_links)
    bo_all = bo_df.reindex(all_links)

    # Find common columns between the two dataframes
    common_columns = fo_all.columns.intersection(bo_all.columns)

    # Prepare prefixed dataframes for comparison
    fo_common_prefixed = fo_all[common_columns].add_prefix("FO_")
    bo_common_prefixed = bo_all[common_columns].add_prefix("BO_")

    # Combine both prefixed dataframes
    fo_bo_combined = pd.concat([fo_common_prefixed, bo_common_prefixed], axis=1)

    # Determine reconciliation status and remarks
    def reconcile_row(row):
        remarks = []
        missing_fo = row[[f"FO_{col}" for col in common_columns]].isnull().all()
        missing_bo = row[[f"BO_{col}" for col in common_columns]].isnull().all()

        if missing_fo or missing_bo:
            return pd.Series(["Not Reconciled", "Missing Value"])

        mismatched_columns = []
        for col in common_columns:
            fo_val = row.get(f"FO_{col}")
            bo_val = row.get(f"BO_{col}")
            if pd.isnull(fo_val) and pd.isnull(bo_val):
                continue
            if fo_val != bo_val:
                mismatched_columns.append(col)

        if mismatched_columns:
            return pd.Series(["Not Reconciled", ", ".join(mismatched_columns)])
        else:
            return pd.Series(["Reconciled", ""])

    # Apply reconciliation logic
    fo_bo_combined[["Reconciliation Status", "Remarks"]] = fo_bo_combined.apply(reconcile_row, axis=1)

    # Reset index to include 'Link Ref.' in the output
    fo_bo_combined.reset_index(inplace=True)

    # ------------------ FO vs GL Reconciliation ------------------
    fo_gl_links = fo_df.index.union(gl_df.index)
    fo_gl_fo = fo_df.reindex(fo_gl_links)
    fo_gl_gl = gl_df.reindex(fo_gl_links)

    fo_gl_combined = pd.DataFrame(index=fo_gl_links)
    fo_gl_combined["FO_CCY1"] = fo_gl_fo["CCY1"]
    fo_gl_combined["FO_Amount1"] = fo_gl_fo["Amount1"]
    fo_gl_combined["GL_Ccy"] = fo_gl_gl["Ccy"]
    fo_gl_combined["GL_Atlas Value"] = fo_gl_gl["Atlas Value"]

    def reconcile_fo_gl(row):
        missing_fo = pd.isnull(row["FO_CCY1"]) and pd.isnull(row["FO_Amount1"])
        missing_gl = pd.isnull(row["GL_Ccy"]) and pd.isnull(row["GL_Atlas Value"])
        if missing_fo or missing_gl:
            return pd.Series(["Not Reconciled", "Missing Value"])
        mismatches = []
        if row["FO_CCY1"] != row["GL_Ccy"]:
            mismatches.append("CCY1 vs Ccy")
        if row["FO_Amount1"] != row["GL_Atlas Value"]:
            mismatches.append("Amount1 vs Atlas Value")
        if mismatches:
            return pd.Series(["Not Reconciled", ", ".join(mismatches)])
        return pd.Series(["Reconciled", ""])

    fo_gl_combined[["Reconciliation Status", "Remarks"]] = fo_gl_combined.apply(reconcile_fo_gl, axis=1)
    fo_gl_combined.reset_index(inplace=True)

    # ------------------ BO vs GL Reconciliation ------------------
    # bo_df["System ID"] = bo_df["System ID"].astype(str)
    # gl_df["Ref."] = gl_df["Ref."].astype(str)

    # bo_gl_links = pd.Index(bo_df["System ID"].dropna().unique()).union(gl_df["Ref."].dropna().unique())
    # bo_gl_bo = bo_df.set_index("System ID").reindex(bo_gl_links)
    # gl_gl_gl = gl_df.set_index("Ref.").reindex(bo_gl_links)
    bo_gl_links = bo_df.index.union(gl_df.index)
    bo_gl_bo = bo_df.reindex(bo_gl_links)
    bo_gl_gl = gl_df.reindex(bo_gl_links)


    bo_gl_combined = pd.DataFrame(index=bo_gl_links)

    bo_gl_combined["BO_System ID"] = bo_gl_bo["System ID"]
    bo_gl_combined["BO_CCY1"] = bo_gl_bo["CCY1"]
    bo_gl_combined["BO_Amount1"] = bo_gl_bo["Amount1"]
    bo_gl_combined["GL_Ref."] = bo_gl_gl["Ref."]
    bo_gl_combined["GL_Ccy"] = bo_gl_gl["Ccy"]
    bo_gl_combined["GL_Atlas Value"] = bo_gl_gl["Atlas Value"]

    def reconcile_bo_gl(row):
        missing_bo = pd.isnull(row["BO_CCY1"]) and pd.isnull(row["BO_Amount1"])
        missing_gl = pd.isnull(row["GL_Ccy"]) and pd.isnull(row["GL_Atlas Value"])
        if missing_bo or missing_gl:
            return pd.Series(["Not Reconciled", "Missing Value"])
        mismatches = []
        if row["BO_CCY1"] != row["GL_Ccy"]:
            mismatches.append("CCY1 vs Ccy")
        if row["BO_Amount1"] != row["GL_Atlas Value"]:
            mismatches.append("Amount1 vs Atlas Value")
        if row["BO_System ID"] != row["GL_Ref."]:
            mismatches.append("System ID vs Ref.")
        if mismatches:
            return pd.Series(["Not Reconciled", ", ".join(mismatches)])
        return pd.Series(["Reconciled", ""])

    bo_gl_combined[["Reconciliation Status", "Remarks"]] = bo_gl_combined.apply(reconcile_bo_gl, axis=1)
    bo_gl_combined.reset_index(inplace=True)

    # ------------------ Save to Excel ------------------
    with pd.ExcelWriter("FOBO_Reconciliation_Output.xlsx", engine="openpyxl") as writer:
        fo_bo_combined.to_excel(writer, sheet_name="FO vs BO", index=False)
        fo_gl_combined.to_excel(writer, sheet_name="FO vs GL", index=False)
        bo_gl_combined.to_excel(writer, sheet_name="BO vs GL", index=False)

def main(input_file: str = None, output_file: str = None):
    """
    Simple CLI entrypoint.
    If input_file is None, it will look for 'Input/FOBO Dummy Data.xlsx' relative to cwd.
    """
    if input_file is None:
        input_file = os.path.join("Input", "FOBO Dummy Data.xlsx")
    result = reconcile_all(input_file, output_file)
    print(f"Wrote reconciliation output to: {result}")
    return result


if __name__ == "__main__":
    main()

# import pandas as pd

# # Load the Excel file and read all three sheets
# file_path = r"Input/FOBO Dummy Data.xlsx"
# fo_df = pd.read_excel(file_path, sheet_name="FO Dump", engine="openpyxl")
# bo_df = pd.read_excel(file_path, sheet_name="BO Dump", engine="openpyxl")
# gl_df = pd.read_excel(file_path, sheet_name="GL Dump", engine="openpyxl")

# # Set 'Link Ref.' as index for FO and BO
# fo_df.set_index("Link Ref.", inplace=True)
# bo_df.set_index("Link Ref.", inplace=True)
# gl_df.set_index("Link Ref.", inplace=True)

# # ------------------ FO vs BO Reconciliation ------------------
# # Get all unique Link Ref. values from both sheets
# all_links = fo_df.index.union(bo_df.index)

# # Reindex both dataframes to include all Link Ref. values
# fo_all = fo_df.reindex(all_links)
# bo_all = bo_df.reindex(all_links)

# # Find common columns between the two dataframes
# common_columns = fo_all.columns.intersection(bo_all.columns)

# # Prepare prefixed dataframes for comparison
# fo_common_prefixed = fo_all[common_columns].add_prefix("FO_")
# bo_common_prefixed = bo_all[common_columns].add_prefix("BO_")

# # Combine both prefixed dataframes
# fo_bo_combined = pd.concat([fo_common_prefixed, bo_common_prefixed], axis=1)

# # Determine reconciliation status and remarks
# def reconcile_row(row):
#     remarks = []
#     missing_fo = row[[f"FO_{col}" for col in common_columns]].isnull().all()
#     missing_bo = row[[f"BO_{col}" for col in common_columns]].isnull().all()

#     if missing_fo or missing_bo:
#         return pd.Series(["Not Reconciled", "Missing Value"])

#     mismatched_columns = []
#     for col in common_columns:
#         fo_val = row.get(f"FO_{col}")
#         bo_val = row.get(f"BO_{col}")
#         if pd.isnull(fo_val) and pd.isnull(bo_val):
#             continue
#         if fo_val != bo_val:
#             mismatched_columns.append(col)

#     if mismatched_columns:
#         return pd.Series(["Not Reconciled", ", ".join(mismatched_columns)])
#     else:
#         return pd.Series(["Reconciled", ""])

# # Apply reconciliation logic
# fo_bo_combined[["Reconciliation Status", "Remarks"]] = fo_bo_combined.apply(reconcile_row, axis=1)

# # Reset index to include 'Link Ref.' in the output
# fo_bo_combined.reset_index(inplace=True)

# # ------------------ FO vs GL Reconciliation ------------------
# fo_gl_links = fo_df.index.union(gl_df.index)
# fo_gl_fo = fo_df.reindex(fo_gl_links)
# fo_gl_gl = gl_df.reindex(fo_gl_links)

# fo_gl_combined = pd.DataFrame(index=fo_gl_links)
# fo_gl_combined["FO_CCY1"] = fo_gl_fo["CCY1"]
# fo_gl_combined["FO_Amount1"] = fo_gl_fo["Amount1"]
# fo_gl_combined["GL_Ccy"] = fo_gl_gl["Ccy"]
# fo_gl_combined["GL_Atlas Value"] = fo_gl_gl["Atlas Value"]

# def reconcile_fo_gl(row):
#     missing_fo = pd.isnull(row["FO_CCY1"]) and pd.isnull(row["FO_Amount1"])
#     missing_gl = pd.isnull(row["GL_Ccy"]) and pd.isnull(row["GL_Atlas Value"])
#     if missing_fo or missing_gl:
#         return pd.Series(["Not Reconciled", "Missing Value"])
#     mismatches = []
#     if row["FO_CCY1"] != row["GL_Ccy"]:
#         mismatches.append("CCY1 vs Ccy")
#     if row["FO_Amount1"] != row["GL_Atlas Value"]:
#         mismatches.append("Amount1 vs Atlas Value")
#     if mismatches:
#         return pd.Series(["Not Reconciled", ", ".join(mismatches)])
#     return pd.Series(["Reconciled", ""])

# fo_gl_combined[["Reconciliation Status", "Remarks"]] = fo_gl_combined.apply(reconcile_fo_gl, axis=1)
# fo_gl_combined.reset_index(inplace=True)

# # ------------------ BO vs GL Reconciliation ------------------
# # bo_df["System ID"] = bo_df["System ID"].astype(str)
# # gl_df["Ref."] = gl_df["Ref."].astype(str)

# # bo_gl_links = pd.Index(bo_df["System ID"].dropna().unique()).union(gl_df["Ref."].dropna().unique())
# # bo_gl_bo = bo_df.set_index("System ID").reindex(bo_gl_links)
# # gl_gl_gl = gl_df.set_index("Ref.").reindex(bo_gl_links)
# bo_gl_links = bo_df.index.union(gl_df.index)
# bo_gl_bo = bo_df.reindex(bo_gl_links)
# bo_gl_gl = gl_df.reindex(bo_gl_links)


# bo_gl_combined = pd.DataFrame(index=bo_gl_links)

# bo_gl_combined["BO_System ID"] = bo_gl_bo["System ID"]
# bo_gl_combined["BO_CCY1"] = bo_gl_bo["CCY1"]
# bo_gl_combined["BO_Amount1"] = bo_gl_bo["Amount1"]
# bo_gl_combined["GL_Ref."] = bo_gl_gl["Ref."]
# bo_gl_combined["GL_Ccy"] = bo_gl_gl["Ccy"]
# bo_gl_combined["GL_Atlas Value"] = bo_gl_gl["Atlas Value"]

# def reconcile_bo_gl(row):
#     missing_bo = pd.isnull(row["BO_CCY1"]) and pd.isnull(row["BO_Amount1"])
#     missing_gl = pd.isnull(row["GL_Ccy"]) and pd.isnull(row["GL_Atlas Value"])
#     if missing_bo or missing_gl:
#         return pd.Series(["Not Reconciled", "Missing Value"])
#     mismatches = []
#     if row["BO_CCY1"] != row["GL_Ccy"]:
#         mismatches.append("CCY1 vs Ccy")
#     if row["BO_Amount1"] != row["GL_Atlas Value"]:
#         mismatches.append("Amount1 vs Atlas Value")
#     if row["BO_System ID"] != row["GL_Ref."]:
#         mismatches.append("System ID vs Ref.")
#     if mismatches:
#         return pd.Series(["Not Reconciled", ", ".join(mismatches)])
#     return pd.Series(["Reconciled", ""])

# bo_gl_combined[["Reconciliation Status", "Remarks"]] = bo_gl_combined.apply(reconcile_bo_gl, axis=1)
# bo_gl_combined.reset_index(inplace=True)

# # ------------------ Save to Excel ------------------
# with pd.ExcelWriter("FOBO_Reconciliation_Output.xlsx", engine="openpyxl") as writer:
#     fo_bo_combined.to_excel(writer, sheet_name="FO vs BO", index=False)
#     fo_gl_combined.to_excel(writer, sheet_name="FO vs GL", index=False)
#     bo_gl_combined.to_excel(writer, sheet_name="BO vs GL", index=False)
