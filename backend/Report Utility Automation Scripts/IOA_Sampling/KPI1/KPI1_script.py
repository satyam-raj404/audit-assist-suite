# KPI1/kpi1.py
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")


def run_kpi1(input_folder: str, output_folder: str, frac: float = 0.3) -> str:
    """
    Runs KPI 1 sampling logic.
    Accepts either:
      - input_folder: a folder that contains "Active GL List.xlsx" and "Old samples selected.xlsx"
      - OR input_folder: the full path to "Active GL List.xlsx" directly
    Returns the path to the written output Excel file.
    """
    input_path = Path(input_folder)
    output_path = Path(output_folder) / "IOA Sampling"
    output_path.mkdir(parents=True, exist_ok=True)

    # Decide whether input_path is a file or a directory
    if input_path.is_file():
        gl_file = input_path
        base_input_dir = input_path.parent
    else:
        base_input_dir = input_path
        gl_file = base_input_dir / "Active GL List.xlsx"

    old_sample_file = base_input_dir / "Old samples selected.xlsx"

    if not gl_file.exists():
        raise FileNotFoundError(f"{gl_file} not found")
    if not old_sample_file.exists():
        raise FileNotFoundError(f"{old_sample_file} not found")

    # read using the resolved paths
    gl_df = pd.read_excel(gl_file)
    old_sample_df = pd.read_excel(old_sample_file)

    manual_gls = gl_df[gl_df.get("system_only_flag") == "N"]

    def keyword_filter(df: pd.DataFrame, keywords):
        if df is None or df.empty:
            return df
        pattern = "|".join(keywords)
        return df[df["IOA_Name"].astype(str).str.upper().str.contains(pattern, case=False, na=False)]

    cash_keywords = ["CASH", "CSH"]
    exchange_keywords = ["Exchange", "EXC"]
    clearing_keywords = ["Clear", "CLG"]

    cash_df = keyword_filter(manual_gls, cash_keywords)
    exchange_df = keyword_filter(manual_gls, exchange_keywords)
    clearing_df = keyword_filter(manual_gls, clearing_keywords)

    old_sample_copy = gl_df.iloc[0:0].copy()
    old_sample_copy = pd.DataFrame(old_sample_df.values, columns=old_sample_copy.columns)

    def stratified_sample(df: pd.DataFrame, group_col: str, frac=0.3):
        if df is None or df.empty:
            return df
        return df.groupby(group_col, group_keys=False).apply(
            lambda x: x.sample(frac=frac, random_state=42)
        )

    cash_sample = stratified_sample(cash_df, "Owner_of_the_GL", frac=frac)
    # exchange_sample = stratified_sample(exchange_df, "Owner_of_the_GL", frac=frac)
    # clearing_sample = stratified_sample(clearing_df, "Owner_of_the_GL", frac=frac)

    def exclude_previous(current_df: pd.DataFrame, previous_df: pd.DataFrame):
        if current_df is None or current_df.empty:
            return current_df
        if previous_df is None or previous_df.empty:
            return current_df.reset_index(drop=True)

        # Use merge-based anti-join on all columns to be safe:
        merged = current_df.merge(previous_df, indicator=True, how="left")
        new_only = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
        return new_only.reset_index(drop=True)

    cash_final = exclude_previous(cash_sample, old_sample_copy)
    # exchange_final = exclude_previous(exchange_sample, old_sample_copy)
    # clearing_final = exclude_previous(clearing_sample, old_sample_copy)

    out_file = output_path / "KPI1_GL_Samples_Updated.xlsx"
    with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
        if cash_final is not None and not cash_final.empty:
            cash_final.to_excel(writer, sheet_name="Cash", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Cash", index=False)

        # Uncomment when exchange/clearing implemented
        # if exchange_final is not None and not exchange_final.empty:
        #     exchange_final.to_excel(writer, sheet_name="Exchange", index=False)
        # if clearing_final is not None and not clearing_final.empty:
        #     clearing_final.to_excel(writer, sheet_name="Clearing", index=False)

    return str(out_file)


def main(input_folder: str = "Input file", output_folder: str = "Output files", frac: float = 0.3) -> str:
    """
    CLI-friendly entrypoint. Returns output file path.
    """
    return run_kpi1(input_folder, output_folder, frac)


if __name__ == "__main__":
    # Example local test — pass either the input folder or the exact file path
    base = Path("IOA_Sampling") / "KPI1"
    input_example = base / "Input file" / "Active GL List.xlsx"   # or: base / "Input file"
    output_example = base / "Output files"
    print(main(str(input_example), str(output_example)))
