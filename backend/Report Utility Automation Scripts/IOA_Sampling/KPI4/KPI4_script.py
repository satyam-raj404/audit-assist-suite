# kpi4.py
import os
from pathlib import Path
from typing import Optional
import pandas as pd


def run_kpi4(input_folder: str, output_folder: Optional[str] = None) -> str:
    """
    Run KPI4 collection filtering.
    - input_folder: folder containing 'agri collection dump.xlsx' and 'Unit wise list of account collection.xlsx'
    - output_folder: destination folder; if None, defaults to 'Output files' in cwd
    Returns the path to the written output Excel file.
    """
    input_folder = Path(input_folder)
    input_path = Path(input_folder)
    if input_path.is_file():
        input_folder = input_path.parent
    else:
        input_folder = input_path

    if output_folder:
        output_folder = Path(output_folder) / "IOA Sampling"
    else:
        output_folder = Path("Output files")
    output_folder.mkdir(parents=True, exist_ok=True)

    agri_file = input_folder / "agri collection dump.xlsx"
    gl_file = input_folder / "Unit wise list of account collection.xlsx"

    if not agri_file.exists():
        raise FileNotFoundError(f"Missing input file: {agri_file}")
    if not gl_file.exists():
        raise FileNotFoundError(f"Missing input file: {gl_file}")

    agri_df = pd.read_excel(agri_file)
    gl_df = pd.read_excel(gl_file)

    agri_df['ACC_NO'] = agri_df['ACC_NO'].astype(str).str.strip()
    gl_df['IOA COMBINATION AGGREGATED'] = gl_df['IOA COMBINATION AGGREGATED'].astype(str).str.strip()

    merged_df = pd.merge(
        agri_df,
        gl_df,
        left_on='ACC_NO',
        right_on='IOA COMBINATION AGGREGATED',
        how='inner'
    )

    filtered_df = merged_df[
        merged_df['TRAN_PARTICULAR']
        .astype(str)
        .str.upper()
        .str.strip()
        .str.contains('REV|REVERSAL', regex=True, na=False)
    ]

    final_df = pd.merge(
        filtered_df[['TRAN_ID']],
        agri_df,
        on='TRAN_ID',
        how='inner'
    ).drop_duplicates()

    out_file = output_folder / "KPI4_collection_output.xlsx"
    final_df.to_excel(out_file, index=False, engine='xlsxwriter')

    return str(out_file)


def main(input_folder: str = "Input files", output_folder: Optional[str] = None) -> str:
    """
    CLI-friendly entrypoint for KPI4.
    """
    result = run_kpi4(input_folder, output_folder)
    print(f"Filtered data saved to {result}")
    return result


if __name__ == "__main__":
    main()

