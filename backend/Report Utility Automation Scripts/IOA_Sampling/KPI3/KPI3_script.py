# kpi3.py
import os
from pathlib import Path
from typing import Optional
import pandas as pd


def run_kpi3(input_folder: str, output_folder: Optional[str] = None) -> str:
    """
    Run KPI3 filtering and write output Excel.
    - input_folder: folder containing 'Agri disbursement dump.xlsx' and 'Unit wise list of accounts.xlsx'
    - output_folder: destination folder; if None, defaults to 'Output files' inside cwd
    Returns the path to the written output file.
    """
    input_folder = Path(input_folder)
    input_path = Path(input_folder)
    if input_path.is_file():
        input_folder = input_path.parent
    else:
        input_folder = input_path

    if output_folder:
        output_path = Path(output_folder) / "IOA Sampling"
    else:
        output_path = Path("Output files")
    output_path.mkdir(parents=True, exist_ok=True)

    agri_file = input_folder / "Agri disbursement dump.xlsx"
    gl_file = input_folder / "Unit wise list of accounts.xlsx"

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
        merged_df['TRAN_ID'].astype(str).str.upper().str.strip().str.startswith('M') &
        (merged_df['PART_TRAN_TYPE'].astype(str).str.upper().str.strip() == 'C')
    ]

    final_df = pd.merge(
        filtered_df[['TRAN_ID']],
        agri_df,
        on='TRAN_ID',
        how='inner'
    ).drop_duplicates()

    out_file = output_path / "KPI3_disbursement_output.xlsx"
    final_df.to_excel(out_file, index=False, engine='xlsxwriter')

    return str(out_file)


def main(input_folder: str = "Input files", output_folder: Optional[str] = None) -> str:
    """
    CLI-friendly entrypoint for KPI3.
    """
    result = run_kpi3(input_folder, output_folder)
    print(f"Filtered data saved to {result}")
    return result


if __name__ == "__main__":
    main()

