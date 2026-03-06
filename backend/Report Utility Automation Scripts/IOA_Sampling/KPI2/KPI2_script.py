import os
from pathlib import Path
from typing import Optional
import pandas as pd


def run_kpi2_simple(input_folder: str, output_folder: Optional[str] = None) -> str:
    """
    Core KPI2 logic (keeps your original simple implementation).
    - input_folder: folder containing 'Agri disbursement dump.xlsx' and 'Unit wise list of accounts.xlsx'
    - output_folder: destination folder for the output file; if None, uses 'Output file' inside input_folder
    Returns the path to the written output Excel file (str).
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
        output_path = input_folder / "Output file"
    output_path.mkdir(parents=True, exist_ok=True)

    agri_file = input_folder / "Agri disbursement dump.xlsx"
    gl_file = input_folder / "Unit wise list of accounts.xlsx"

    if not agri_file.exists():
        raise FileNotFoundError(f"Missing input file: {agri_file}")
    if not gl_file.exists():
        raise FileNotFoundError(f"Missing input file: {gl_file}")

    agri_df = pd.read_excel(agri_file)
    gl_df = pd.read_excel(gl_file)

    agri_df["ACC_NO"] = agri_df["ACC_NO"].astype(str).str.strip()
    gl_df["IOA COMBINATION AGGREGATED"] = gl_df["IOA COMBINATION AGGREGATED"].astype(str).str.strip()

    joined_df = pd.merge(
        agri_df,
        gl_df,
        left_on="ACC_NO",
        right_on="IOA COMBINATION AGGREGATED",
        how="inner"
    )

    keywords = [
        'REV', 'REVERSAL', 'NEFT.*RETURNED', 'NEFT.*RETURN', 'NEFT.*RTN', 'REFUND',
        'ACCOUNT.*DOES.*NOT.*EXIST', 'UNKNOWNENDCUSTOMER', 'FUND.*TRF',
        'FINACLE.*SETTELMENT.*ACC', 'NOT.*SPECIFIED.*REASON.*CUSTOMER.*GENERA',
        'INCORRECT.*ACCOUNT.*NUMBER', 'PAYMENT.*STOPPED', 'PAYMENT.*STOP', 'FUND/TRF'
    ]

    pattern = '|'.join([kw.replace(' ', r'\s*').replace('/', r'[/]?') for kw in keywords])

    # Filter rows
    filtered_df = joined_df[
        joined_df['TRAN_PARTICULAR']
        .astype(str)
        .str.upper()
        .str.strip()
        .str.contains(pattern, regex=True, na=False)
    ]

    final_df = pd.merge(
        filtered_df[['TRAN_ID']],
        agri_df,
        on='TRAN_ID',
        how='inner'
    ).drop_duplicates()

    output_file = output_path / 'KPI2_disbursement_output.xlsx'
    final_df.to_excel(output_file, index=False, engine='xlsxwriter')

    return str(output_file)


def main(input_folder: str = "Input file", output_folder: Optional[str] = None) -> str:
    """
    CLI-friendly entrypoint. Calls run_kpi2_simple and prints the output path.
    Returns the output path string.
    """
    result_path = run_kpi2_simple(input_folder, output_folder)
    print(f"Filtered data saved to {result_path}")
    return result_path


if __name__ == "__main__":
    # Simple local test using the same folder names you used previously
    in_folder = r"IOA_Sampling\KPI2\Input file"              # folder that contains the two input xlsx files
    out_folder = r"IOA_Sampling\KPI2\Output file"            # destination folder for output

    # If you run from project root and the folders are under KPI2, update paths accordingly:
    # in_folder = r"IOA_Sampling\KPI2\Input file"
    # out_folder = r"IOA_Sampling\KPI2\Output file"

    main(in_folder, out_folder)
