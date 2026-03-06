import pandas as pd
import logging
from pathlib import Path

def setup_logging():
    """Configure logging format and level."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )


def locate_header_row(df_raw: pd.DataFrame, required_cols: set) -> int:
    """
    Scan the first 20 rows of df_raw to find the header row
    which contains all required_cols. Returns its index.
    """
    for idx, row in df_raw.head(20).iterrows():
        headers = set(str(x).strip() for x in row.values if pd.notna(x))
        if required_cols.issubset(headers):
            logging.info(f"Detected header row at index {idx}")
            return idx
    raise ValueError(f"Failed to locate header row with columns: {required_cols}")

def read_swift_sheet(file_a: Path, sheet_name: str, key_cols: set) -> pd.DataFrame:
    """
    Read the 'Swift Dumps' sheet from File A, locate the true header row,
    and return a DataFrame with only the key columns retained.
    """
    # 1. Load without header to detect correct header row
    df_raw = pd.read_excel(file_a, sheet_name=sheet_name, header=None)
    header_idx = locate_header_row(df_raw, key_cols)

    # 2. Reload with header
    df = pd.read_excel(file_a, sheet_name=sheet_name, header=header_idx, dtype=str)

    # 3. Keep only rows where at least one key column is non-null
    cols = list(key_cols)
    mask = df[cols].notna().any(axis=1)
    df_clean = df.loc[mask, cols].drop_duplicates().reset_index(drop=True)
    logging.info(f"Loaded {len(df_clean)} reference rows from '{sheet_name}'")
    return df_clean


def match_all_sheets(df_ref: pd.DataFrame, file_b: Path, key_cols: set) -> dict:
    """
    For every sheet in File B:
      1) detect the real header row via locate_header_row()
      2) re-load with that header
      3) strip blank rows & trim column names
      4) match on any intersection of key_cols
      5) collect reconciled rows with status flag
    """
    matched_results = {}
    xls = pd.ExcelFile(file_b)

    for sheet_name in xls.sheet_names:
        # 1. Load raw (no header) for header detection
        df_raw = pd.read_excel(file_b, sheet_name=sheet_name, header=None, dtype=str)
        try:
            header_idx = locate_header_row(df_raw, key_cols)
        except ValueError:
            logging.warning(f"Skipped '{sheet_name}': no header with {key_cols}")
            continue

        # 2. Reload with correct header row
        df = pd.read_excel(file_b, sheet_name=sheet_name,
                           header=header_idx, dtype=str)

        # 3. Clean up
        df.columns = df.columns.str.strip()            # trim column names
        df = df.dropna(how="all").reset_index(drop=True)  # drop fully blank rows

        # 4. Find which keys exist in this sheet
        common = [c for c in key_cols if c in df.columns]
        if not common:
            logging.warning(f"Skipped '{sheet_name}': no key columns present")
            continue

        # 5. Build mask: any-key match
        mask = pd.Series(False, index=df.index)
        for col in common:
            ref_vals = df_ref[col].dropna().astype(str).unique()
            mask |= df[col].astype(str).isin(ref_vals)

        # 6. Collect reconciled rows
        df_matched = df.loc[mask].copy()
        if df_matched.empty:
            logging.info(f"'{sheet_name}': no reconciled rows")
            continue

        df_matched["Reconciliation Status"] = "Reconciled"
        matched_results[sheet_name] = df_matched.reset_index(drop=True)
        logging.info(f"'{sheet_name}': {len(df_matched)} rows reconciled")

    return matched_results

def save_results(matched: dict, output_path: Path):
    """
    Save each reconciled DataFrame to its own sheet in a new Excel file.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in matched.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    logging.info(f"Reconciliation file written to: {output_path}")

def recon_main(tracker_path , output_folder_path):
    setup_logging()
    tracker_path = Path(tracker_path)
    input_path = Path(tracker_path)
    output_path = Path(output_folder_path)
    if input_path.is_file():
        input_folder = input_path.parent
    else:
        input_folder = input_path

    # if output_folder:
    #     output_folder = Path(output_folder)
    # else:
    #     output_folder = input_folder / "Output file"
    # output_folder.mkdir(parents=True, exist_ok=True)

    file_a = input_folder / "A.xlsx"
    file_b = input_folder / "B.xlsx"

    # # File paths
    # file_a = tracker_path
    # file_b = recon_path
    result_file = output_path / "C.xlsx"

    # Key columns for matching
    key_columns = {"Reference", "MUR", "Transaction Reference"}

    # Step 1: Read reference data from File A
    df_reference = read_swift_sheet(file_a, sheet_name="Swift Dumps", key_cols=key_columns)

    # Step 2: Match against every sheet in File B
    matched_sheets = match_all_sheets(df_reference, file_b, key_columns)

    # Step 3: Save reconciled rows
    if matched_sheets:
        save_results(matched_sheets, result_file)
    else:
        logging.warning("No reconciled data found; no output file generated.")


if __name__ == "__main__":
    recon_main(r"C:\Users\satyambarnwal\Downloads\FS Utility Input\FS Utility Input\Swift\Input\A.xlsx" , r"C:\Users\satyambarnwal\Downloads\FS Utility Input\FS Utility Input\Swift\Output") 
