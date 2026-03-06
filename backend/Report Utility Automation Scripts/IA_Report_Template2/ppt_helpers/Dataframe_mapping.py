import pandas as pd
import numpy as np

def make_ordered_hashmap(df_obs):
    """
    Return dict: {header: [value1, value2, ...]} preserving row order.
    NaN values are skipped; duplicates are kept.
    """
    hashmap = {}
    for col in df_obs.columns:
        vals = df_obs[col].dropna().astype(str).tolist()
        hashmap[col] = vals
    return hashmap

# Map all columns in df_obs to their headers 
def map_columns(df_obs): 
    """
    Return a dict mapping column index -> column header for df_obs. 
    """ 
    return {idx: col for idx, col in enumerate(df_obs.columns)}


def _find_table_for_mapping(slide, header_cell_map):
    """
    Return the table shape on `slide` that can accommodate all (r,c) coords
    in header_cell_map. If multiple tables qualify, return the one with the
    smallest area (rows*cols). Return None if no table fits.
    - slide: pptx.slide.Slide
    - header_cell_map: dict mapping header -> (row_index, col_index)
    """
    candidates = []
    # iterate shapes and collect table shapes that can fit all mapped coords
    for sh in slide.shapes:
        if not getattr(sh, "has_table", False):
            continue
        tbl = sh.table
        rows = len(tbl.rows)
        cols = len(tbl.columns)
        fits_all = True
        for (r, c) in header_cell_map.values():
            # ensure r,c are ints and within bounds
            if not (isinstance(r, int) and isinstance(c, int) and 0 <= r < rows and 0 <= c < cols):
                fits_all = False
                break
        if fits_all:
            candidates.append((rows * cols, sh))  # area, shape

    if not candidates:
        return None
    # choose the candidate with smallest area (best fit)
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]

def remove_whitespace_cells(df: pd.DataFrame, drop_empty_rows=True, drop_empty_cols=True):
    """
    Replace cells that contain only whitespace with empty string, then
    optionally drop rows/columns that are entirely empty.
    """
    if df is None:
        return pd.DataFrame()

    # make a copy to avoid mutating original
    out = df.copy(deep=True)

    # strip strings and convert whitespace-only strings to empty string
    def _strip_or_keep(x):
        if isinstance(x, str):
            s = x.strip()
            return s if s != "" else ""
        return x

    out = out.applymap(_strip_or_keep)

    # normalize empty-like values to NaN for easy dropping
    out = out.replace("", np.nan)

    if drop_empty_rows:
        out = out.dropna(axis=0, how="all").reset_index(drop=True)
    if drop_empty_cols:
        out = out.dropna(axis=1, how="all")

    # restore empty strings instead of NaN if you prefer
    out = out.fillna("")

    return out
