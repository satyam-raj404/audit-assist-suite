### Project Technical Documentation

This document describes the purpose and usage of the main Python modules you provided. Each file section lists the functions it exposes, their **purpose**, **parameters**, **return values**, **behavior**, and **implementation notes** to help you integrate, test, and maintain the code.

---

### FS_Utility.py

**Context**  
GUI launcher that collects user inputs (Excel, PPTX, output folder, audit/utility/report types, month/year) and dispatches to the appropriate processing function for each audit utility.

#### Key functions

##### `browse_excel_csv()`
- **Purpose**: Open file dialog to select an Excel or CSV file and populate the GUI entry.
- **Parameters**: None (uses GUI widgets `excel_entry`, `pptx_entry`, etc.).
- **Returns**: None.
- **Behavior**: Enables PPTX selection controls after a valid Excel/CSV is chosen.

##### `browse_pptx()`
- **Purpose**: Open file dialog to select a PowerPoint template and populate the GUI entry.
- **Parameters**: None.
- **Returns**: None.
- **Behavior**: Enables output folder controls after a valid PPTX is chosen.

##### `browse_output_folder()`
- **Purpose**: Open directory chooser and populate the output folder entry.
- **Parameters**: None.
- **Returns**: None.

##### `submit_paths()`
- **Purpose**: Validate GUI inputs and call the appropriate main function for the selected audit utility.
- **Parameters**: None (reads GUI state: `excel_entry`, `pptx_entry`, `output_entry`, `audit_var`, `utility_var`, `report_var`, `month_var`, `year_var`).
- **Returns**: None.
- **Behavior**:
  - Handles **FOBO**: validates Excel and output folder, calls `fo_bo_main`, shows result message, closes GUI.
  - Handles **ICOFR**: validates Excel, PPTX, output folder, calls `ICOFR_main`, shows result message, closes GUI.
  - Handles other audit types: validates required inputs depending on `audit_type` and `utility_type`, calls `main(...)` (the generic entry point), shows result message, closes GUI.
  - Uses `messagebox` for warnings/errors and `logging.exception` for exceptions.
- **Notes**:
  - `main(...)` is expected to return an output filename (string) or raise on error.
  - `submit_paths` constructs `final_output_path` using `output_folder` and the returned filename.

##### `add_placeholder(entry, placeholder)`
- **Purpose**: Add placeholder text behavior to an `Entry` widget (grey placeholder that clears on focus).
- **Parameters**: `entry` (tk.Entry), `placeholder` (str).
- **Returns**: None.

##### `toggle_report_type(*args)`
- **Purpose**: Show/hide GUI fields depending on selected `audit_type`.
- **Parameters**: `*args` (used as callback signature).
- **Returns**: None.
- **Behavior**: Enables/disables and grids GUI frames and menu options for different audit types (FOBO, ICOFR, FD - Sampling, KYC, CCIL, IOA KPI1-4, Swift, Concurrent Audit, Internal Audit variants).

**Implementation notes**
- GUI state is stored in global tkinter variables and widgets referenced by these functions.
- Add `output_folder_path = Path(existing_output_folder) / "IA_Zensar"` where needed to include the `IA_Zensar` subfolder (create with `mkdir(parents=True, exist_ok=True)`).

---

### IA_T1 (IA_Report_Template2 helpers)

**Context**  
Utility functions for PowerPoint manipulation: duplicating slides, deleting slides, replacing titles, updating tables, formatting cells, and handling page numbers and overflow.

#### Key functions

##### `duplicate_slide(ppt, slide)`
- **Purpose**: Create a deep copy of `slide` in `ppt` preserving layout and shapes.
- **Parameters**:
  - `ppt`: `Presentation` object (python-pptx).
  - `slide`: `Slide` object to duplicate.
- **Returns**: `Slide` object (the newly created slide).
- **Behavior**:
  - Adds a new slide with the same layout.
  - Removes auto-created placeholders from the new slide.
  - Copies each shape by parsing the source shape XML and inserting into the new slide.
- **Notes**: Uses `parse_xml(shape.element.xml)` to duplicate shape XML.

##### `delete_slide(ppt, index)`
- **Purpose**: Remove a slide safely by index without triggering PowerPoint repair mode.
- **Parameters**:
  - `ppt`: `Presentation` object.
  - `index`: zero-based slide index to delete.
- **Returns**: `ppt` (modified Presentation).
- **Behavior**:
  - Removes the slide ID element from the presentation XML and drops the relationship.

##### `replace_and_format_title(shape, replacement_text)`
- **Purpose**: Replace a shape's title text and apply KPMG styling.
- **Parameters**:
  - `shape`: shape object with `has_text_frame`.
  - `replacement_text`: string to set.
- **Returns**: None.
- **Behavior**:
  - Clears text frame, inserts `replacement_text`, sets font to *Univers for KPMG*, size `18 pt`, color `#002060`.

##### `update_table_texts(shapes_with_tables, replacing_text, review_text)`
- **Purpose**: Replace table cell text matching `replacing_text` with `review_text` and apply formatting.
- **Parameters**:
  - `shapes_with_tables`: iterable of shapes.
  - `replacing_text`: string to match.
  - `review_text`: replacement string (empty or value).
- **Returns**: None.
- **Behavior**:
  - For each table cell equal to `replacing_text`, sets new text (or blank) and applies font `Univers for KPMG`, `11 pt`, color `#002060`.

##### `update_risk_rating(shapes_with_textbox, RR)`
- **Purpose**: Update small textbox shapes used for risk rating (H/M/L) with background color and formatted text.
- **Parameters**:
  - `shapes_with_textbox`: iterable of shapes.
  - `RR`: one of `'H'`, `'M'`, `'L'`.
- **Returns**: None.
- **Behavior**:
  - Sets shape fill color based on `RR`, clears text frame, centers and writes `RR` with `Univers for KPMG`, `16 pt`, bold, white.

##### `update_risk_table(shapes_with_tables, replacing_text, review_text)`
- **Purpose**: For table cells matching `replacing_text`, set cell fill color when `review_text == "Yes"`.
- **Parameters**: same pattern as `update_table_texts`.
- **Returns**: None.

##### `delete_extra_rows(table, keep_row_count)`
- **Purpose**: Remove rows beyond `keep_row_count` from a PPTX table.
- **Parameters**:
  - `table`: `pptx.table.Table` object.
  - `keep_row_count`: integer number of rows to keep.
- **Returns**: None.
- **Behavior**: Manipulates the underlying `<a:tbl>` XML to remove `<a:tr>` elements.

##### `update_page_no(ppt, slide_idx, reference_text, replace_text)`
- **Purpose**: Find slide index containing `reference_text` and write that page number into a table cell on `slide_idx` where cell text equals `replace_text`.
- **Parameters**:
  - `ppt`: `Presentation`.
  - `slide_idx`: target slide index to update.
  - `reference_text`: text to search across slides to determine page number.
  - `replace_text`: placeholder text in table cell to replace with page number.
- **Returns**: None.
- **Behavior**:
  - Searches slides for `reference_text`, computes 1-based page number, writes into matching table cell and formats the cell text.

##### `check_overflow(cell, max_lines)`
- **Purpose**: Detect if a cell's text exceeds `max_lines` and split into first/remaining lines.
- **Parameters**:
  - `cell`: table cell object with `.text`.
  - `max_lines`: integer.
- **Returns**: tuple `(overflow_flag: bool, first_lines: str, remaining_lines: str)`.

##### `format_cell(cell, font_name, font_size, font_color)`
- **Purpose**: Apply consistent font formatting to all runs in a table cell.
- **Parameters**:
  - `cell`: table cell.
  - `font_name`: default `"Univers for KPMG"`.
  - `font_size`: default `11`.
  - `font_color`: default `RGBColor(0x00,0x20,0x60)`.
- **Returns**: None.

**Implementation notes**
- Functions rely on `python-pptx` internals and direct XML manipulation for robust duplication and deletion.
- Keep `__init__.py` present in `IA_Report_Template2` and `ppt_helpers` to maintain package imports.

---

### Swift

**Context**  
Utilities to read and reconcile Swift dump sheets across Excel files. Designed to detect header rows dynamically and match rows across multiple sheets.

#### Key functions

##### `locate_header_row(df_raw, required_cols)`
- **Purpose**: Detect the header row index within the first 20 rows of a raw DataFrame that contains all `required_cols`.
- **Parameters**:
  - `df_raw`: `pd.DataFrame` loaded with `header=None`.
  - `required_cols`: `set` of required column names (strings).
- **Returns**: integer header row index.
- **Raises**: `ValueError` if header row not found.
- **Behavior**: Scans first 20 rows, trims values, logs detection.

##### `read_swift_sheet(file_a, sheet_name, key_cols)`
- **Purpose**: Read a Swift reference sheet, detect header row, and return a cleaned DataFrame containing only `key_cols`.
- **Parameters**:
  - `file_a`: `Path` to Excel file.
  - `sheet_name`: sheet name string.
  - `key_cols`: set/list of key column names.
- **Returns**: `pd.DataFrame` with unique rows and only key columns.
- **Behavior**:
  - Loads sheet twice (no header to detect header row, then with header).
  - Filters rows where at least one key column is non-null, drops duplicates.

##### `match_all_sheets(df_ref, file_b, key_cols)`
- **Purpose**: For each sheet in `file_b`, detect header, reload, clean, and match rows against `df_ref` on any key column intersection.
- **Parameters**:
  - `df_ref`: reference DataFrame (from `read_swift_sheet`).
  - `file_b`: `Path` to Excel file to reconcile.
  - `key_cols`: set of key column names.
- **Returns**: `dict` mapping sheet name → reconciled `pd.DataFrame` (with `"Reconciliation Status"` column).
- **Behavior**:
  - Detects header row per sheet.
  - Trims column names, drops blank rows.
  - Builds mask by checking membership of values in `df_ref` for each common key column.
  - Adds `"Reconciliation Status" = "Reconciled"` for matched rows.

##### `save_results(matched, output_path)`
- **Purpose**: Write matched reconciliation results to an Excel workbook with one sheet per matched sheet.
- **Parameters**:
  - `matched`: dict of sheet_name → DataFrame.
  - `output_path`: `Path` to write Excel file.
- **Returns**: None.
- **Behavior**: Uses `pd.ExcelWriter` with `openpyxl` engine.

**Implementation notes**
- All reads use `dtype=str` where appropriate to avoid numeric coercion.
- Logging is used to report skipped sheets and counts.

---

### KYC

**Context**  
KYC-specific KPI computations reading multiple sheets and producing a consolidated Excel result.

#### Key function

##### `KYC_main(excel_file, output_folder_path)`
- **Purpose**: Compute four KPIs from the provided Excel workbook and save results to `KYC_Result.xlsx` in `output_folder_path`.
- **Parameters**:
  - `excel_file`: path to input Excel workbook.
  - `output_folder_path`: folder path where `KYC_Result.xlsx` will be written.
- **Returns**: `full_path` (string) to the written Excel file.
- **Behavior**:
  - Loads `Sheet2`, `Sheet3`, `Sheet4`.
  - **KPI 1**: rows where same `LCIN` have multiple `PAN No` values.
  - **KPI 2**: rows where same `LCIN` have multiple `Final Risk Rating` values.
  - **KPI 3**: maps PAN Type using 4th character of `PAN No` via mapping from `Sheet3`.
  - **KPI 4**: flags `DBSIC Description` as `"High Risk"` if any word matches descriptions in `Sheet4`.
  - Writes `KPI_1`, `KPI_2`, `KPI_3`, `KPI_4` sheets into `KYC_Result.xlsx`.
- **Notes**:
  - `mapping_dict` expects `PAN 4 character list` entries split by `–` (en dash). Validate input format.
  - Uses `openpyxl` engine for writing.

---

### FOBO

**Context**  
FO vs BO vs GL reconciliation logic that compares three sheets and writes a consolidated reconciliation workbook.

#### Key functions

##### `reconcile_all(input_path, output_path=None)`
- **Purpose**: Perform FO-BO, FO-GL, and BO-GL reconciliations and save results to an Excel workbook.
- **Parameters**:
  - `input_path`: path to Excel file containing sheets `"FO Dump"`, `"BO Dump"`, `"GL Dump"`.
  - `output_path`: optional path to write output Excel. If `None`, writes `FOBO_Reconciliation_Output.xlsx` next to input file.
- **Returns**: path to the written output file (string).
- **Behavior**:
  - Loads three sheets using `openpyxl`.
  - Sets `"Link Ref."` as index for FO, BO, GL.
  - **FO vs BO**:
    - Reindexes both to union of link refs, prefixes columns with `FO_` and `BO_`, compares common columns.
    - Produces `"Reconciliation Status"` and `"Remarks"` per row (Missing Value or list of mismatched columns).
  - **FO vs GL**:
    - Compares `CCY1` vs `Ccy` and `Amount1` vs `Atlas Value`.
  - **BO vs GL**:
    - Compares `BO_System ID` vs `GL_Ref.`, `BO_CCY1` vs `GL_Ccy`, `BO_Amount1` vs `GL_Atlas Value`.
  - Writes three sheets: `"FO vs BO"`, `"FO vs GL"`, `"BO vs GL"`.
- **Notes**:
  - The function expects specific column names (`Link Ref.`, `CCY1`, `Amount1`, `Ccy`, `Atlas Value`, `System ID`, `Ref.`). Validate input schema before running.
  - Output filename default is `"FOBO_Reconciliation_Output.xlsx"`.

---

## General Integration Notes

- **Package imports**: Ensure `IA_Report_Template2` and its subpackages (e.g., `ppt_helpers`) contain `__init__.py` and that your project root is on `sys.path` when running `FS_Utility.py`. Prefer running the launcher as a module from the project root:
  ```bash
  python -m Utility_Code.FS_Utility
  ```
  or add the project root to `sys.path` at the top of `FS_Utility.py`:
  ```py
  from pathlib import Path
  import sys
  PROJECT_ROOT = Path(__file__).resolve().parents[1]
  sys.path.insert(0, str(PROJECT_ROOT))
  ```

- **Paths and output folder**: To include the `IA_Zensar` subfolder in outputs, construct `output_folder_path` using `pathlib`:
  ```py
  from pathlib import Path
  output_folder_path = Path(existing_output_folder) / "IA_Zensar"
  output_folder_path.mkdir(parents=True, exist_ok=True)
  ```

- **PowerPoint geometry**: If you manipulate shapes and sizes, convert EMU ↔ pt consistently (12700 EMU per point) when mixing `python-pptx` geometry and manual point values.

- **Robustness**:
  - Validate input Excel schemas (expected sheet names and column names) before processing.
  - Wrap heavy processing calls in `try/except` and log exceptions with `logging.exception` as done in `FS_Utility.py`.
  - Use `dtype=str` when reading Excel sheets where numeric IDs or leading zeros matter.

---

I can now generate a single-file patch to (a) replace the relative imports in `inspect_placeholders.py` with absolute imports, (b) add the `sys.path` insertion to `FS_Utility.py`, or (c) add the `IA_Zensar` subfolder creation to the place where you compute `output_folder_path`. Tell me which change you want and I’ll produce the exact code diff.