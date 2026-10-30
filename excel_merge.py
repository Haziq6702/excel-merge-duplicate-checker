#!/usr/bin/env python3

"""

excel_merge file.py

========



Terminal tool with two modes:



  Mode 1 - Merge multiple Excel FILES into one worksheet.

           Files are picked with Windows file/folder dialogs and can come

           from ANY number of different folders.

  Mode 2 - Merge multiple SHEETS from a single Excel file into one worksheet.



Both modes share the same core engine: header-name column alignment,

key-column selection/validation, duplicate checking (within-source and

cross-source), a detailed "Duplicate Report" sheet (full row data + where

each row came from), and a timestamped output file.



Run with:

    python "excel_merge file.py"



Requires:

    pip install pandas openpyxl numpy xlsxwriter

    (optional, much faster reading: pip install python-calamine)

    (tkinter ships with the standard Windows Python installer)

"""



import os

import re

import sys

import time

import traceback

import warnings

from datetime import datetime

from dataclasses import dataclass



import numpy as np

import pandas as pd

from openpyxl import load_workbook

from openpyxl.styles import PatternFill



# Source workbooks sometimes contain cells formatted as dates but holding

# an out-of-range serial value (e.g. a blank/zero cell with a date format

# applied). openpyxl warns loudly about every such cell; the cell is

# harmless (it just comes through as blank/NaN), so the warning is

# suppressed here to keep the terminal output readable. Fix the source

# formatting in Excel later if you want those cells to show real dates.

warnings.filterwarnings(

    "ignore",

    message=r"Cell .* is marked as a date but the serial value .* is outside the limits for dates.*",

    category=UserWarning,

)



# ------------------------------------------------------------------------

# Configuration

# ------------------------------------------------------------------------



OUTPUT_DIR = ""  # Please enter your preferred output folder path here. Leave blank to use the current working directory.



EXCEL_EXTENSIONS = (".xlsx", ".xlsm")



# Excel allows 1,048,576 rows per worksheet; one is the header row.

# Anything bigger is automatically continued on a new sheet.

EXCEL_MAX_DATA_ROWS = 1_048_575



# Remembers the last folder used in a dialog so the next dialog opens there.

_last_dialog_dir = ""



# Fixed columns at the front of the Duplicate Report.

REPORT_FIXED_COLUMNS = [

    "Key Value",

    "Source File",

    "Source Folder",

    "Source Sheet",

    "Source Row",

    "Occurrences",

    "Duplicate Type",

    "Action",

]



# Cell colours used to mark duplicate rows (by Duplicate Type).

DUP_COLORS = {

    "Cross-source": "F8CBAD",                 # light orange

    "Within-source": "FFF2CC",                # light yellow

    "Cross-source + Within-source": "F4B084", # stronger orange

}





# ------------------------------------------------------------------------

# Data model

# ------------------------------------------------------------------------



@dataclass

class SourceEntry:

    """

    One unit of data to be merged.

    In Mode 1, a source = one worksheet of one Excel FILE.

        label = "\<full path> | \<sheet>"  (unique per file+sheet)

    In Mode 2, a source = one WORKSHEET (label = sheet name).



    path / file_name / folder / sheet_name are used for the detailed

    Duplicate Report ("where did this row come from?").

    """

    label: str

    df: pd.DataFrame

    clean_df: pd.DataFrame = None

    rows_found: int = 0

    rows_added: int = 0

    path: str = ""

    file_name: str = ""

    folder: str = ""

    sheet_name: str = ""





# ------------------------------------------------------------------------

# Small display / input helpers

# ------------------------------------------------------------------------



LINE = "=" * 60

DASH = "-" * 60





def print_banner(title):

    print()

    print(LINE)

    print(title)

    print(LINE)





def print_section(title):

    print()

    print(DASH)

    print(title)

    print(DASH)





def info(msg):

    print(f"[INFO] {msg}")





def ok(msg):

    print(f"[OK] {msg}")





def warning(msg):

    print(f"[WARNING] {msg}")





def error(msg):

    print(f"[ERROR] {msg}")





def ask(prompt):

    try:

        return input(prompt).strip()

    except (EOFError, KeyboardInterrupt):

        print()

        info("Cancelled by user. Exiting.")

        sys.exit(0)





def ask_menu_choice(prompt, valid_choices):

    while True:

        choice = ask(prompt)

        if choice in valid_choices:

            return choice

        error(f"Invalid choice '{choice}'. Please try again.")





def ask_path(prompt):

    """

    Prompts for a file path and cleans it up: strips whitespace and any

    surrounding quotes (e.g. from Windows' "Copy as path", which wraps

    the path in double quotes).

    """

    raw = ask(prompt)

    cleaned = raw.strip()

    if len(cleaned) >= 2 and cleaned[0] == cleaned[-1] and cleaned[0] in ("'", '"'):

        cleaned = cleaned[1:-1].strip()

    return cleaned





# ------------------------------------------------------------------------

# Input files and folders are selected by the user at runtime.
# No personal or company-specific input path is stored in this template.
# Windows file / folder dialogs  (Mode 1)

# ------------------------------------------------------------------------



def _create_hidden_tk_root():

    """

    Creates an invisible Tk root window so only the dialog is shown

    (no blank Tkinter window). 'topmost' makes the dialog appear in front

    of the terminal instead of hiding behind it.

    """

    import tkinter as tk



    root = tk.Tk()

    root.withdraw()

    try:

        root.attributes("-topmost", True)

    except Exception:

        pass

    root.update()

    return root





def pick_excel_file():

    """

    Opens a Windows 'Open file' dialog limited to .xlsx / .xlsm.

    Returns the selected path, or None if the user cancelled.

    Falls back to typing the path if tkinter is unavailable.

    """

    global _last_dialog_dir

    try:

        from tkinter import filedialog

        root = _create_hidden_tk_root()

    except Exception as exc:

        warning(f"File dialog is not available ({exc}). Please type the path instead.")

        path = ask_path("Enter Excel file path (blank to cancel): ")

        return path or None



    try:

        path = filedialog.askopenfilename(

            parent=root,

            title="Select an Excel file",

            initialdir=_last_dialog_dir or None,

            filetypes=[("Excel files", "*.xlsx *.xlsm")],

        )

    finally:

        root.destroy()



    if not path:

        return None



    path = os.path.normpath(path)

    _last_dialog_dir = os.path.dirname(path)

    return path





def pick_folder():

    """

    Opens a Windows 'Select folder' dialog.

    Returns the selected folder path, or None if the user cancelled.

    Falls back to typing the path if tkinter is unavailable.

    """

    global _last_dialog_dir

    try:

        from tkinter import filedialog

        root = _create_hidden_tk_root()

    except Exception as exc:

        warning(f"Folder dialog is not available ({exc}). Please type the path instead.")

        path = ask_path("Enter folder path (blank to cancel): ")

        return path or None



    try:

        folder = filedialog.askdirectory(

            parent=root,

            title="Select a folder containing Excel files",

            initialdir=_last_dialog_dir or None,

            mustexist=True,

        )

    finally:

        root.destroy()



    if not folder:

        return None



    folder = os.path.normpath(folder)

    _last_dialog_dir = folder

    return folder





def scan_folder_for_excel(folder):

    """

    Returns a sorted list of full paths to .xlsx / .xlsm files directly

    inside 'folder'. Temporary Excel files (names starting with '~$') are

    ignored.

    """

    try:

        names = os.listdir(folder)

    except Exception as exc:

        error(f"Could not read the folder: {exc}")

        return []



    found = []

    for name in names:

        if name.startswith("~$"):

            continue

        if not name.lower().endswith(EXCEL_EXTENSIONS):

            continue

        full = os.path.join(folder, name)

        if os.path.isfile(full):

            found.append(full)



    found.sort(key=lambda p: os.path.basename(p).lower())

    return found





# ------------------------------------------------------------------------

# File / sheet loading & validation  (shared by both modes)

# ------------------------------------------------------------------------



def validate_and_open_excel(path):

    """

    Validates that the given path is a usable, non-empty Excel file.

    Returns a pandas.ExcelFile on success, or None on failure.

    """

    if not path:

        error("No path entered.")

        return None



    if not os.path.exists(path):

        error("File not found.")

        print()

        print("Please enter a valid Excel file path.")

        return None



    if not os.path.isfile(path):

        error("The path given is not a file.")

        return None



    ext = os.path.splitext(path)[1].lower()

    if ext not in EXCEL_EXTENSIONS:

        error(f"Unsupported file type '{ext}'. Only .xlsx and .xlsm files are supported.")

        return None



    try:

        excel_file = pd.ExcelFile(path, engine="openpyxl")

    except Exception as exc:

        error("Could not open the file. It may be corrupted or in use.")

        info(f"Details: {exc}")

        return None



    if not excel_file.sheet_names:

        error("The workbook does not contain any worksheets.")

        return None



    return excel_file





def get_sheet_preview_dims(path, sheet_name):

    """Fast row/column count for display purposes using openpyxl."""

    try:

        wb = load_workbook(path, read_only=True, data_only=True)

        ws = wb[sheet_name]

        rows = ws.max_row or 0

        cols = ws.max_column or 0

        data_rows = max(rows - 1, 0)

        wb.close()

        return data_rows, cols

    except Exception:

        return None, None





def load_sheet_dataframe(path, sheet_name):

    """

    Loads a sheet as a DataFrame of strings/objects (no type coercion),

    validates it isn't empty, and returns it (or None + printed error).

    """

    try:

        df = pd.read_excel(path, sheet_name=sheet_name, dtype=object, engine="openpyxl")

    except Exception as exc:

        error(f"Could not read worksheet '{sheet_name}'.")

        info(f"Details: {exc}")

        return None



    if df.shape[1] == 0:

        error(f"Worksheet '{sheet_name}' has no columns.")

        return None



    df.columns = [str(c).strip() for c in df.columns]



    if df.dropna(how="all").shape[0] == 0:

        error(f"Worksheet '{sheet_name}' contains no usable data.")

        return None



    return df





def display_file_information(path, excel_file):

    filename = os.path.basename(path)

    default_sheet = excel_file.sheet_names[0]

    rows, cols = get_sheet_preview_dims(path, default_sheet)



    print_banner("FILE LOADED")

    print(f"File name    : {filename}")

    print(f"File path    : {path}")

    if rows is not None:

        print(f"Total rows   : {rows:,}")

        print(f"Total columns: {cols:,}")

    print()

    print("Available sheets:")

    print()

    for idx, name in enumerate(excel_file.sheet_names, start=1):

        print(f"{idx}. {name}")





# ------------------------------------------------------------------------

# Sheet / column selection  (shared by both modes)

# ------------------------------------------------------------------------



def select_sheet(excel_file):

    """Single-sheet selection, used by Mode 1 (one sheet per file)."""

    sheet_names = excel_file.sheet_names



    if len(sheet_names) == 1:

        info(f"Only one worksheet found ('{sheet_names[0]}'). Using it automatically.")

        return sheet_names[0]



    print()

    print("Select sheet:")

    print()

    for idx, name in enumerate(sheet_names, start=1):

        print(f"{idx}. {name}")

    print()



    valid = [str(i) for i in range(1, len(sheet_names) + 1)]

    choice = ask_menu_choice("Enter selection: ", valid)

    return sheet_names[int(choice) - 1]





def parse_index_list(raw, max_idx):

    """

    Parses a comma-separated list of 1-based indices (e.g. "1,2,4").

    Returns a sorted-by-appearance list of valid ints, or None + printed

    error if the input is malformed / out of range / has duplicates.

    """

    parts = [p.strip() for p in raw.split(",") if p.strip() != ""]

    if not parts:

        error("No selection entered. Please try again.")

        return None



    try:

        indices = [int(p) for p in parts]

    except ValueError:

        error("Please enter numbers only (e.g. 1 or 1,2,4).")

        return None



    if any(i < 1 or i > max_idx for i in indices):

        error(f"Selection out of range. Enter numbers between 1 and {max_idx}.")

        return None



    if len(set(indices)) != len(indices):

        error("Duplicate numbers entered. Please select distinct items.")

        return None



    return indices





def select_multiple_sheets(excel_file):

    """

    Mode 2: lets the user pick 2+ worksheets (by number, comma separated)

    from a single workbook. Returns the list of chosen sheet names, in

    the order they were entered.

    """

    sheet_names = excel_file.sheet_names



    while True:

        print()

        print("Select the sheets you want to merge.")

        print()

        print("Example:")

        print("1,2,4")

        print()

        print("Available sheets:")

        print()

        for idx, name in enumerate(sheet_names, start=1):

            print(f"{idx}. {name}")

        print()



        raw = ask("Enter sheet numbers: ")

        indices = parse_index_list(raw, len(sheet_names))

        if indices is None:

            continue



        if len(indices) < 2:

            warning("Only one sheet selected. At least two sheets are recommended to merge, "

                     "but you may continue with just one.")



        selected_names = [sheet_names[i - 1] for i in indices]



        print()

        print("You selected:")

        print()

        for i in indices:

            print(f"{i}. {sheet_names[i - 1]}")

        print()

        print("Confirm selection?")

        print()

        print("1. Continue")

        print("2. Select again")

        print()

        choice = ask_menu_choice("Enter your choice: ", ["1", "2"])

        if choice == "1":

            return selected_names

        # else: loop back and select again





def display_columns(df, label=None):

    if label:

        print()

        print(f"{label}:")

        print()

    else:

        print_banner("AVAILABLE COLUMNS")

    for idx, col in enumerate(df.columns, start=1):

        print(f"{idx}. {col}")





def select_key_columns(df):

    """

    Prompts for one or more column numbers (comma-separated) to use as

    the merge/key column(s). Returns a list of column header names.

    """

    columns = list(df.columns)

    max_idx = len(columns)



    while True:

        print()

        print("Enter the column number(s) to use as the merge/key column.")

        print()

        print("Example:")

        print("1")

        print()

        print("For multiple columns:")

        print("1,2")

        print()

        raw = ask("Selection: ")



        indices = parse_index_list(raw, max_idx)

        if indices is None:

            continue



        return [columns[i - 1] for i in indices]





def confirm_key_columns(selected_cols):

    print()

    print("You selected:")

    print()

    for idx, col in enumerate(selected_cols, start=1):

        print(f"{idx}. {col}")

    print()

    print("Please confirm:")

    print()

    print("1. Continue")

    print("2. Select again")

    print()

    choice = ask_menu_choice("Enter your choice: ", ["1", "2"])

    return choice == "1"





def choose_key_columns_with_confirmation(df):

    """Runs select -> confirm in a loop until the user accepts a selection."""

    while True:

        selected = select_key_columns(df)

        if confirm_key_columns(selected):

            return selected





def validate_key_columns_present(df, key_columns, label):

    """

    Checks that every key column name exists in df's columns.

    Returns True if all present, otherwise prints an error and returns False.

    """

    missing = [c for c in key_columns if c not in df.columns]

    if not missing:

        return True



    error(f"{label} does not contain the selected key column(s):")

    print()

    for col in missing:

        print(f"  - {col}")

    print()

    print("Available columns:")

    for idx, col in enumerate(df.columns, start=1):

        print(f"{idx}. {col}")

    return False





# ------------------------------------------------------------------------

# Blank-row cleaning & column alignment  (shared by both modes)

# ------------------------------------------------------------------------



def is_blank_row(row):

    for v in row:

        if pd.isna(v):

            continue

        if isinstance(v, str) and v.strip() == "":

            continue

        return False

    return True





def clean_dataframe(df):

    """

    Drops fully-blank rows, keeping original row index for row-number tracing.



    Fast version: a row is kept as soon as ONE cell has content. Column by

    column, only the rows that are still undecided are examined, so for

    normal data (first column filled) almost no Python-level work is done.

    """

    n = len(df)

    if n == 0:

        return df



    has_content = np.zeros(n, dtype=bool)

    for c in range(df.shape[1]):

        undecided = np.flatnonzero(~has_content)

        if undecided.size == 0:

            break

        vals = df.iloc[undecided, c].to_numpy(dtype=object)

        notna = pd.notna(vals)

        for j in np.flatnonzero(notna):

            v = vals[j]

            if isinstance(v, str) and v.strip() == "":

                notna[j] = False

        has_content[undecided[notna]] = True



    return df[has_content]





def build_master_column_order(entries):

    """

    Builds the merged output's column order: the first source's headers

    first (in original order), then any additional headers encountered

    later, in the order first seen.

    """

    master = []

    seen = set()

    for entry in entries:

        for col in entry.df.columns:

            if col not in seen:

                seen.add(col)

                master.append(col)

    return master





# ------------------------------------------------------------------------

# Merge engine  (shared by both modes)

# ------------------------------------------------------------------------



def merge_sources(entries, master_columns, process_title, item_noun):

    """

    Cleans, aligns, and concatenates every SourceEntry's dataframe.

    process_title: banner text, e.g. "MERGE PROCESS" / "MERGE SHEETS PROCESS"

    item_noun: "Processing" label prefix stays constant; item_noun used

               only for readability of comments/logs here.

    Returns (merged_df, total_processed).



    Helper columns added to merged_df (never written to 'Merged Data'):

        __source__      unique source label (used to tell sources apart)

        __file__        source file name

        __folder__      source folder

        __sheet__       source worksheet name

        __source_row__  Excel row number in the source sheet

    """

    print_banner(process_title)



    merged_frames = []

    total_processed = 0

    n = len(entries)



    for i, entry in enumerate(entries, start=1):

        rows_found = entry.df.shape[0]

        cleaned = clean_dataframe(entry.df)

        rows_added = cleaned.shape[0]



        entry.clean_df = cleaned

        entry.rows_found = rows_found

        entry.rows_added = rows_added

        total_processed += rows_added



        print(f"[{i}/{n}] Processing {entry.label}")

        print(f"      Rows found : {rows_found:,}")

        print(f"      Rows added : {rows_added:,}")

        print()



        aligned = cleaned.reindex(columns=master_columns)

        aligned["__source__"] = entry.label

        aligned["__file__"] = entry.file_name

        aligned["__folder__"] = entry.folder

        aligned["__sheet__"] = entry.sheet_name

        aligned["__source_row__"] = cleaned.index.to_numpy() + 2

        merged_frames.append(aligned)



    merged_df = (

        pd.concat(merged_frames, ignore_index=True)

        if merged_frames

        else pd.DataFrame(columns=master_columns)

    )



    print(DASH)

    print(f"TOTAL ROWS PROCESSED: {total_processed:,}")

    print(DASH)



    return merged_df, total_processed





# ------------------------------------------------------------------------

# Duplicate checking  (shared by both modes)

# ------------------------------------------------------------------------



DUP_LABELS = np.array(["", "Cross-source", "Within-source", "Cross-source + Within-source"], dtype=object)

ACTION_LABELS = np.array(["", "Marked", "Kept", "Removed"], dtype=object)

LABEL_TO_CODE = {label: code for code, label in enumerate(DUP_LABELS)}





@dataclass

class DupInfo:

    """

    Result of the duplicate check, stored as arrays aligned with the rows

    of the merged dataframe (fast even for millions of rows).



    types: 0 = not duplicate, 1 = Cross-source, 2 = Within-source,

           3 = Cross-source + Within-source

    """

    keys: np.ndarray

    key_codes: np.ndarray

    src_codes: np.ndarray

    types: np.ndarray

    occurrences: np.ndarray

    n_cross: int

    n_within: int

    unique_count: int





def compute_key_array(merged_df, key_columns):

    """

    Key value of every merged row (blank cells -> "", values stripped;

    several key columns are joined with ' | '). Vectorised.

    """

    key = None

    for col in key_columns:

        s = merged_df[col]

        part = s.where(s.notna(), "").astype(str).str.strip()

        key = part if key is None else key + " | " + part

    return key.to_numpy(dtype=object)





def check_duplicates(merged_df, key_columns, source_noun):

    """

    source_noun: "file" or "sheet" - controls wording only.

    Returns a DupInfo. Same definitions as before:

      cross-source : same key in 2+ different sources

      within-source: same key 2+ times inside one source

    """

    print_banner("DUPLICATE CHECK")

    t0 = time.time()



    info("Building key values...")

    keys = compute_key_array(merged_df, key_columns)

    n = len(keys)



    info("Finding duplicates...")

    key_codes, _key_uniques = pd.factorize(keys)

    src_codes, src_uniques = pd.factorize(merged_df["__source__"].to_numpy())

    n_src = max(len(src_uniques), 1)

    n_keys = int(key_codes.max()) + 1 if n else 0

    valid = keys != ""



    types = np.zeros(n, dtype=np.int8)

    occurrences = np.zeros(n, dtype=np.int64)

    n_cross = n_within = unique_count = 0



    if n and valid.any():

        kc = key_codes[valid]

        sc = src_codes[valid]



        total_per_key = np.bincount(kc, minlength=n_keys)

        pair = kc.astype(np.int64) * n_src + sc

        pair_uniq, pair_counts = np.unique(pair, return_counts=True)

        sources_per_key = np.bincount(pair_uniq // n_src, minlength=n_keys)



        cross_key = sources_per_key > 1

        in_source_count = pair_counts[np.searchsorted(pair_uniq, pair)]

        is_cross = cross_key[kc]

        is_within = in_source_count > 1



        types[valid] = np.where(

            is_cross & is_within, 3,

            np.where(is_cross, 1, np.where(is_within, 2, 0)),

        ).astype(np.int8)

        occurrences[valid] = total_per_key[kc]



        n_cross = int(cross_key.sum())

        n_within = int((pair_counts > 1).sum())

        unique_count = int((total_per_key > 0).sum())



    if n_cross:

        print(f"DUPLICATE DATA: {n_cross:,} cross-{source_noun} duplicate(s) found.")

    else:

        ok(f"No duplicate key values found across the selected {source_noun}s.")



    if n_within:

        print(f"DUPLICATE DATA: {n_within:,} within-{source_noun} duplicate(s) found.")



    print_banner("DUPLICATE SUMMARY")

    print(f"{source_noun.capitalize()}s checked      : {len(src_uniques)}")

    print(f"Unique key values    : {unique_count:,}")

    print(f"Duplicate key values : {n_cross:,}")

    print()

    print(f"Cross-{source_noun} duplicates:")

    print(f"{n_cross:,}")

    print()

    print(f"Within-{source_noun} duplicates:")

    print(f"{n_within:,}")

    info(f"Duplicate check finished in {time.time() - t0:,.0f}s.")



    return DupInfo(keys, key_codes, src_codes, types, occurrences,

                   n_cross, n_within, unique_count)





def ask_duplicate_handling():

    """

    Last question before the merge starts: what to do with duplicate rows.

    Returns "mark", "delete_first" or "delete_last".

    """

    print_banner("DUPLICATE HANDLING")

    print("What should happen to duplicate rows (same key value)?")

    print()

    print("1. Keep all rows, but MARK duplicates with a cell colour")

    print("2. DELETE duplicates (keep only one row per key value)")

    print()

    choice = ask_menu_choice("Enter your choice: ", ["1", "2"])

    if choice == "1":

        return "mark"



    print()

    print("Which row should be kept for each duplicated key?")

    print()

    print("1. Keep the FIRST occurrence (earliest file in your list / top row)")

    print("2. Keep the LAST occurrence")

    print()

    keep = ask_menu_choice("Enter your choice: ", ["1", "2"])

    return "delete_first" if keep == "1" else "delete_last"





def describe_duplicate_handling(mode, removed):

    if mode == "mark":

        return "Duplicate rows marked with colour (nothing deleted)"

    which = "first" if mode == "delete_first" else "last"

    return f"{removed:,} duplicate row(s) deleted (kept the {which} occurrence)"





def apply_duplicate_handling(merged_df, dup, mode):

    """

    Applies the user's duplicate choice.



    Returns (final_df, actions, removed_count). actions is an int8 array

    aligned with the ORIGINAL merged rows:

      0 = "", 1 = Marked, 2 = Kept, 3 = Removed

    final_df is merged_df with removed rows dropped (mark mode: unchanged).

    """

    n = len(dup.types)

    has_dup = dup.types > 0

    actions = np.zeros(n, dtype=np.int8)



    if mode == "mark":

        actions[has_dup] = 1

        return merged_df, actions, 0



    idx = np.flatnonzero(has_dup)

    keep = "first" if mode == "delete_first" else "last"

    is_extra = pd.Series(dup.key_codes[idx]).duplicated(keep=keep).to_numpy()

    actions[idx] = np.where(is_extra, 3, 2)



    keep_rows = actions != 3

    removed = int((~keep_rows).sum())

    final_df = merged_df[keep_rows].reset_index(drop=True)

    return final_df, actions, removed





def build_duplicate_report_df(merged_df, dup, actions, master_columns):

    """

    Detailed, long-format duplicate report. One line per duplicated ROW:



      Key Value | Source File | Source Folder | Source Sheet | Source Row |

      Occurrences | Duplicate Type | Action | \<ALL ORIGINAL DATA COLUMNS...>



    Duplicate Type:

      "Cross-source"                 key appears in 2+ different sources

      "Within-source"                key appears 2+ times in the same source only

      "Cross-source + Within-source" key is cross-source AND this row's own

                                     source also holds it more than once

    Occurrences = total times the key appears in the whole merge.

    Action = Marked (kept + coloured) / Kept / Removed, per the user's choice.

    Rows are grouped by key (in order of first appearance), then by source.

    """

    pos = np.flatnonzero(dup.types)

    if pos.size == 0:

        return pd.DataFrame([{"Key Value": "No duplicates found."}])



    info(f"Building Duplicate Report ({pos.size:,} rows)...")

    order = np.lexsort((pos, dup.src_codes[pos], dup.key_codes[pos]))

    pos = pos[order]

    dup_rows = merged_df.iloc[pos]



    front = pd.DataFrame({

        "Key Value": dup.keys[pos],

        "Source File": dup_rows["__file__"].to_numpy(),

        "Source Folder": dup_rows["__folder__"].to_numpy(),

        "Source Sheet": dup_rows["__sheet__"].to_numpy(),

        "Source Row": dup_rows["__source_row__"].to_numpy(),

        "Occurrences": dup.occurrences[pos],

        "Duplicate Type": DUP_LABELS[dup.types[pos]],

        "Action": ACTION_LABELS[actions[pos]],

    })



    # The actual data of each duplicated row. If an original column happens

    # to share a name with one of the report's own columns, label it clearly.

    data = dup_rows[master_columns].reset_index(drop=True)

    data.columns = [

        f"{c} (data)" if c in REPORT_FIXED_COLUMNS else c

        for c in data.columns

    ]



    return pd.concat([front, data], axis=1)





# ------------------------------------------------------------------------

# Output saving  (shared by both modes)

# ------------------------------------------------------------------------



def _fill_row(ws, excel_row, n_cols, hex_color):

    """openpyxl fallback: colour every cell of one row."""

    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    for c in range(1, n_cols + 1):

        ws.cell(row=excel_row, column=c).fill = fill





def ask_output_layout():

    """

    Mode 1: how should the merged data be laid out in the output workbook?

    Returns "single" or "folder".

    """

    print_banner("OUTPUT LAYOUT")

    print("Excel allows about 1 million rows per sheet. How should the merged")

    print("data be arranged in the output file?")

    print()

    print("1. One 'Merged Data' sheet (continues on 'Merged Data 2', ... only if it gets too big)")

    print("2. One sheet PER FOLDER (each sheet is named after its folder)")

    print()

    choice = ask_menu_choice("Enter your choice: ", ["1", "2"])

    return "single" if choice == "1" else "folder"





def make_sheet_name(raw, used):

    """Excel-safe, unique (case-insensitive) sheet name of max 31 characters."""

    base = re.sub(r"[**\\\**/*?:**\[\\]**]", "_", str(raw)).strip().strip("'") or "Sheet"

    base = base[:31]

    candidate = base

    n = 2

    while candidate.lower() in used:

        suffix = f" ({n})"

        candidate = base[:31 - len(suffix)] + suffix

        n += 1

    used.add(candidate.lower())

    return candidate





def split_into_chunks(positions):

    """Splits an array of row positions into pieces that fit on one sheet."""

    limit = EXCEL_MAX_DATA_ROWS

    if len(positions) == 0:

        return [positions]

    return [positions[i:i + limit] for i in range(0, len(positions), limit)]





def plan_data_sheets(merged_df, layout, used):

    """

    Decides which rows go on which sheet.

    Returns a list of (sheet_name, numpy array of row positions).

    """

    plan = []



    if layout == "folder":

        codes, folders = pd.factorize(merged_df["__folder__"].to_numpy())

        for i, folder in enumerate(folders):

            positions = np.flatnonzero(codes == i)

            folder_name = os.path.basename(str(folder).rstrip("\\\\/")) or str(folder)

            for k, chunk in enumerate(split_into_chunks(positions)):

                raw = folder_name if k == 0 else f"{folder_name} {k + 1}"

                plan.append((make_sheet_name(raw, used), chunk))

    else:

        for k, chunk in enumerate(split_into_chunks(np.arange(len(merged_df)))):

            raw = "Merged Data" if k == 0 else f"Merged Data {k + 1}"

            plan.append((make_sheet_name(raw, used), chunk))



    return plan





def _color_rows_openpyxl(writer, sheet_name, type_codes, n_cols):

    """openpyxl fallback: colour the rows of one sheet by duplicate type code."""

    rows = np.flatnonzero(type_codes > 0)

    if rows.size == 0:

        return

    ws = writer.sheets[sheet_name]

    for r, c in zip((rows + 2).tolist(), type_codes[rows].tolist()):

        _fill_row(ws, r, n_cols, DUP_COLORS[DUP_LABELS[c]])





def _write_sheet_fast(book, sheet_name, df, type_codes, fmts, header_fmt):

    """

    Writes one dataframe to a new worksheet with xlsxwriter, row by row

    (much faster and lighter on memory than pandas.to_excel).

    type_codes (optional int array, one per row): 1/2/3 -> row colour.

    """

    ws = book.add_worksheet(sheet_name)

    ws.write_row(0, 0, [str(c) for c in df.columns], header_fmt)



    vals = df.to_numpy(dtype=object)

    if vals.size:

        vals[pd.isna(vals)] = None

    codes = type_codes.tolist() if type_codes is not None else None



    total = len(vals)

    progress_step = 250_000

    for i in range(total):

        if codes is not None and codes[i]:

            ws.set_row(i + 1, None, fmts[codes[i]])

        row = vals[i]

        try:

            ws.write_row(i + 1, 0, row)

        except Exception:

            # An unusual cell value: write that row cell by cell as text.

            for j, v in enumerate(row):

                try:

                    ws.write(i + 1, j, v)

                except Exception:

                    ws.write_string(i + 1, j, str(v))

        if (i + 1) % progress_step == 0:

            info(f"   ... {i + 1:,} / {total:,} rows")





def save_output(merged_df, duplicate_report_df, master_columns, filename_prefix,

                row_types=None, layout="single"):

    """

    layout: "single" -> 'Merged Data' (continues on 'Merged Data 2', ... if

                         the Excel row limit would be exceeded)

            "folder" -> one sheet per source folder (also split if needed)

    row_types: optional int array (aligned with merged_df rows) of duplicate

               type codes; those rows get a cell colour in the data sheets.

    The 'Duplicate Report' rows are always coloured by Duplicate Type, and

    the report is also continued on 'Duplicate Report 2', ... if too long.



    Uses xlsxwriter when installed (fast); otherwise falls back to

    pandas + openpyxl (works, but MUCH slower on big data).

    """

    # Set OUTPUT_DIR above to a folder on your own computer.
    # If left blank, the result is saved in the folder where this script is run.
    output_dir = OUTPUT_DIR.strip() or os.getcwd()
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, filename)



    n_cols = len(master_columns)



    used = {"duplicate report"}  # reserved so a folder can't take this name

    data_plan = plan_data_sheets(merged_df, layout, used)



    report_names = []

    report_chunks = split_into_chunks(np.arange(len(duplicate_report_df)))

    for k in range(len(report_chunks)):

        if k == 0:

            report_names.append("Duplicate Report")

        else:

            report_names.append(make_sheet_name(f"Duplicate Report {k + 1}", used))



    if len(data_plan) > 1:

        info(f"Merged data is being written to {len(data_plan)} sheets "

             f"(Excel allows max {EXCEL_MAX_DATA_ROWS + 1:,} rows per sheet).")



    def report_codes(chunk):

        if "Duplicate Type" not in chunk.columns:

            return None

        return chunk["Duplicate Type"].map(LABEL_TO_CODE).fillna(0).astype(int).to_numpy()



    try:

        import xlsxwriter

        fast = True

    except ImportError:

        fast = False

        warning("'xlsxwriter' is not installed, so saving will be MUCH slower.")

        warning("Install it with:  pip install xlsxwriter")



    t0 = time.time()



    if fast:

        info("Writing Excel file with 'xlsxwriter' (large files can take a while)...")

        book = xlsxwriter.Workbook(output_path, {

            "constant_memory": True,          # stream to disk, low RAM

            "use_zip64": True,                # allow very large sheets

            "strings_to_formulas": False,     # keep text like '=abc' as text

            "strings_to_urls": False,

            "nan_inf_to_errors": True,

            "remove_timezone": True,

            "default_date_format": "yyyy-mm-dd hh:mm:ss",

        })

        fmts = {

            c: book.add_format({"bg_color": "#" + DUP_COLORS[DUP_LABELS[c]]})

            for c in (1, 2, 3)

        }

        header_fmt = book.add_format({"bold": True})



        for sheet_name, positions in data_plan:

            info(f"Writing sheet '{sheet_name}' ({len(positions):,} rows)...")

            chunk = merged_df.iloc[positions][master_columns]

            codes = row_types[positions] if row_types is not None else None

            _write_sheet_fast(book, sheet_name, chunk, codes, fmts, header_fmt)



        for sheet_name, positions in zip(report_names, report_chunks):

            info(f"Writing sheet '{sheet_name}' ({len(positions):,} rows)...")

            chunk = duplicate_report_df.iloc[positions]

            _write_sheet_fast(book, sheet_name, chunk, report_codes(chunk), fmts, header_fmt)



        info("Saving file (finalising)...")

        book.close()

    else:

        info("Writing Excel file with 'openpyxl' (very slow for large data)...")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

            for sheet_name, positions in data_plan:

                info(f"Writing sheet '{sheet_name}' ({len(positions):,} rows)...")

                chunk = merged_df.iloc[positions][master_columns]

                chunk.to_excel(writer, sheet_name=sheet_name, index=False)

                if row_types is not None:

                    _color_rows_openpyxl(writer, sheet_name, row_types[positions], n_cols)



            for sheet_name, positions in zip(report_names, report_chunks):

                info(f"Writing sheet '{sheet_name}' ({len(positions):,} rows)...")

                chunk = duplicate_report_df.iloc[positions]

                chunk.to_excel(writer, sheet_name=sheet_name, index=False)

                codes = report_codes(chunk)

                if codes is not None:

                    _color_rows_openpyxl(writer, sheet_name, codes, chunk.shape[1])



    info(f"Excel file written in {time.time() - t0:,.0f}s.")

    ok("Sheets written: " + ", ".join([n for n, _ in data_plan] + report_names))

    return output_path, len(merged_df)





# ------------------------------------------------------------------------

# MODE 1 - Merge multiple Excel FILES (file / folder pickers)

# ------------------------------------------------------------------------



def show_current_files(entries):

    print_section("CURRENT FILES TO MERGE")

    print()

    for idx, e in enumerate(entries, start=1):

        print(f"{idx}. {e.path} | {e.sheet_name}")





def select_sheets_for_file(excel_file):

    """

    Mode 1: lets the user pick ONE OR MORE worksheets from a single file

    (e.g. "2" or "1,3"; A = all sheets). Returns a list of sheet names.

    """

    sheet_names = excel_file.sheet_names



    if len(sheet_names) == 1:

        info(f"Only one worksheet found ('{sheet_names[0]}'). Using it automatically.")

        return list(sheet_names)



    while True:

        print()

        print("Select the sheet(s) to merge from this file.")

        print("Example: 2   |   1,3   |   A = all sheets")

        print()

        for idx, name in enumerate(sheet_names, start=1):

            print(f"{idx}. {name}")

        print()

        raw = ask("Enter sheet number(s): ")



        if raw.lower() in ("a", "all"):

            return list(sheet_names)



        indices = parse_index_list(raw, len(sheet_names))

        if indices is None:

            continue

        return [sheet_names[i - 1] for i in indices]





def add_sheet_entry(path, filename, folder, sheet_name, entries, key_columns):

    """

    Loads one worksheet of a file and validates it (key columns), then

    appends a SourceEntry. The very first source also sets the key

    column(s). A later sheet that lacks the key column(s) can be removed

    (skipped) or the whole merge cancelled.

    Returns (key_columns, cancelled).

    """

    if any(e.path == path and e.sheet_name == sheet_name for e in entries):

        warning(f"'{filename}' | '{sheet_name}' is already in the list. Skipped.")

        return key_columns, False



    info(f"Reading worksheet: {sheet_name}")

    df = load_sheet_dataframe(path, sheet_name)

    if df is None:

        return key_columns, False



    if key_columns is None:

        display_columns(df)

        key_columns = choose_key_columns_with_confirmation(df)

    else:

        info("Detecting columns...")

        if not validate_key_columns_present(df, key_columns, f"'{filename}' | '{sheet_name}'"):

            warning("This sheet does NOT contain the required key column(s). "

                    "Merging it could produce incorrect data.")

            print()

            print("What would you like to do?")

            print()

            print("1. Remove this sheet (skip it)")

            print("2. Cancel the whole merge")

            print()

            choice = ask_menu_choice("Enter your choice: ", ["1", "2"])

            if choice == "2":

                return key_columns, True

            info(f"Skipped: {filename} | {sheet_name}")

            return key_columns, False



    entries.append(SourceEntry(

        label=f"{path} | {sheet_name}",

        df=df,

        path=path,

        file_name=filename,

        folder=folder,

        sheet_name=sheet_name,

    ))



    print()

    print("Added:")

    print(path)

    print(f"Sheet: {sheet_name}")

    return key_columns, False





def add_source_from_path(path, entries, key_columns):

    """

    Opens one Excel file for Mode 1, shows its sheets, lets the user pick

    one OR MORE worksheets, and adds each chosen sheet as its own source.



    Returns (key_columns, cancelled).

    """

    path = os.path.abspath(path)

    excel_file = validate_and_open_excel(path)

    if excel_file is None:

        return key_columns, False



    filename = os.path.basename(path)

    folder = os.path.dirname(path)



    display_file_information(path, excel_file)

    sheet_names = select_sheets_for_file(excel_file)



    for sheet_name in sheet_names:

        key_columns, cancelled = add_sheet_entry(

            path, filename, folder, sheet_name, entries, key_columns

        )

        if cancelled:

            return key_columns, True



    return key_columns, False





def choose_files_from_folder(folder, entries):

    """

    Lists the Excel files found in 'folder' and lets the user pick which

    ones to include. Returns a list of full paths (possibly empty).

    """

    files = scan_folder_for_excel(folder)



    print_section(f"EXCEL FILES IN: {folder}")

    if not files:

        warning("No .xlsx or .xlsm files found in this folder.")

        return []



    already = {e.path for e in entries}

    print()

    for idx, p in enumerate(files, start=1):

        note = "  (already added)" if os.path.abspath(p) in already else ""

        print(f"{idx}. {os.path.basename(p)}{note}")



    while True:

        print()

        print("Select the files to include.")

        print("Example: 1,3   |   A = all files   |   0 = go back")

        print()

        raw = ask("Enter file numbers: ")



        if raw == "0":

            return []

        if raw.lower() in ("a", "all"):

            return list(files)



        indices = parse_index_list(raw, len(files))

        if indices is None:

            continue

        return [files[i - 1] for i in indices]





def get_file_entries_for_mode1():

    """

    Main input loop for Mode 1. Files/folders can be added in any order and

    from any number of different folders.

    Returns (entries, key_columns) or (None, None) if cancelled.

    """

    print_banner("MERGE EXCEL FILES")



    entries = []

    key_columns = None



    while True:

        print()

        print("Choose input:")

        print()

        print("1. Select Excel File")

        print("2. Select Folder")

        print("3. Continue to merge process")

        print("4. Cancel")

        print()

        choice = ask_menu_choice("Enter your choice: ", ["1", "2", "3", "4"])



        added_something = False

        cancelled = False



        if choice == "1":

            path = pick_excel_file()

            if not path:

                info("No file selected.")

                continue

            print()

            print("Selected file:")

            print(path)

            before = len(entries)

            key_columns, cancelled = add_source_from_path(path, entries, key_columns)

            added_something = len(entries) > before



        elif choice == "2":

            folder = pick_folder()

            if not folder:

                info("No folder selected.")

                continue

            print()

            print("Selected folder:")

            print(folder)

            chosen = choose_files_from_folder(folder, entries)

            before = len(entries)

            for path in chosen:

                key_columns, cancelled = add_source_from_path(path, entries, key_columns)

                if cancelled:

                    break

            added_something = len(entries) > before



        elif choice == "3":

            if not entries:

                warning("No files have been selected yet.")

                continue

            if len(entries) < 2:

                warning("Only one file/sheet was selected. Merging will proceed with a single source.")

            break



        else:  # "4"

            info("Merge cancelled.")

            return None, None



        if cancelled:

            info("Merge cancelled.")

            return None, None



        if added_something:

            show_current_files(entries)

            print()

            print("What next?")

            print()

            print("1. Add another file/folder")

            print("2. Continue to merge process")

            print("3. Cancel")

            print()

            nxt = ask_menu_choice("Enter your choice: ", ["1", "2", "3"])

            if nxt == "2":

                if len(entries) < 2:

                    warning("Only one file/sheet was selected. Merging will proceed with a single source.")

                break

            if nxt == "3":

                info("Merge cancelled.")

                return None, None

            # nxt == "1": loop back to the input menu



    return entries, key_columns





def display_final_summary_mode1(entries, total_processed, output_rows, cross_count, within_count, output_path, dup_note):

    n_files = len({e.path for e in entries})

    n_folders = len({e.folder for e in entries})

    print_banner("MERGE FILES COMPLETED")

    print(f"Files merged          : {n_files}")

    print(f"Folders used          : {n_folders}")

    print(f"Sources (file+sheet)  : {len(entries)}")

    print(f"Total rows processed  : {total_processed:,}")

    print(f"Total rows in output  : {output_rows:,}")

    print()

    print(f"Duplicate key values  : {cross_count:,}")

    print(f"Cross-file duplicates : {cross_count:,}")

    print(f"Within-file duplicates: {within_count:,}")

    print(f"Duplicate handling    : {dup_note}")

    print()

    print("Output:")

    print()

    print(output_path)

    print()

    ok("Process completed successfully.")

    print(LINE)





def run_mode1():

    entries, key_columns = get_file_entries_for_mode1()

    if entries is None:

        return  # user cancelled back to main menu



    layout = ask_output_layout()

    dup_mode = ask_duplicate_handling()



    master_columns = build_master_column_order(entries)

    merged_df, total_processed = merge_sources(entries, master_columns, "MERGE PROCESS", "file")

    dup = check_duplicates(merged_df, key_columns, "file")

    final_df, actions, removed = apply_duplicate_handling(merged_df, dup, dup_mode)

    duplicate_report_df = build_duplicate_report_df(merged_df, dup, actions, master_columns)

    output_path, output_rows = save_output(

        final_df, duplicate_report_df, master_columns, "Merged_Files",

        row_types=dup.types if dup_mode == "mark" else None, layout=layout,

    )



    display_final_summary_mode1(

        entries, total_processed, output_rows, dup.n_cross, dup.n_within, output_path,

        describe_duplicate_handling(dup_mode, removed),

    )





# ------------------------------------------------------------------------

# MODE 2 - Merge multiple SHEETS from one Excel file

# ------------------------------------------------------------------------



def get_sheet_entries_for_mode2():

    """

    Main flow for Mode 2: pick one file, pick 2+ sheets, load each,

    display info/columns, choose key columns from the first sheet and

    validate against the rest.

    Returns (entries, key_columns, source_path) or (None, None, None) if cancelled.

    """

    print_banner("MERGE SHEETS FROM ONE EXCEL FILE")



    while True:

        path = ask_path("Enter Excel file path: ")

        excel_file = validate_and_open_excel(path)

        if excel_file is not None:

            break

        retry = ask_menu_choice(

            "\n1. Try a different file path\n2. Cancel and return to main menu\n\nEnter your choice: ",

            ["1", "2"],

        )

        if retry == "2":

            return None, None, None



    abs_path = os.path.abspath(path)

    filename = os.path.basename(abs_path)

    folder = os.path.dirname(abs_path)

    ok(f"File loaded: {filename}")

    print()

    print("Available worksheets:")

    print()

    for idx, name in enumerate(excel_file.sheet_names, start=1):

        print(f"{idx}. {name}")



    if len(excel_file.sheet_names) < 2:

        error("This workbook only has one worksheet - there is nothing to merge.")

        return None, None, None



    while True:

        selected_sheet_names = select_multiple_sheets(excel_file)



        # Load every selected sheet.

        entries = []

        load_failed = False

        for name in selected_sheet_names:

            info(f"Loading sheet '{name}'...")

            df = load_sheet_dataframe(path, name)

            if df is None:

                load_failed = True

                break

            ok(f"Sheet '{name}' loaded ({df.shape[0]:,} rows, {df.shape[1]:,} columns).")

            entries.append(SourceEntry(

                label=name,

                df=df,

                path=abs_path,

                file_name=filename,

                folder=folder,

                sheet_name=name,

            ))



        if load_failed:

            error("One of the selected sheets could not be used. Please select sheets again.")

            continue



        # Display per-sheet information & columns (section 7 of the spec).

        print_banner("SHEET INFORMATION")

        for entry in entries:

            rows, cols = get_sheet_preview_dims(path, entry.label)

            print()

            print(f"Sheet: {entry.label}")

            if rows is not None:

                print(f"Rows: {rows:,}")

                print(f"Columns: {cols:,}")

            else:

                print(f"Rows: {entry.df.shape[0]:,}")

                print(f"Columns: {entry.df.shape[1]:,}")



        for entry in entries:

            display_columns(entry.df, label=entry.label)



        # Choose key columns from the first selected sheet, then validate

        # every other selected sheet has those same headers.

        first_entry = entries[0]

        key_columns = choose_key_columns_with_confirmation(first_entry.df)



        all_valid = True

        for entry in entries[1:]:

            if not validate_key_columns_present(entry.df, key_columns, f"Sheet '{entry.label}'"):

                all_valid = False

                break



        if not all_valid:

            retry = ask_menu_choice(

                "\n1. Select sheets/key columns again\n2. Cancel and return to main menu\n\nEnter your choice: ",

                ["1", "2"],

            )

            if retry == "2":

                return None, None, None

            continue



        return entries, key_columns, path





def display_final_summary_mode2(path, entries, total_processed, output_rows, cross_count, within_count, output_path, dup_note):

    print_banner("MERGE SHEETS COMPLETED")

    print(f"Source file            : {os.path.basename(path)}")

    print()

    print(f"Sheets merged          : {len(entries)}")

    print()

    print("Sheets:")

    for entry in entries:

        print(f"- {entry.label}")

    print()

    print(f"Total rows processed   : {total_processed:,}")

    print(f"Total rows in output   : {output_rows:,}")

    print()

    print(f"Duplicate key values   : {cross_count:,}")

    print(f"Cross-sheet duplicates : {cross_count:,}")

    print(f"Within-sheet duplicates: {within_count:,}")

    print(f"Duplicate handling     : {dup_note}")

    print()

    print("Output:")

    print()

    print(output_path)

    print()

    ok("Process completed successfully.")

    print(LINE)





def run_mode2():

    entries, key_columns, source_path = get_sheet_entries_for_mode2()

    if entries is None:

        return  # user cancelled back to main menu



    dup_mode = ask_duplicate_handling()



    master_columns = build_master_column_order(entries)

    merged_df, total_processed = merge_sources(entries, master_columns, "MERGE SHEETS PROCESS", "sheet")

    dup = check_duplicates(merged_df, key_columns, "sheet")

    final_df, actions, removed = apply_duplicate_handling(merged_df, dup, dup_mode)

    duplicate_report_df = build_duplicate_report_df(merged_df, dup, actions, master_columns)

    output_path, output_rows = save_output(

        final_df, duplicate_report_df, master_columns, "Merged_Sheets",

        row_types=dup.types if dup_mode == "mark" else None,

    )



    display_final_summary_mode2(

        source_path, entries, total_processed, output_rows, dup.n_cross, dup.n_within, output_path,

        describe_duplicate_handling(dup_mode, removed),

    )





# ------------------------------------------------------------------------

# Main menu

# ------------------------------------------------------------------------



def show_main_menu():

    print_banner("PYTHON EXCEL MERGE TOOL")

    print()

    print("Please select the operation:")

    print()

    print("1. Merge multiple Excel files")

    print("2. Merge multiple sheets from one Excel file")

    print("3. Exit")

    print()

    return ask_menu_choice("Enter your choice: ", ["1", "2", "3"])





def main():

    try:

        while True:

            choice = show_main_menu()

            if choice == "1":

                run_mode1()

            elif choice == "2":

                run_mode2()

            else:

                info("Goodbye.")

                break



            print()

            again = ask_menu_choice(

                "Return to main menu? (1. Yes  2. Exit)\n\nEnter your choice: ", ["1", "2"]

            )

            if again == "2":

                info("Goodbye.")

                break



    except SystemExit:

        raise

    except Exception:

        print()

        error("An unexpected problem occurred. Details below (for troubleshooting):")

        print()

        traceback.print_exc()

        sys.exit(1)





if __name__ == "__main__":

    main()
