import pandas as pd
import os
import argparse
import csv
import re

# Constants
PARAM_MARKS = {
    "Timely completion, punctuality": 2,
    "Performance, involvement, efficiency": 4,
    "Oral Presentation": 3,
    "Documentation, neatness": 1,
}
SESSION_PREFIX = "Session "  # maps experiment_1.xlsx -> Session 1
ROLL_ALIASES = ["Roll", "Roll No", "Roll No.", "RollNumber"]
TOTAL_COL_NAME = "Total"

def extract_roll_number(roll):
    if pd.isna(roll):
        return None
    s = str(roll)
    m = re.search(r"(\d+)", s)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None

def normalize_presence(val):
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return val != 0
    s = str(val).strip().lower()
    return s in {"p", "present", "yes", "y", "1", "attended", "true"}

def load_attendance(path):
    """
    Parses attendance CSV with multi-row header as in the provided sample.
    Returns DataFrame with Session 1..N boolean flags and '__roll_num_key' integer.
    """
    with open(path, encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        rows = list(reader)

    # Locate header row (which has 'Roll No.' etc.)
    header_idx = None
    for i, row in enumerate(rows):
        if len(row) >= 1 and any(isinstance(c, str) and c.strip().lower().startswith("roll") for c in row[:3]):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find attendance header row with 'Roll No.'")

    # Data starts after header + two subheader rows
    data_start = header_idx + 3
    raw_rows = rows[data_start:]

    cleaned = []
    for row in raw_rows:
        if not row or len(row) < 3:
            continue
        roll_raw = row[0].strip()
        name = row[1].strip()
        surname = row[2].strip()
        session_presence = {}
        # Assume 4 sessions, each has Attendance at indices 3,5,7,9 etc.
        for session_num in range(1, 5):
            att_idx = 3 + (session_num - 1) * 2
            val = row[att_idx] if att_idx < len(row) else None
            session_presence[f"{SESSION_PREFIX}{session_num}"] = normalize_presence(val)
        cleaned.append({
            "Roll No.": roll_raw,
            "Name": name,
            "Surname": surname,
            **session_presence
        })

    df = pd.DataFrame(cleaned)
    df["__roll_num_key"] = df["Roll No."].apply(extract_roll_number)
    return df

def find_roll_column(df):
    # exact alias match first
    for alias in ROLL_ALIASES:
        for col in df.columns:
            if col.strip().lower() == alias.strip().lower():
                return col
    # fallback: any column containing 'roll'
    for col in df.columns:
        if "roll" in col.strip().lower():
            return col
    return None

def process_experiment(excel_path, attendance_df, session_col, output_path=None):
    exp_df = pd.read_excel(excel_path)

    # Identify roll column
    roll_col = find_roll_column(exp_df)
    if roll_col is None:
        raise KeyError(f"No roll number column found in '{excel_path}'.")

    # Build numeric roll key
    exp_df["__roll_num_key"] = exp_df[roll_col].apply(extract_roll_number)

    # Attendance presence map for this session
    if session_col not in attendance_df.columns:
        raise KeyError(f"'{session_col}' not found in attendance data.")
    attendance_map = attendance_df.set_index("__roll_num_key")[session_col].to_dict()

    # Ensure parameter columns exist (append if missing)
    for param in PARAM_MARKS:
        if param not in exp_df.columns:
            exp_df[param] = 0

    # Ensure Total exists
    if TOTAL_COL_NAME not in exp_df.columns:
        exp_df[TOTAL_COL_NAME] = 0

    # In-place mark assignment preserving column order
    for idx, row in exp_df.iterrows():
        key = row.get("__roll_num_key", None)
        if key is None or pd.isna(key):
            continue  # no valid roll
        if key not in attendance_map:
            continue  # no attendance info; leave existing marks
        present = normalize_presence(attendance_map.get(key))
        # overwrite parameter columns
        for param, mark in PARAM_MARKS.items():
            exp_df.at[idx, param] = mark if present else 0
        # recompute total
        total_score = sum(exp_df.at[idx, param] for param in PARAM_MARKS)
        exp_df.at[idx, TOTAL_COL_NAME] = total_score

    # Drop helper column without disturbing order
    exp_df.drop(columns=["__roll_num_key"], inplace=True, errors="ignore")

    # Save
    out_path = output_path if output_path else excel_path
    exp_df.to_excel(out_path, index=False)
    print(f"[+] Processed '{os.path.basename(excel_path)}' (session '{session_col}') → saved to '{out_path}'")

def main():
    parser = argparse.ArgumentParser(description="Apply attendance-based marks to experiment sheets using roll number matching.")
    parser.add_argument("--attendance", required=True, help="Path to attendance CSV file")
    parser.add_argument("--experiments-dir", required=True, help="Directory containing experiment_1.xlsx ...")
    parser.add_argument("--output-dir", help="Optional directory to write updated experiment files (if omitted, overwrites originals)")
    args = parser.parse_args()

    if not os.path.exists(args.attendance):
        raise FileNotFoundError(f"Attendance file '{args.attendance}' not found.")
    attendance_df = load_attendance(args.attendance)
    if "__roll_num_key" not in attendance_df.columns:
        raise RuntimeError("Attendance parsing failed to create roll key.")

    for fname in os.listdir(args.experiments_dir):
        if not fname.lower().startswith("experiment_") or not fname.lower().endswith((".xls", ".xlsx")):
            continue
        try:
            num = os.path.splitext(fname)[0].split("_")[1]
            session_col = f"{SESSION_PREFIX}{num}"
        except Exception:
            print(f"[!] Skipping '{fname}' due to unexpected naming.")
            continue

        exp_path = os.path.join(args.experiments_dir, fname)
        if args.output_dir:
            os.makedirs(args.output_dir, exist_ok=True)
            out_path = os.path.join(args.output_dir, fname)
        else:
            out_path = exp_path

        try:
            process_experiment(
                excel_path=exp_path,
                attendance_df=attendance_df,
                session_col=session_col,
                output_path=out_path
            )
        except Exception as e:
            print(f"[!] Error processing '{fname}': {e}")

if __name__ == "__main__":
    main()


