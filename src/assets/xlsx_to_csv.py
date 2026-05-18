import pandas as pd
import os
import glob
import sys


def xlsx_to_csv(input_path, output_path=None, encoding="utf-8-sig"):
    """
    Convert a single .xlsx / .xls file to CSV.

    Returns:
        List of (output_path, rows, cols) tuples on success, or empty list on error.
    """
    if not os.path.exists(input_path):
        print(f"[ERROR] File not found: {input_path}")
        return []

    if not input_path.lower().endswith((".xlsx", ".xls")):
        print(f"[ERROR] Not an Excel file: {input_path}")
        return []

    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".csv"

    try:
        excel_file = pd.ExcelFile(input_path)
        sheet_names = excel_file.sheet_names
        exported = []

        if len(sheet_names) == 1:
            df = pd.read_excel(input_path, sheet_name=0)
            df.to_csv(output_path, index=False, encoding=encoding)
            print(f"[OK] {input_path} → {output_path}  ({len(df)} rows × {len(df.columns)} cols)")
            exported.append((output_path, len(df), len(df.columns)))
        else:
            base = os.path.splitext(output_path)[0]
            for sheet in sheet_names:
                df = pd.read_excel(input_path, sheet_name=sheet)
                sheet_out = f"{base}_{sheet}.csv"
                df.to_csv(sheet_out, index=False, encoding=encoding)
                print(f"[OK] sheet '{sheet}' → {sheet_out}  ({len(df)} rows × {len(df.columns)} cols)")
                exported.append((sheet_out, len(df), len(df.columns)))

        return exported

    except Exception as e:
        print(f"[ERROR] Failed to convert {input_path}: {e}")
        return []


def convert_folder(folder_path, output_dir=None, prefix="", encoding="utf-8-sig"):
    """
    Convert all .xlsx / .xls files inside folder_path to CSV.

    Args:
        folder_path : source folder containing Excel files
        output_dir  : destination folder (defaults to folder_path)
        prefix      : optional filename prefix for every output CSV
        encoding    : csv encoding, default utf-8-sig (Excel-safe)

    Returns:
        (exported_list, success_count, total_count)
        exported_list is a list of (output_path, rows, cols) tuples.
    """
    xlsx_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    xlsx_files += glob.glob(os.path.join(folder_path, "*.xls"))

    if not xlsx_files:
        print(f"[INFO] No Excel files found in: {folder_path}")
        return [], 0, 0

    if output_dir is None:
        output_dir = folder_path

    os.makedirs(output_dir, exist_ok=True)

    print(f"[INFO] Found {len(xlsx_files)} Excel file(s) in '{folder_path}'\n")
    success = 0
    all_exported = []

    for file in xlsx_files:
        base_name = os.path.splitext(os.path.basename(file))[0]
        output_path = os.path.join(output_dir, f"{prefix}{base_name}.csv")
        result = xlsx_to_csv(file, output_path, encoding=encoding)
        if result:
            success += 1
            all_exported.extend(result)

    print(f"\n[DONE] Converted {success}/{len(xlsx_files)} file(s) successfully.")
    return all_exported, success, len(xlsx_files)


# ─── CLI Entry Point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    args = sys.argv[1:]

    if len(args) == 0:
        convert_folder(".")
    elif len(args) == 1:
        target = args[0]
        if os.path.isdir(target):
            convert_folder(target)
        else:
            xlsx_to_csv(target)
    elif len(args) == 2:
        xlsx_to_csv(args[0], args[1])
    else:
        print("Usage:")
        print("  python xlsx_to_csv.py                         # convert all in current folder")
        print("  python xlsx_to_csv.py file.xlsx               # convert single file")
        print("  python xlsx_to_csv.py file.xlsx output.csv    # custom output name")
        print("  python xlsx_to_csv.py ./folder                # convert all in folder")