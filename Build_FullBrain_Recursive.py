import pandas as pd
from pathlib import Path

# ====================================================================================
# Base folder to recursively scan
# ====================================================================================
BASE_DIR = Path(
    r"C:\Users\npettersson\Desktop\_PROJECTS\New Inventory Project"
)

# Output file
OUT_FILE = BASE_DIR / "MASTER_FULL_BRAIN.xlsx"


def get_all_excel_files(base_path: Path):
    """
    Recursively find all .xlsx files in the entire directory tree.
    Excludes:
        - The output master XLSX file
        - Temporary Excel lock files
        - Anything with '~$' in the name
    """
    excel_files = []
    for path in base_path.rglob("*.xlsx"):
        name = path.name.lower()

        if "~$" in name:  # temporary Excel lock files
            continue
        if "master_full_brain" in name:  # do not include the output file
            continue

        excel_files.append(path)

    return excel_files


def process_excel_file(path: Path):
    """Extract first 5 columns as standardized fields from all sheets."""
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        print(f"[ERROR] Cannot open {path.name}: {e}")
        return []

    frames = []

    for sheet in xls.sheet_names:
        try:
            df = xls.parse(sheet_name=sheet, header=None, dtype=str)
        except Exception as e:
            print(f"[WARN] Failed reading sheet '{sheet}' in {path.name}: {e}")
            continue

        # Need at least 5 usable columns
        if df.empty or df.shape[1] < 5:
            continue

        sub = df.iloc[:, :5].copy()
        sub.columns = ["ModelAndBuild", "Description", "Price", "PartNumber", "Year"]

        # Drop empty rows
        sub = sub.dropna(how="all")
        if sub.empty:
            continue

        # Add file + sheet metadata
        sub.insert(0, "SourceFile", path.name)
        sub.insert(1, "SheetName", sheet)

        frames.append(sub)

    return frames


def main():
    print(f"\n[INFO] Scanning folder recursively:\n{BASE_DIR}\n")

    # Find all Excel files
    excel_files = get_all_excel_files(BASE_DIR)

    if not excel_files:
        print("[ERROR] No Excel files found.")
        return

    print("[INFO] Excel files found:")
    for f in excel_files:
        print(f"  - {f}")

    print("\n[INFO] Extracting data...")

    frames = []

    for f in excel_files:
        print(f"[INFO] Processing {f.name} ...")
        extracted_frames = process_excel_file(f)
        frames.extend(extracted_frames)

    if not frames:
        print("[ERROR] No data extracted from any Excel file.")
        return

    combined = pd.concat(frames, ignore_index=True)
    print(f"\n[INFO] Rows before dedupe: {len(combined)}")

    # Remove duplicate rows
    combined = combined.drop_duplicates()
    print(f"[INFO] Rows after dedupe:  {len(combined)}")

    # Clean up stray "Field4_text" artifacts
    if "Year" in combined.columns:
        combined = combined[combined["Year"] != "Field4_text"]

    # Save master file
    OUT_FILE.unlink(missing_ok=True)
    combined.to_excel(OUT_FILE, index=False)

    print(f"\n[OK] Master brain created:\n{OUT_FILE}\n")


if __name__ == "__main__":
    main()
