import csv
from pathlib import Path
from datetime import datetime
import hashlib

# Root where ALL your Excel/CSV files live
ROOT = r"C:\Users\npettersson\Desktop\_PROJECTS\New Inventory Project"
# Index file will be created here
OUTPUT = r"C:\Users\npettersson\Desktop\_PROJECTS\New Inventory Project\_FileIndex.csv"

# Only Excel/CSV for now
EXT_WHITELIST = {".xlsx", ".xls", ".csv"}


def make_file_id(path: Path) -> str:
    """Create a stable id for each file based on its full path."""
    h = hashlib.md5(str(path).encode("utf-8")).hexdigest()
    return h[:12]


def guess_vendor(path: Path) -> str:
    """
    Guess vendor from folder names.
    Extend this as you add more brands / patterns.
    """
    p = str(path).lower()
    if "genie" in p:
        return "GENIE"
    if "jlg" in p:
        return "JLG"
    if "skyjack" in p:
        return "SKYJACK"
    if "snorkel" in p:
        return "SNORKEL"
    if "haulotte" in p:
        return "HAULOTTE"
    if "magni" in p:
        return "MAGNI"
    return "UNKNOWN"


def build_index(root: Path):
    rows = []

    for path in root.rglob("*"):
        if not path.is_file():
            continue

        ext = path.suffix.lower()
        if ext not in EXT_WHITELIST:
            continue

        # Skip temp Excel files like ~$
        if path.name.startswith("~$"):
            continue

        stat = path.stat()
        modified = datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds")

        file_id = make_file_id(path)
        vendor = guess_vendor(path)

        rows.append({
            "file_id": file_id,
            "full_path": str(path),
            "file_name": path.name,
            "extension": ext,
            "vendor": vendor,
            "size_bytes": stat.st_size,
            "last_modified": modified,
            "ingested": 0,  # 0 = not ingested yet, 1 = ingested
        })

    return rows


def main():
    root = Path(ROOT)
    if not root.exists():
        raise SystemExit(f"Root folder does not exist: {root}")

    rows = build_index(root)
    if not rows:
        raise SystemExit("No Excel/CSV files found under root.")

    fieldnames = [
        "file_id",
        "full_path",
        "file_name",
        "extension",
        "vendor",
        "size_bytes",
        "last_modified",
        "ingested",
    ]

    with open(OUTPUT, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {OUTPUT}")


if __name__ == "__main__":
    main()
