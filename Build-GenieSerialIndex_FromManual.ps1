# Build-GenieSerialIndex_FromManual.ps1

Write-Host "=== Building Genie Serial Index (from manual text) ==="

# Path to Python
$pythonExe = "C:\Users\npettersson\AppData\Local\Programs\Python\Python312\python.exe"

if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERROR] Python not found at $pythonExe"
    Read-Host "Press Enter to exit"
    exit 1
}

# Root folder of Genie manuals
$rootDir = "C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE"

# Temp Python script path
$pyPath = Join-Path $env:TEMP "build_genie_serial_index_from_manual.py"

# Python code block
$pyCode = @"
import os
import csv
import re
import pdfplumber

ROOT_DIR = r"C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE"
OUT_CSV = os.path.join(ROOT_DIR, "_SerialIndex_FromManual.csv")

def parse_folder_name(folder_name):
    # Example: '001 - 32223GT - Genie S-40, S-45'
    m = re.match(r'^\s*\d+\s*-\s*([^-]+?)\s*-\s*(.+)$', folder_name)
    if not m:
        return None, None
    manual_no = m.group(1).strip()
    model_title = m.group(2).strip()
    return manual_no, model_title

def collect_serial_lines(pdf_path, max_pages=5):
    rows = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for idx, page in enumerate(pdf.pages[:max_pages], start=1):
                text = page.extract_text() or ""
                for line in text.splitlines():
                    if "serial" in line.lower():
                        rows.append({
                            "page": idx,
                            "serial_line": line.strip()
                        })
    except Exception as e:
        print(f"[WARN] Could not read {pdf_path}: {e}")
    return rows

def build_index(root_dir, out_csv):
    out_rows = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        folder_name = os.path.basename(dirpath)
        manual_no, model_title = parse_folder_name(folder_name)
        if not manual_no:
            continue

        for fname in filenames:
            if not fname.lower().endswith(".pdf"):
                continue

            pdf_path = os.path.join(dirpath, fname)
            serial_lines = collect_serial_lines(pdf_path)

            if not serial_lines:
                # still capture a record that we scanned this manual
                out_rows.append({
                    "manual_no": manual_no,
                    "model_title": model_title,
                    "pdf_path": pdf_path,
                    "page": "",
                    "serial_line": ""
                })
            else:
                for s in serial_lines:
                    out_rows.append({
                        "manual_no": manual_no,
                        "model_title": model_title,
                        "pdf_path": pdf_path,
                        "page": s["page"],
                        "serial_line": s["serial_line"]
                    })

    fieldnames = ["manual_no", "model_title", "pdf_path", "page", "serial_line"]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)

    print(f"[INFO] Wrote {len(out_rows)} rows to {out_csv}")

if __name__ == "__main__":
    build_index(ROOT_DIR, OUT_CSV)
"@

# Write the Python script
Set-Content -Path $pyPath -Encoding UTF8 -Value $pyCode

Write-Host "[INFO] Running Python serial index builder (from manual)..."
& $pythonExe $pyPath

Write-Host ""
Write-Host "Done. Check _SerialIndex_FromManual.csv under:"
Write-Host "  $rootDir"
Read-Host "Press Enter to exit"
