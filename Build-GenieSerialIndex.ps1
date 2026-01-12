# Build-GenieSerialIndex.ps1

Write-Host "=== Building Genie _SerialIndex_Master.csv ==="

# Path to Python
$pythonExe = "C:\Users\npettersson\AppData\Local\Programs\Python\Python312\python.exe"

# Root folder of Genie manuals
$rootDir = "C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE"

# Temp Python script path
$pyPath = Join-Path $env:TEMP "build_genie_serial_index.py"

# Python code block
$pyCode = @"
import os, csv, re
import pdfplumber

ROOT_DIR = r"C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE"
OUT_CSV = os.path.join(ROOT_DIR, "_SerialIndex_Master.csv")

def parse_folder_name(folder_name):
    m = re.match(r'^\s*\d+\s*-\s*([^-]+?)\s*-\s*(.+)$', folder_name)
    if not m:
        return None, None
    return m.group(1).strip(), m.group(2).strip()

def extract_serial_ranges_from_pdf(pdf_path, max_pages=3):
    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:max_pages]:
                text = page.extract_text() or ""
                for line in text.splitlines():
                    u = line.upper()
                    if "SERIAL" not in u:
                        continue
                    if not any(x in u for x in ["RANGE", "NO.", "NUMBER"]):
                        continue
                    m = re.search(r'([A-Z0-9\-]+)\s*(?:TO|[-])\s*([A-Z0-9\-]+)', u)
                    if m:
                        serial_from = m.group(1)
                        serial_to = m.group(2)
                    else:
                        serial_from = ""
                        serial_to = ""
                    results.append({
                        "serial_from": serial_from,
                        "serial_to": serial_to,
                        "raw_line": line.strip()
                    })
    except Exception as e:
        print(f"[WARN] Could not read {pdf_path}: {e}")
    return results

def build_serial_index(root_dir, out_csv):
    rows = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        folder_name = os.path.basename(dirpath)
        manual_no, model_title = parse_folder_name(folder_name)
        if not manual_no:
            continue
        for fname in filenames:
            if not fname.lower().endswith(".pdf"):
                continue
            pdf_path = os.path.join(dirpath, fname)
            serials = extract_serial_ranges_from_pdf(pdf_path)
            if not serials:
                rows.append({
                    "manual_no": manual_no,
                    "model_title": model_title,
                    "pdf_path": pdf_path,
                    "serial_from": "",
                    "serial_to": "",
                    "raw_line": ""
                })
            else:
                for s in serials:
                    rows.append({
                        "manual_no": manual_no,
                        "model_title": model_title,
                        "pdf_path": pdf_path,
                        "serial_from": s["serial_from"],
                        "serial_to": s["serial_to"],
                        "raw_line": s["raw_line"]
                    })

    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "manual_no", "model_title", "pdf_path",
            "serial_from", "serial_to", "raw_line"
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"[INFO] Saved: {out_csv}")

if __name__ == "__main__":
    build_serial_index(ROOT_DIR, OUT_CSV)
"@

# Write Python script
Set-Content -Path $pyPath -Encoding UTF8 -Value $pyCode

Write-Host "[INFO] Running Python builder..."
& $pythonExe $pyPath

Write-Host "Done. Press Enter to exit."
Read-Host
