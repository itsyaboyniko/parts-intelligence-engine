# Build-GenieSerialIndex_FromSidecars.ps1

Write-Host "=== Building Genie _SerialIndex_Master.csv from sidecar SN files ==="

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
$pyPath = Join-Path $env:TEMP "build_genie_serial_index_from_sidecars.py"

# Python code block (keep strings mostly single-quoted to avoid escaping issues)
$pyCode = @"
import os
import csv

ROOT_DIR = r'C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE'
OUT_CSV = os.path.join(ROOT_DIR, '_SerialIndex_Master.csv')

def parse_sidecar_txt(txt_path):
    manual_full = ''
    token = ''
    serial_text = ''
    pdf_path = ''

    try:
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                line = line.strip()
                lower = line.lower()

                if lower.startswith('manual'):
                    # Manual: 32223GT - Genie S-40, S-45
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        manual_full = parts[1].strip()

                elif lower.startswith('token'):
                    # Token : 32223GT
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        token = parts[1].strip()

                elif lower.startswith('serial'):
                    # Serial: SN: Prior to 0830
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        serial_text = parts[1].strip()

                elif lower.startswith('file'):
                    # File  : C:\...\32223GT - Genie S-40, S-45.pdf
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        pdf_path = parts[1].strip()
    except Exception as e:
        print(f'[WARN] Failed to read {txt_path}: {e}')

    # Derive model_title from manual_full if possible
    model_title = ''
    manual_no = token

    # Example manual_full: 32223GT - Genie S-40, S-45
    if manual_full:
        if ' - ' in manual_full:
            first, rest = manual_full.split(' - ', 1)
            if not manual_no:
                manual_no = first.strip()
            model_title = rest.strip()
        else:
            # If no dash, treat the whole thing as model title if token already known
            model_title = manual_full.strip()

    # Fallback: if no model_title, but we have txt folder name, caller can adjust
    return {
        'manual_no': manual_no,
        'model_title': model_title,
        'serial_text': serial_text,
        'pdf_path': pdf_path,
        'sidecar_path': txt_path,
    }

def build_index(root_dir, out_csv):
    rows = []

    for dirpath, dirnames, filenames in os.walk(root_dir):
        for fname in filenames:
            # Look for 'SN - 32223GT.txt' style files
            lower = fname.lower()
            if not lower.startswith('sn -') or not lower.endswith('.txt'):
                continue

            txt_path = os.path.join(dirpath, fname)
            info = parse_sidecar_txt(txt_path)

            # If we still don't have a manual_no, try to pull it from the filename
            if not info['manual_no']:
                # Example: 'SN - 32223GT.txt'
                base = os.path.splitext(fname)[0]  # 'SN - 32223GT'
                parts = base.split('-', 1)
                if len(parts) == 2:
                    info['manual_no'] = parts[1].strip()

            # If model_title is empty, try using folder name pattern
            if not info['model_title']:
                folder_name = os.path.basename(dirpath)
                # Example: '001 - 32223GT - Genie S-40, S-45'
                if ' - ' in folder_name:
                    chunks = folder_name.split(' - ', 2)
                    if len(chunks) == 3:
                        # chunks[1] should be manual no, chunks[2] model title
                        info['manual_no'] = info['manual_no'] or chunks[1].strip()
                        info['model_title'] = chunks[2].strip()

            rows.append(info)

    fieldnames = ['manual_no', 'model_title', 'serial_text', 'pdf_path', 'sidecar_path']
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f'[INFO] Wrote {len(rows)} rows to {out_csv}')

if __name__ == '__main__':
    build_index(ROOT_DIR, OUT_CSV)
"@

# Write the Python script
Set-Content -Path $pyPath -Encoding UTF8 -Value $pyCode

Write-Host "[INFO] Running Python serial index builder (from sidecar SN txt files)..."
& $pythonExe $pyPath

Write-Host ""
Write-Host "Done. Check _SerialIndex_Master.csv under:"
Write-Host "  $rootDir"
Read-Host "Press Enter to exit"
