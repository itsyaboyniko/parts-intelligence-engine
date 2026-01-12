import os
import csv

# Root folder of Genie manuals
ROOT_DIR = r"C:\Users\npettersson\Desktop\_MANUALS\MANUALS\GENIE"
OUT_CSV = os.path.join(ROOT_DIR, "_SerialIndex_Master.csv")


def parse_sidecar_txt(txt_path):
    """
    Parse a sidecar txt file like:
      Manual: 32223GT - Genie S-40, S-45
      Token : 32223GT
      Link  : ...
      Serial: SN: Prior to 0830
      File  : C:\...\32223GT - Genie S-40, S-45.pdf
    """
    manual_full = ""
    token = ""
    serial_text = ""
    pdf_path = ""

    try:
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                lower = line.lower()

                if lower.startswith("manual"):
                    # Manual: 32223GT - Genie S-40, S-45
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        manual_full = parts[1].strip()

                elif lower.startswith("token"):
                    # Token : 32223GT
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        token = parts[1].strip()

                elif lower.startswith("serial"):
                    # Serial: SN: Prior to 0830
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        serial_text = parts[1].strip()

                elif lower.startswith("file"):
                    # File  : C:\...\32223GT - Genie S-40, S-45.pdf
                    parts = line.split(":", 1)
                    if len(parts) == 2:
                        pdf_path = parts[1].strip()

    except Exception as e:
        print(f"[WARN] Failed to read {txt_path}: {e}")

    # Derive model_title from manual_full if possible
    model_title = ""
    manual_no = token

    # Example manual_full: "32223GT - Genie S-40, S-45"
    if manual_full:
        if " - " in manual_full:
            first, rest = manual_full.split(" - ", 1)
            if not manual_no:
                manual_no = first.strip()
            model_title = rest.strip()
        else:
            model_title = manual_full.strip()

    return {
        "manual_no": manual_no,
        "model_title": model_title,
        "serial_text": serial_text,
        "pdf_path": pdf_path,
        "sidecar_path": txt_path,
    }


def build_index(root_dir, out_csv):
    # Force overwrite by deleting old file first if it exists
    try:
        if os.path.exists(out_csv):
            os.remove(out_csv)
    except Exception as e:
        print(f"[ERROR] Could not delete old index file: {e}")
        return

    rows = []

    for dirpath, dirnames, filenames in os.walk(root_dir):
        for fname in filenames:
            # Look for "SN - 32223GT.txt" style files
            lower = fname.lower()
            if not lower.startswith("sn -") or not lower.endswith(".txt"):
                continue

            txt_path = os.path.join(dirpath, fname)
            info = parse_sidecar_txt(txt_path)

            # If manual_no still empty, try filename: "SN - 32223GT.txt"
            if not info["manual_no"]:
                base = os.path.splitext(fname)[0]  # "SN - 32223GT"
                parts = base.split("-", 1)
                if len(parts) == 2:
                    info["manual_no"] = parts[1].strip()

            # If model_title empty, try folder name: "001 - 32223GT - Genie S-40, S-45"
            if not info["model_title"]:
                folder_name = os.path.basename(dirpath)
                chunks = folder_name.split(" - ", 2)
                if len(chunks) == 3:
                    # chunks[1] = manual no, chunks[2] = model title
                    if not info["manual_no"]:
                        info["manual_no"] = chunks[1].strip()
                    info["model_title"] = chunks[2].strip()

            rows.append(info)

    fieldnames = ["manual_no", "model_title", "serial_text", "pdf_path", "sidecar_path"]
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"[INFO] Wrote {len(rows)} rows to {out_csv}")


if __name__ == "__main__":
    build_index(ROOT_DIR, OUT_CSV)
