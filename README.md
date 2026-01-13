# Parts Intelligence Engine

The Parts Intelligence Engine converts unstructured OEM parts manuals (PDFs, XMLs, and sidecar files) into a structured, queryable database. It automatically extracts part numbers, descriptions, model and serial applicability, and source locations, then stores everything in a local SQLite database that can be searched, exported, and extended with distributor cross-references.

---

## What it does

This project ingests large collections of OEM parts manuals and builds a machine-readable **“parts brain.”**  
Instead of hunting through thousands of PDFs, you get a normalized database that knows which parts belong to which machines, where they appear in the manuals, and how they relate to specific models and serial ranges.

---

## How to run

```bash
pip install pdfplumber openpyxl tqdm
python classify_manuals.py
python build_parts_brain.py
