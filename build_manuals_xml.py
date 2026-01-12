import os
import re
import time
import hashlib
import requests
from datetime import datetime, timezone
from urllib.parse import urljoin
from xml.sax.saxutils import escape

import pdfplumber  # pip install pdfplumber

BASE_URL = "https://www.dana.com"
INPUT_HTML_PATH = r"C:\Users\npettersson\Downloads\pdf links dana.txt"
   # <-- your saved page source file
OUT_DIR = r"manuals_pdfs"
OUT_XML = r"manuals_text.xml"

# Be polite to the server
SLEEP_SECONDS = 0.6
TIMEOUT = 60
MAX_RETRIES = 3

PDF_HREF_REGEX = re.compile(r'href="([^"]+\.pdf)"', re.IGNORECASE)

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def safe_filename(url: str) -> str:
    # Use the last part of the URL, but keep it filesystem-safe
    name = url.split("/")[-1]
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return name

def download_pdf(session: requests.Session, url: str, out_path: str) -> None:
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            with open(out_path, "wb") as f:
                f.write(r.content)
            return
        except Exception as e:
            if attempt == MAX_RETRIES:
                raise
            time.sleep(1.5 * attempt)

def extract_text_from_pdf(path: str) -> str:
    texts = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            if t.strip():
                texts.append(t)
    return "\n\n".join(texts)

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    with open(INPUT_HTML_PATH, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    rel_links = PDF_HREF_REGEX.findall(html)
    rel_links = list(dict.fromkeys(rel_links))  # de-dupe, keep order

    abs_links = [urljoin(BASE_URL, rel) for rel in rel_links]

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (manual-indexer; +https://www.dana.com)"
    })

    # Build XML
    now = datetime.now(timezone.utc).isoformat()

    with open(OUT_XML, "w", encoding="utf-8") as x:
        x.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        x.write(f'<manuals source="{escape(BASE_URL)}" fetchedAt="{escape(now)}">\n')

        for i, url in enumerate(abs_links, start=1):
            fname = safe_filename(url)
            pdf_path = os.path.join(OUT_DIR, fname)

            # Download if missing
            if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 1024:
                print(f"[{i}/{len(abs_links)}] Downloading: {url}")
                download_pdf(session, url, pdf_path)
                time.sleep(SLEEP_SECONDS)
            else:
                print(f"[{i}/{len(abs_links)}] Exists: {fname}")

            # Extract
            print(f"  Extracting text: {fname}")
            try:
                text = extract_text_from_pdf(pdf_path)
            except Exception as e:
                text = f"[EXTRACTION_FAILED] {e}"

            doc_hash = sha256_file(pdf_path)
            size_bytes = os.path.getsize(pdf_path)

            # Wrap big text in CDATA so you don’t have to escape everything
            x.write(f'  <manual id="{i}" url="{escape(url)}" file="{escape(fname)}" '
                    f'sha256="{doc_hash}" sizeBytes="{size_bytes}">\n')
            x.write("    <text><![CDATA[\n")
            x.write(text)
            x.write("\n    ]]></text>\n")
            x.write("  </manual>\n")

        x.write("</manuals>\n")

    print(f"\nDONE -> {OUT_XML}\nPDFs -> {OUT_DIR}\nLinks found -> {len(abs_links)}")

if __name__ == "__main__":
    main()
