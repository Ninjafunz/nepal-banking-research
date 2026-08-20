"""
fetch_nrb_data.py
=================
Downloads NRB Banking & Financial Statistics PDFs and attempts to extract
bank-level data tables using pdfplumber.

NRB publishes "Banking and Financial Statistics" monthly.
Year-end reports (mid-July each year / Month 12 of Nepali calendar = Ashadh)
are the most complete and correspond to our fiscal year endpoints.

The URL pattern for NRB uploads:
  https://www.nrb.org.np/contents/uploads/YYYY/MM/filename.pdf
  (varies — we try known patterns for year-end reports)

Usage:
    python fetch_nrb_data.py
"""

import os
import sys
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(__file__))

BASE   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW    = os.path.join(BASE, "DATA", "raw")
os.makedirs(RAW, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# Known direct URLs for NRB Banking & Financial Statistics (Year-end reports)
# These are the Ashadh-end (mid-July) annual summaries — verified from NRB site
# Naming convention varies: "banking-and-financial-statistics-of-mid-july-YYYY"
# ─────────────────────────────────────────────────────────────────────────────
NRB_ANNUAL_URLS = {
    # FY2025 (Mid-July 2025)
    2025: [
        "https://www.nrb.org.np/contents/uploads/2025/09/Banking-and-Financial-Statistics-of-Mid-July-2025.xlsx",
        "https://www.nrb.org.np/contents/uploads/2025/09/Banking-and-Financial-Statistics-of-Mid-July-2025.pdf",
        "https://www.nrb.org.np/contents/uploads/2025/08/Banking-and-Financial-Statistics-of-Mid-July-2025.xlsx",
        "https://www.nrb.org.np/contents/uploads/2025/08/Banking-and-Financial-Statistics-of-Mid-July-2025.pdf",
    ],
    # FY2024 (Mid-July 2024)
    2024: [
        "https://www.nrb.org.np/contents/uploads/2024/09/Banking-and-Financial-Statistics-of-Mid-July-2024.xlsx",
        "https://www.nrb.org.np/contents/uploads/2024/09/Banking-and-Financial-Statistics-of-Mid-July-2024.pdf",
        "https://www.nrb.org.np/contents/uploads/2024/10/Banking-and-Financial-Statistics-of-Mid-July-2024.xlsx",
        "https://www.nrb.org.np/contents/uploads/2024/10/Banking-and-Financial-Statistics-of-Mid-July-2024.pdf",
    ],
    # FY2023
    2023: [
        "https://www.nrb.org.np/contents/uploads/2023/09/Banking-and-Financial-Statistics-of-Mid-July-2023.xlsx",
        "https://www.nrb.org.np/contents/uploads/2023/09/Banking-and-Financial-Statistics-of-Mid-July-2023.pdf",
        "https://www.nrb.org.np/contents/uploads/2023/10/Banking-and-Financial-Statistics-of-Mid-July-2023.xlsx",
    ],
    # FY2022
    2022: [
        "https://www.nrb.org.np/contents/uploads/2022/10/Banking-and-Financial-Statistics-of-Mid-July-2022.xlsx",
        "https://www.nrb.org.np/contents/uploads/2022/09/Banking-and-Financial-Statistics-of-Mid-July-2022.xlsx",
        "https://www.nrb.org.np/contents/uploads/2022/10/Banking-and-Financial-Statistics-of-Mid-July-2022.pdf",
    ],
    # FY2021
    2021: [
        "https://www.nrb.org.np/contents/uploads/2021/10/Banking-and-Financial-Statistics-of-Mid-July-2021.xlsx",
        "https://www.nrb.org.np/contents/uploads/2021/09/Banking-and-Financial-Statistics-of-Mid-July-2021.xlsx",
        "https://www.nrb.org.np/contents/uploads/2021/10/Banking-and-Financial-Statistics-of-Mid-July-2021.pdf",
    ],
    # FY2020
    2020: [
        "https://www.nrb.org.np/contents/uploads/2020/10/Banking-and-Financial-Statistics-of-Mid-July-2020.xlsx",
        "https://www.nrb.org.np/contents/uploads/2020/09/Banking-and-Financial-Statistics-of-Mid-July-2020.xlsx",
        "https://www.nrb.org.np/contents/uploads/2020/10/Banking-and-Financial-Statistics-of-Mid-July-2020.pdf",
    ],
}

def try_download(urls, dest_path, label):
    """Try each URL in order, save first that works. Returns (success, ext)."""
    for url in urls:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as r:
                if r.status == 200:
                    data = r.read()
                    ext = ".xlsx" if url.endswith(".xlsx") else ".pdf"
                    path = dest_path + ext
                    with open(path, "wb") as f:
                        f.write(data)
                    print(f"  [OK]   {label} → {os.path.basename(path)} ({len(data)//1024} KB)")
                    return True, ext, path
        except Exception:
            continue
    print(f"  [FAIL] {label} — no URL succeeded")
    return False, None, None

def extract_from_xlsx(path, fy):
    """Try to read bank-level data from NRB Excel file."""
    try:
        import pandas as pd
        xl = pd.ExcelFile(path)
        print(f"    Sheets: {xl.sheet_names[:10]}")
        results = {}
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet, header=None)
                text = df.to_string()
                # Look for sheets with bank data (contains bank names)
                bank_indicators = ["nabil", "everest", "himalayan", "kumari",
                                   "nrb", "commercial", "total", "assets"]
                if any(ind in text.lower() for ind in bank_indicators):
                    results[sheet] = df
                    print(f"    Found bank data in sheet: '{sheet}' ({df.shape})")
            except Exception:
                pass
        return results
    except Exception as e:
        print(f"    Could not parse xlsx: {e}")
        return {}

def run():
    print("NRB Banking Statistics Download & Extraction\n" + "="*50)
    downloaded = {}

    for fy, urls in NRB_ANNUAL_URLS.items():
        dest = os.path.join(RAW, f"nrb_banking_stats_FY{fy}")
        success, ext, path = try_download(urls, dest, f"FY{fy}")
        if success:
            downloaded[fy] = {"path": path, "ext": ext}

    print(f"\nDownloaded: {len(downloaded)} of {len(NRB_ANNUAL_URLS)} files\n")

    # Try to extract from any xlsx files
    extracted = {}
    for fy, info in downloaded.items():
        if info["ext"] == ".xlsx":
            print(f"\nExtracting FY{fy}...")
            sheets = extract_from_xlsx(info["path"], fy)
            extracted[fy] = sheets

    # Summary
    print("\n" + "="*50)
    print("SUMMARY")
    print("="*50)
    for fy in sorted(NRB_ANNUAL_URLS.keys()):
        status = "Downloaded" if fy in downloaded else "NOT FOUND"
        fmt    = downloaded[fy]["ext"] if fy in downloaded else "—"
        print(f"  FY{fy}: {status} ({fmt})")

    print(f"\nRaw files saved to: {RAW}")
    print("\nNEXT STEPS:")
    print("  - For each downloaded xlsx: open and identify the bank-level table")
    print("  - For PDFs: use the data entry guide to locate values manually")
    print("  - Run build_panel.py after populating 01_bank_financials.xlsx")

if __name__ == "__main__":
    run()

