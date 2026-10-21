# -*- coding: utf-8 -*-
"""
fetch_october_batch.py - Tra cuu va tai ve cac cong van canh bao tu 27/09/2026 den 07/10/2026
Luu vao thu muc: D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026\2026-10
"""
import os
import sys
import json
from pathlib import Path

IOFFICE_SCRIPTS = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\tra-cuu-ioffice-dong-nai\scripts")
sys.path.insert(0, str(IOFFICE_SCRIPTS))

from ioffice_client import IOfficeClient

OUT_DIR = Path(r"D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026\2026-10")
OUT_DIR.mkdir(parents=True, exist_ok=True)

RESULTS_JSON = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts\search_results_2026_10.json")
PROGRESS_JSON = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts\download_progress_2026_10.json")

def main():
    client = IOfficeClient(headless=True)
    with client:
        client.ensure_logged_in()
        print("Logged in successfully to VNPT iOffice!")
        
        # Search documents
        results = client.search(
            trich_yeu="cảnh báo",
            from_date="27/09/2026",
            to_date="07/10/2026",
            max_results=100
        )
        print(f"\nFound {len(results)} search results.")
        with open(RESULTS_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        # Load progress if any
        progress = {}
        if PROGRESS_JSON.exists():
            try:
                with open(PROGRESS_JSON, "r", encoding="utf-8") as f:
                    progress = json.load(f)
            except Exception:
                progress = {}
        
        download_summary = []
        for idx, doc in enumerate(results, 1):
            doc_id = doc.get("doc_id")
            so_ky_hieu = doc.get("so_ky_hieu", f"doc_{idx}")
            so_den = doc.get("so_den", f"{idx}")
            trich_yeu = doc.get("trich_yeu", "")
            
            print(f"\n[{idx}/{len(results)}] {so_ky_hieu} (So den: {so_den}) - doc_id: {doc_id}")
            print(f"    Trich yeu: {trich_yeu[:80]}")
            
            if not doc_id:
                print("    -> Khong co doc_id, bo qua.")
                continue
                
            if str(doc_id) in progress:
                print(f"    -> Da tai truoc do: {len(progress[str(doc_id)].get('downloaded_files', []))} tep.")
                doc["downloaded_files"] = progress[str(doc_id)].get("downloaded_files", [])
                download_summary.append(doc)
                continue
                
            prefix = f"[{so_den}]_{so_ky_hieu}"
            try:
                files = client.download_document(doc_id, OUT_DIR, filename_prefix=prefix)
                doc["downloaded_files"] = [str(f) for f in files]
                download_summary.append(doc)
                progress[str(doc_id)] = doc
                
                with open(PROGRESS_JSON, "w", encoding="utf-8") as f:
                    json.dump(progress, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"    -> LOI khi tai doc_id={doc_id}: {e}")
                
        print(f"\nHOAN TAT: Da xu ly {len(download_summary)}/{len(results)} van ban ve {OUT_DIR}.")

if __name__ == "__main__":
    main()
