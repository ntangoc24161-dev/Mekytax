import json
import os
import re
import sys
from pathlib import Path

# Add scripts dir to path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from ioffice_client import IOfficeClient
from search_police_docs import extract_current_table_rows

def main():
    client = IOfficeClient(headless=True)
    with client:
        client.ensure_logged_in()
        client.navigate_to_tra_cuu()

        queries = [
            {"name": "VPCQCSĐT", "so_ky_hieu": "VPCQCSĐT", "start_date": "", "end_date": ""},
            {"name": "Số KH chứa CSĐT", "so_ky_hieu": "CSĐT", "start_date": "", "end_date": ""},
            {"name": "Số KH chứa CSKT", "so_ky_hieu": "CSKT", "start_date": "", "end_date": ""},
            {"name": "Số KH chứa PC03", "so_ky_hieu": "PC03", "start_date": "", "end_date": ""},
            {"name": "Số KH chứa YC", "so_ky_hieu": "YC", "start_date": "", "end_date": ""},
            {"name": "Văn bản đến từ Thuế TP (DON) trong tuần", "so_ky_hieu": "DON", "start_date": "27/09/2026", "end_date": "02/10/2026"},
        ]

        all_matches = {}
        for q in queries:
            print(f"=== Query: {q['name']} ===")
            client.page.evaluate('''() => {
                const adv = document.querySelector('#div_timkiem_nangcao');
                if (adv) adv.style.display = 'block';
                if (typeof resetInput === 'function') resetInput();
            }''')
            client.page.wait_for_timeout(500)
            client.page.evaluate('''(q) => {
                if (q.so_ky_hieu) {
                    const el = document.querySelector('#txt_so_kyhieu');
                    if (el) el.value = q.so_ky_hieu;
                }
                if (q.start_date) {
                    const s = document.querySelector('#txt_start_date_ngayden');
                    if (s) s.value = q.start_date;
                }
                if (q.end_date) {
                    const e = document.querySelector('#txt_end_date_ngayden');
                    if (e) e.value = q.end_date;
                }
                const sel = document.querySelector('#selectPageRec');
                if (sel) sel.value = '50';
            }''', q)
            client.page.click("#btnClickLoading")
            client.page.wait_for_timeout(3500)
            rows = extract_current_table_rows(client.page)
            print(f"Results: {len(rows)}")
            for r in rows:
                key = (r['so_ky_hieu'], r['so_den'], r['doc_id'])
                if key not in all_matches:
                    all_matches[key] = r
                    print(f"  Số KH: {r['so_ky_hieu']} | Số đến: {r['so_den']} | Ngày đến: {r['ngay_den']} | CQ: {r['don_vi_ban_hanh']} | doc_id: {r['doc_id']}")
                    print(f"  Trích yếu: {r['trich_yeu']}")

        print(f"\nTotal unique dispatches found: {len(all_matches)}")
        
        # Save to file
        with open(Path(__file__).resolve().parent / "deep_check_results.json", "w", encoding="utf-8") as f:
            json.dump(list(all_matches.values()), f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
