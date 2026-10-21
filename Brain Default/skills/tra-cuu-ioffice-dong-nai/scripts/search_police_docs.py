import json
import os
import re
import sys
import time
from pathlib import Path

# Add scripts dir to path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from ioffice_client import IOfficeClient

def extract_current_table_rows(page):
    return page.evaluate('''() => {
        const t = document.querySelectorAll('table')[7];
        if (!t) return [];
        const res = [];
        t.querySelectorAll('tr').forEach((r, idx) => {
            if (idx === 0) return;
            const cells = Array.from(r.querySelectorAll('td')).map(td => td.innerText.trim().replace(/\\s+/g, ' '));
            if (cells.length < 5) return;

            let doc_id = null;
            r.querySelectorAll('a').forEach(a => {
                const h = (a.getAttribute('href') || '') + (a.getAttribute('onclick') || '');
                const m = h.match(/allFileDownload\\((\\d+)\\)/);
                if (m && !doc_id) doc_id = m[1];
            });

            res.push({
                stt: cells[0] || '',
                so_eoffice: cells[1] || '',
                trich_yeu: cells[2] || '',
                so_den: cells[3] || '',
                so_ky_hieu: cells[4] || '',
                phuc_dap: cells[5] || '',
                ky_so: cells[6] || '',
                ngay_vb: cells[7] || '',
                trang_thai: cells[8] || '',
                ngay_den: cells[9] || '',
                do_khan: cells[10] || '',
                don_vi_ban_hanh: cells[11] || '',
                xlc: cells[12] || '',
                doc_id: doc_id
            });
        });
        return res;
    }''')

def main():
    client = IOfficeClient(headless=True)
    with client:
        client.ensure_logged_in()
        client.navigate_to_tra_cuu()

        all_collected_docs = {}

        # List of search queries to run
        search_configs = [
            {"desc": "Tra cứu tất cả văn bản đến 27/09 - 02/10", "trichyeu": "", "start_date": "27/09/2026", "end_date": "02/10/2026"},
            {"desc": "Tra cứu từ khóa 'công an'", "trichyeu": "công an", "start_date": "", "end_date": ""},
            {"desc": "Tra cứu từ khóa 'CSĐT'", "trichyeu": "CSĐT", "start_date": "", "end_date": ""},
            {"desc": "Tra cứu từ khóa 'PC03'", "trichyeu": "PC03", "start_date": "", "end_date": ""},
            {"desc": "Tra cứu từ khóa 'xác minh'", "trichyeu": "xác minh", "start_date": "27/09/2026", "end_date": "02/10/2026"},
            {"desc": "Tra cứu từ khóa 'cung cấp'", "trichyeu": "cung cấp", "start_date": "27/09/2026", "end_date": "02/10/2026"},
        ]

        for cfg in search_configs:
            print(f"\n--- Đang thực hiện: {cfg['desc']} ---")
            client.page.evaluate('''() => {
                const adv = document.querySelector('#div_timkiem_nangcao');
                if (adv) adv.style.display = 'block';
                if (typeof resetInput === 'function') resetInput();
            }''')
            client.page.wait_for_timeout(1000)

            # Set inputs
            client.page.evaluate('''(cfg) => {
                if (cfg.trichyeu) {
                    const ty = document.querySelector('#txt_trichyeu');
                    if (ty) ty.value = cfg.trichyeu;
                }
                if (cfg.start_date) {
                    const s1 = document.querySelector('#txt_start_date_ngayden');
                    if (s1) s1.value = cfg.start_date;
                }
                if (cfg.end_date) {
                    const e1 = document.querySelector('#txt_end_date_ngayden');
                    if (e1) e1.value = cfg.end_date;
                }
                const sel = document.querySelector('#selectPageRec');
                if (sel) sel.value = '50';
            }''', cfg)
            
            client.page.wait_for_timeout(500)
            client.page.click("#btnClickLoading")
            client.page.wait_for_timeout(4000)

            rows = extract_current_table_rows(client.page)
            print(f"-> Thu được {len(rows)} kết quả.")
            for r in rows:
                key = (r['so_ky_hieu'], r['so_den'], r['doc_id'])
                if key not in all_collected_docs:
                    all_collected_docs[key] = r

        print(f"\n=== TỔNG HỢP TOÀN BỘ VĂN BẢN THU THẬP ĐƯỢC: {len(all_collected_docs)} VĂN BẢN ===")
        output_file = Path(__file__).resolve().parent / "all_collected_docs.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(list(all_collected_docs.values()), f, ensure_ascii=False, indent=2)

        # Filter strictly for Police / Investigation / Dossier Provision dispatches
        police_keywords = [
            "công an", "ca ", "catp", "csđt", "cskt", "cshs", "pc03", "pc01", "pc02", "pa04",
            "điều tra", "cung cấp hồ sơ", "cung cấp thông tin", "xác minh", "phối hợp cung cấp",
            "viện kiểm sát", "vks", "tòa án", "cơ quan cảnh sát điều tra"
        ]

        police_docs = []
        for doc in all_collected_docs.values():
            full_text = f"{doc['trich_yeu']} {doc['so_ky_hieu']} {doc['don_vi_ban_hanh']}".lower()
            if any(k in full_text for k in police_keywords):
                # Exclude internal non-police routine admin notices if any
                police_docs.append(doc)

        print(f"\n=== KẾT QUẢ LỌC: {len(police_docs)} CÔNG VĂN CÔNG AN / CUNG CẤP HỒ SƠ ===")
        for idx, d in enumerate(police_docs, 1):
            print(f"[{idx}] Số KH: {d['so_ky_hieu']} | Số đến: {d['so_den']} | Ngày đến: {d['ngay_den']} | Ngày VB: {d['ngay_vb']}")
            print(f"    Cơ quan ban hành: {d['don_vi_ban_hanh']}")
            print(f"    Trích yếu: {d['trich_yeu']}")
            print(f"    doc_id: {d['doc_id']}")
            print("-" * 70)

        # Save police docs to JSON
        police_file = Path(__file__).resolve().parent / "police_docs_filtered.json"
        with open(police_file, "w", encoding="utf-8") as f:
            json.dump(police_docs, f, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()

