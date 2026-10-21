import os
import re
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, Page, BrowserContext

DEFAULT_URL = "https://dongnai.vnptioffice.vn/qlvbdh_dni/main?lang=vi"
DEFAULT_SESSION_FILE = Path(__file__).resolve().parent.parent / "ioffice_session.json"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

class IOfficeClient:
    def __init__(self, username="ntangoc3.dni.don", password="Dni@1234", session_file=None, headless=True):
        self.username = username
        self.password = password
        self.session_file = Path(session_file) if session_file else DEFAULT_SESSION_FILE
        self.headless = headless
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def start(self):
        self.playwright = sync_playwright().start()
        launch_kwargs = {
            "headless": self.headless,
        }
        if os.path.exists(CHROME_PATH):
            launch_kwargs["executable_path"] = CHROME_PATH

        self.browser = self.playwright.chromium.launch(**launch_kwargs)

        context_kwargs = {
            "ignore_https_errors": True,
            "accept_downloads": True,
        }
        if self.session_file.exists():
            context_kwargs["storage_state"] = str(self.session_file)

        self.context = self.browser.new_context(**context_kwargs)
        self.page = self.context.new_page()

    def close(self):
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def is_logged_in(self) -> bool:
        """Kiểm tra xem session hiện tại có đang hoạt động hay không."""
        try:
            self.page.goto(DEFAULT_URL, wait_until="domcontentloaded", timeout=15000)
            self.page.wait_for_timeout(1000)
            title = self.page.title()
            if "Trang chủ" in title:
                return True
            # Kiểm tra xem có ô login không
            if self.page.query_selector("#userName"):
                return False
            return "Trang chủ" in self.page.inner_text("body")
        except Exception:
            return False

    def login(self) -> bool:
        """Thực hiện đăng nhập hệ thống iOffice và lưu session cookie."""
        print(f"Đang kết nối tới cổng VNPT iOffice: {DEFAULT_URL}...")
        self.page.goto(DEFAULT_URL, wait_until="domcontentloaded", timeout=25000)
        self.page.wait_for_timeout(1500)
        
        # Nếu đã ở Trang chủ
        if "Trang chủ" in self.page.title():
            print("Đã đăng nhập sẵn từ phiên trước!")
            self.context.storage_state(path=str(self.session_file))
            return True

        # Điền thông tin đăng nhập
        if self.page.query_selector("#userName"):
            print(f"Đang nhập tài khoản: {self.username}...")
            self.page.fill("#userName", self.username)
            self.page.fill("#passWord", self.password)
            self.page.click("#submitBtn")
            self.page.wait_for_timeout(3000)

        title = self.page.title()
        if "Trang chủ" in title or "qlvbdh" in self.page.url:
            print(f"Đăng nhập thành công! Tiêu đề trang: {title}")
            self.session_file.parent.mkdir(parents=True, exist_ok=True)
            self.context.storage_state(path=str(self.session_file))
            print(f"Đã lưu phiên làm việc (session state) tại: {self.session_file}")
            return True
        else:
            print(f"Đăng nhập thất bại hoặc cần xác thực thêm. Tiêu đề: {title}")
            return False

    def ensure_logged_in(self):
        """Đảm bảo đang ở trạng thái đăng nhập hợp lệ."""
        if not self.is_logged_in():
            print("Phiên làm việc hết hạn hoặc chưa đăng nhập. Tiến hành đăng nhập lại...")
            if not self.login():
                raise RuntimeError("Không thể đăng nhập vào VNPT iOffice.")

    def navigate_to_tra_cuu(self):
        """Chuyển đến trang Tra cứu văn bản."""
        self.ensure_logged_in()
        print("Đang điều hướng tới phân hệ 'Tra cứu văn bản'...")
        
        if self.page.query_selector("#btnClickLoading"):
            print("Đã ở trang Tra cứu văn bản.")
            return

        tra_cuu_url = "https://dongnai.vnptioffice.vn/qlvbdh_dni/main?IzL1Dx9w5BxmCEtw5A9c6Bnb=CEt1CzAwJyHx4yjbTq9vCBtuTt9fCcPbUo..&IyLlCc5f5w5fCES.=DBny4Y9wDctV4zAxJybk5EA0&CBAkTA9f5o..=m2286&4c9lTFLwDctm=2&TFbm5B5xCcLw6B9k=0"
        try:
            self.page.goto(tra_cuu_url, wait_until="domcontentloaded", timeout=25000)
            self.page.wait_for_timeout(2000)
        except Exception:
            pass

        if not self.page.query_selector("#btnClickLoading"):
            found = False
            for link in self.page.query_selector_all("a"):
                if "Tra cứu văn bản" in link.inner_text():
                    link.click()
                    found = True
                    break
            
            if not found:
                self.page.evaluate('if (typeof link === "function") { link("m2286","DBny4Y9wDctV4zAxJybk5EA0","","",0,0,"&4c9lTFLwDctm=2&TFbm5B5xCcLw6B9k=0"); }')

            self.page.wait_for_timeout(3000)

        self.page.wait_for_selector("#btnClickLoading", timeout=30000)
        print("Đã vào trang Tra cứu văn bản thành công.")

    def search(self, trich_yeu=None, so_ky_hieu=None, from_date=None, to_date=None, loai_vb=None, max_results=50):
        """
        Tra cứu văn bản theo các tiêu chí:
        - trich_yeu: Từ khóa trong trích yếu (vd: 'cảnh báo', 'xác minh')
        - so_ky_hieu: Số ký hiệu văn bản (vd: '5977/TCS2-QLDN')
        - from_date: Ngày ban hành hoặc ngày đến bắt đầu (DD/MM/YYYY)
        - to_date: Ngày ban hành hoặc ngày đến kết thúc (DD/MM/YYYY)
        - loai_vb: Loại văn bản ('Văn bản đến', 'Văn bản đi')
        """
        self.navigate_to_tra_cuu()

        # Reset form tìm kiếm trước
        try:
            self.page.evaluate('if (typeof resetInput === "function") { resetInput(); }')
            self.page.wait_for_timeout(1000)
        except Exception:
            pass

        print(f"Thiết lập điều kiện tra cứu: trích yếu='{trich_yeu}', số KH='{so_ky_hieu}', từ={from_date}, đến={to_date}")

        if trich_yeu:
            self.page.fill("#txt_trichyeu", trich_yeu)

        if so_ky_hieu:
            self.page.fill("#txt_so_kyhieu", so_ky_hieu)

        if from_date:
            if self.page.query_selector("#txt_start_date_banhanh"):
                self.page.fill("#txt_start_date_banhanh", from_date)
            if self.page.query_selector("#txt_start_date_ngayden"):
                self.page.fill("#txt_start_date_ngayden", from_date)

        if to_date:
            if self.page.query_selector("#txt_end_date_banhanh"):
                self.page.fill("#txt_end_date_banhanh", to_date)
            if self.page.query_selector("#txt_end_date_ngayden"):
                self.page.fill("#txt_end_date_ngayden", to_date)

        # Chọn số bản ghi hiển thị nếu có
        try:
            select_rec = self.page.query_selector("#selectPageRec")
            if select_rec:
                select_rec.select_option("50")
        except Exception:
            pass

        # Nhấn nút Tìm kiếm
        print("Đang thực hiện tra cứu...")
        self.page.click("#btnClickLoading")
        self.page.wait_for_timeout(4000)

        # Trích xuất danh sách kết quả từ bảng
        results = []
        tables = self.page.query_selector_all("table")
        target_table = None

        for t in tables:
            headers = [h.inner_text().strip() for h in t.query_selector_all("th")]
            if any("Trích yếu" in h for h in headers) or any("Số ký hiệu" in h for h in headers):
                target_table = t
                break

        if not target_table and len(tables) > 7:
            target_table = tables[7]

        if not target_table:
            print("Không tìm thấy bảng kết quả tra cứu.")
            return []

        rows = target_table.query_selector_all("tr")
        print(f"Tổng số dòng trong bảng kết quả: {len(rows) - 1}")

        for r in rows[1:]:
            cells = r.query_selector_all("td")
            if len(cells) < 5:
                continue

            stt = cells[0].inner_text().strip()
            so_eoffice = cells[1].inner_text().strip()
            trich_yeu_text = cells[2].inner_text().strip().replace("\n", " ")
            so_den = cells[3].inner_text().strip()
            so_ky_hieu_text = cells[4].inner_text().strip()

            # Lấy doc_id từ link allFileDownload
            doc_id = None
            for a in r.query_selector_all("a"):
                href = a.get_attribute("href") or ""
                onclick = a.get_attribute("onclick") or ""
                m = re.search(r"allFileDownload\((\d+)\)", href + onclick)
                if m:
                    doc_id = m.group(1)
                    break

            item = {
                "stt": stt,
                "so_eoffice": so_eoffice,
                "trich_yeu": trich_yeu_text,
                "so_den": so_den,
                "so_ky_hieu": so_ky_hieu_text,
                "doc_id": doc_id,
            }
            results.append(item)
            if len(results) >= max_results:
                break

        return results

    def download_document(self, doc_id: str, output_dir: Path, filename_prefix: str = None) -> list:
        """
        Tải toàn bộ tệp đính kèm của một văn bản qua doc_id.
        Tự động giải nén nếu file tải về là .zip và bảo toàn các tệp PDF/Word.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        downloaded_files = []

        print(f"Đang tải văn bản doc_id={doc_id}...")
        try:
            with self.page.expect_download(timeout=30000) as download_info:
                self.page.evaluate(f"allFileDownload({doc_id})")
            
            download = download_info.value
            original_filename = download.suggested_filename or f"doc_{doc_id}.zip"
            
            # Đặt tên tệp lưu trữ
            safe_prefix = re.sub(r'[\\/*?:"<>|]', "_", filename_prefix) if filename_prefix else f"doc_{doc_id}"
            save_path = output_dir / f"{safe_prefix}_{original_filename}"
            download.save_as(str(save_path))
            print(f"  -> Đã tải tệp về: {save_path.name}")
            downloaded_files.append(save_path)

            # Nếu là file zip, tự động giải nén để lấy các file PDF/DOCX
            if save_path.suffix.lower() == ".zip":
                extract_folder = output_dir / f"{safe_prefix}_files"
                extract_folder.mkdir(parents=True, exist_ok=True)
                try:
                    with zipfile.ZipFile(save_path, 'r') as zip_ref:
                        zip_ref.extractall(extract_folder)
                    print(f"  -> Đã giải nén tệp đính kèm vào thư mục: {extract_folder.name}")
                    for root, _, files in os.walk(extract_folder):
                        for f in files:
                            downloaded_files.append(Path(root) / f)
                except Exception as ze:
                    print(f"  Lỗi khi giải nén ZIP: {ze}")

        except Exception as e:
            print(f"Lỗi khi tải văn bản {doc_id}: {e}")

        return downloaded_files

    def search_and_download_batch(self, trich_yeu=None, so_ky_hieu=None, from_date=None, to_date=None, output_dir=None, max_docs=20) -> list:
        """
        Quy trình trọn gói: Tra cứu theo điều kiện và tự động tải toàn bộ văn bản về thư mục đích.
        """
        docs = self.search(trich_yeu=trich_yeu, so_ky_hieu=so_ky_hieu, from_date=from_date, to_date=to_date, max_results=max_docs)
        if not docs:
            print("Không tìm thấy văn bản nào thỏa mãn điều kiện tra cứu.")
            return []

        print(f"\nTìm thấy {len(docs)} văn bản. Bắt đầu tải tệp về...")
        download_summary = []

        for idx, doc in enumerate(docs, 1):
            doc_id = doc.get("doc_id")
            if not doc_id:
                print(f"[{idx}/{len(docs)}] Văn bản {doc.get('so_ky_hieu')} không có tệp đính kèm tải được.")
                continue

            prefix = f"[{doc.get('so_den', idx)}]_{doc.get('so_ky_hieu', 'CV')}"
            files = self.download_document(doc_id, output_dir, filename_prefix=prefix)
            doc["downloaded_files"] = [str(f) for f in files]
            download_summary.append(doc)

        return download_summary
