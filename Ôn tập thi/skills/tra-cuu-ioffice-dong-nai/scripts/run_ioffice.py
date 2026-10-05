import argparse
import json
import sys
from pathlib import Path

# Thêm thư mục hiện tại vào sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from ioffice_client import IOfficeClient

def main():
    parser = argparse.ArgumentParser(description="Công cụ tự động hóa tra cứu và tải văn bản VNPT iOffice Đồng Nai")
    parser.add_argument("--action", choices=["login", "search", "download", "status"], default="search", help="Hành động cần thực hiện")
    parser.add_argument("--keyword", "--trichyeu", dest="keyword", default=None, help="Từ khóa tìm kiếm trong trích yếu (vd: 'cảnh báo')")
    parser.add_argument("--sokyhieu", dest="sokyhieu", default=None, help="Số ký hiệu văn bản cần tìm")
    parser.add_argument("--from-date", dest="from_date", default=None, help="Ngày bắt đầu (DD/MM/YYYY)")
    parser.add_argument("--to-date", dest="to_date", default=None, help="Ngày kết thúc (DD/MM/YYYY)")
    parser.add_argument("--output-dir", dest="output_dir", default=r"D:\HO SO CONG VAN CANH BAO\Chuyen_doi", help="Thư mục lưu tệp tải về")
    parser.add_argument("--max-docs", dest="max_docs", type=int, default=10, help="Số lượng văn bản tối đa cần xử lý")
    parser.add_argument("--headless", action="store_true", default=True, help="Chạy ở chế độ không mở cửa sổ trình duyệt (mặc định)")
    parser.add_argument("--headed", dest="headless", action="store_false", help="Mở cửa sổ trình duyệt hiển thị")

    args = parser.parse_args()

    client = IOfficeClient(headless=args.headless)
    with client:
        if args.action == "status":
            logged_in = client.is_logged_in()
            print(f"Trạng thái phiên làm việc iOffice: {'ĐANG ĐĂNG NHẬP' if logged_in else 'CHƯA ĐĂNG NHẬP / HẾT HẠN'}")
            return

        if args.action == "login":
            success = client.login()
            if success:
                print("Đăng nhập thành công và đã cập nhật session.")
            else:
                print("Đăng nhập không thành công.")
            return

        if args.action == "search":
            client.ensure_logged_in()
            results = client.search(
                trich_yeu=args.keyword,
                so_ky_hieu=args.sokyhieu,
                from_date=args.from_date,
                to_date=args.to_date,
                max_results=args.max_docs
            )
            print(f"\n--- KẾT QUẢ TRA CỨU ({len(results)} VĂN BẢN) ---")
            for idx, r in enumerate(results, 1):
                print(f"[{idx}] Số KH: {r.get('so_ky_hieu')} | Số đến: {r.get('so_den')} | doc_id: {r.get('doc_id')}")
                print(f"    Trích yếu: {r.get('trich_yeu')}")
            return

        if args.action == "download":
            client.ensure_logged_in()
            print(f"Thư mục lưu trữ: {args.output_dir}")
            summary = client.search_and_download_batch(
                trich_yeu=args.keyword,
                so_ky_hieu=args.sokyhieu,
                from_date=args.from_date,
                to_date=args.to_date,
                output_dir=Path(args.output_dir),
                max_docs=args.max_docs
            )
            print(f"\n--- BÁO CÁO KẾT QUẢ TẢI VĂN BẢN ({len(summary)} ĐÃ TẢI) ---")
            for idx, item in enumerate(summary, 1):
                print(f"[{idx}] {item.get('so_ky_hieu')} (Đến: {item.get('so_den')}):")
                for f in item.get("downloaded_files", []):
                    print(f"    + {f}")

if __name__ == "__main__":
    main()
