---
type: wiki
status: active
tags: [wiki, TTKT, playbook, skill]
created: '2026-10-15'
updated: '2026-10-15'
source: '[[phan-tich-ho-so-thue]]'
---
# Playbook lệnh các skill phân tích hồ sơ thuế

Trang này là **mục lục** lệnh copy-paste của mọi skill phục vụ phân tích hồ sơ kiểm tra thuế. Lệnh đầy đủ nằm ở file vận hành `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md`. Chạy theo bước khi công ty mới: `skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md`. Bản đồ lớp vận hành: [[Bản đồ tổng hợp quy trình phân tích hồ sơ thuế]].

Cập nhật: **2026-10-15**.

Placeholder: `<FOLDER>` = đường dẫn tuyệt đối `<TEN>_<MST>`; `<MST>` `<TEN>` `<KY>`.

---

## Đường ống 0-10

| Bước | Skill | Playbook riêng |
|---|---|---|
| 0 | `to-chuc-ho-so-kiem-tra-thue` | `skills/to-chuc-ho-so-kiem-tra-thue/playbook-lenh.md` |
| 0b+1 | `tong-hop-dau-vao-hddt-bang-ke-nld` | `skills/tong-hop-dau-vao-hddt-bang-ke-nld/playbook-lenh.md` |
| 1b | `lam-sach-cd1-mst-hoa-don` | `skills/lam-sach-cd1-mst-hoa-don/playbook-lenh.md` |
| 2 | `doi-chieu-gtgt-tndn-cbrr` | `skills/doi-chieu-gtgt-tndn-cbrr/playbook-lenh.md` |
| 3 | `doi-chieu-ncc-mua-vao-rui-ro` | `skills/doi-chieu-ncc-mua-vao-rui-ro/playbook-lenh.md` |
| 4 | `tong-hop-tncn-da-nguon` | `skills/tong-hop-tncn-da-nguon/playbook-lenh.md` |
| 5 | `doi-chieu-511-632` | `skills/doi-chieu-511-632/playbook-lenh.md` |
| 6 | `phan-tich-xnt-kho` | `skills/phan-tich-xnt-kho/playbook-lenh.md` |
| 7 | `phan-tich-gia-thanh-dinh-muc` | `skills/phan-tich-gia-thanh-dinh-muc/playbook-lenh.md` |
| 8 | `phan-tich-cong-no-tien-mat` | `skills/phan-tich-cong-no-tien-mat/playbook-lenh.md` |
| 9-10 | `phan-tich-ho-so-thue` | `skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md` |

Skill trục giữ Master **8 sheet**, **28 tiêu chí**, đề cương. [[Phân tích hồ sơ kiểm tra thuế theo module độc lập]]

---

## Skill cạnh

| Skill | Việc | Playbook |
|---|---|---|
| `mo-ho-so-ttkt` | Checklist nghiệp vụ A-F, GDLK, nhánh xử lý | `skills/mo-ho-so-ttkt/playbook-lenh.md` |
| `tong-hop-ben-mua-tu-so-cvcbrr` | Bên mua trên sổ CVCBRR | `skills/tong-hop-ben-mua-tu-so-cvcbrr/playbook-lenh.md` |
| `quan-ly-rui-ro-tuan-thu-thue` | Khung chọn đối tượng / bộ chỉ số TCT | `skills/quan-ly-rui-ro-tuan-thu-thue/playbook-lenh.md` |

Không dùng cho phân tích 1 DN: `xu-phat-vphc-thue-hoa-don`, `tra-loi-cong-van`. Chính sách TNCN: skill chung `thue-tncn`.

---

## Câu siêu ngắn

| Skill | Câu |
|---|---|
| `to-chuc-ho-so-kiem-tra-thue` | `Tổ chức / kiểm kê checklist <FOLDER>` |
| `tong-hop-dau-vao-hddt-bang-ke-nld` | `Tổng hợp đầu vào HĐĐT-Bảng kê-NLĐ <FOLDER>` |
| `lam-sach-cd1-mst-hoa-don` | `Chuẩn hóa MST + quét HĐ trùng <FOLDER>` |
| `doi-chieu-gtgt-tndn-cbrr` | `Đối chiếu CBRR GTGT-TNDN-HĐĐT <FOLDER>` |
| `doi-chieu-ncc-mua-vao-rui-ro` | `Đối chiếu NCC F0/F1/CA chuẩn VPT <FOLDER>` |
| `tong-hop-tncn-da-nguon` | `Tổng hợp TNCN <MST> <FOLDER>` |
| `doi-chieu-511-632` | `Đối chiếu 511-632 <FOLDER>` |
| `phan-tich-xnt-kho` | `Phân tích XNT kho <FOLDER>` |
| `phan-tich-gia-thanh-dinh-muc` | `Phân tích giá thành/định mức <FOLDER>` |
| `phan-tich-cong-no-tien-mat` | `Quét 111, cụm cùng ngày NCC và công nợ 131/331 <FOLDER>` |
| `phan-tich-ho-so-thue` | `Tái lập Master 8 sheet đủ 28 TC chuẩn VPT <FOLDER>` |
| `mo-ho-so-ttkt` | `Checklist mở hồ sơ TTKT A-F <FOLDER>` |

Một lệnh = một file. Thiếu nguồn thì ghi thiếu, không để số 0 giả.
