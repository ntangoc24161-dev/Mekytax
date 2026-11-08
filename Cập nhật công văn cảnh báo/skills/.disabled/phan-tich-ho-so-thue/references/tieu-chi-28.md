# 28 tiêu chí rủi ro (sheet TUDIEN, tool 4.4)

Nguồn: `2. PHAN MEM ... 4.4.xlsb` / sheet `TUDIEN` (`TC1_N` ... `TC28_N`). Không bịa thêm ngưỡng.

| ID | Tên | Ngưỡng ghi trên TUDIEN |
|---|---|---|
| TC1 | Chênh lệch DT GTGT [34] vs DT BCTC KQKD [01] | lệch ≠ 0 |
| TC2 | Người mua trả tiền trước > 5% DT | 5% |
| TC3 | HTK + XDCB dở dang > 30% DT | 30% |
| TC4 | DT chưa thực hiện > 0 | > 0 |
| TC5 | DT thuần - giá vốn < 0 | < 0 |
| TC6 | Chi phí phải trả trích trước lớn | số dư [315][333] |
| TC7 | Lãi vay chưa vốn hóa | [230][242] |
| TC8 | Ưu đãi / miễn giảm TNDN | chỉ tiêu C2/C11 mẫu 03/TNDN |
| TC9 | VAT khấu trừ TK vs BCTC [152] | [22]/[43] vs đầu/cuối [152] |
| TC10 | VAT khấu trừ vs 10%×(HTK+TSCĐ) | ×10% |
| TC11 | Có DT KCT [26] nhưng [25]=[24] | |
| TC12 | Điều chỉnh giảm LN trước thuế | B7/B8 > ngưỡng (số không ghi) |
| TC13 | Trích lập dự phòng | sum mã dự phòng |
| TC14 | CN vốn [22]>0 và [25]<0 | |
| TC15 | Bù trừ lãi BĐS / lỗ SXKD | |
| TC16 | Chuyển giá: [251][252]>0 và [421]<0 | |
| TC17 | CP lãi vay vs 30% EBITDA | |
| TC18 | B11>0 không có PL GDLK | |
| TC19 | Thuế TNDN hoãn lại | (mô tả trống trên TUDIEN) |
| TC20 | Phải trả người bán lớn | [311]+[331] |
| TC21 | Lãi vay khi thiếu VĐL | cần vốn điều lệ TMS - **bỏ** |
| TC22 | QTT TNCN vs TK tháng/quý | |
| TC23 | Dòng tiền chi NLĐ vs TNCT | 20% (TC23_M) |
| TC24 | KHBS vs hạch toán TMS | **không chạy** (cấm TMS) |
| TC25 | Bổ sung GTGT chưa điều chỉnh | |[22]+[25]-[35]-[37]+[38]| vs |[40]+[42]+[43]| |
| TC26 | Tỷ trọng HĐ mua NNT TT 03/06 | cần trạng thái MST - **bỏ nếu không có file** |
| TC27 | DT TK GTGT vs HĐĐT bán ra | lệch ≠ 0 |
| TC28 | Phân tích ngang/dọc BCTC | TC28_M1 giá vốn tăng nhanh hơn DT |

Engine v1 chạy: TC1, TC2, TC3, TC4, TC5, TC9, TC10, TC11, TC27, cộng M1 (trùng HĐ, suất, âm, ngoài kỳ, lệch mua vào).
