# Dap an tu giai BT5-BT9 (CD1 Excel kiem tra thue)

Tao luc: 2026-10-07 18:08 Asia/Ho_Chi_Minh
Nguon du lieu: Du lieu_HV (ban hoc vien, chua co dap an GV)

## Ket qua chay
- BT5: 1240 dong; 9 dong trung; 4 nhom -> BT5_Hoa_don_trung_DAP_AN.xlsx
- BT6: 751 dong; 417 MST chinh -> BT6_Chuan_hoa_MST_DAP_AN.xlsx
- BT7A: 751 dong; 62 dong/>20tr; 34 nhom -> BT7A_HD_tren_20tr_DAP_AN.xlsx
- BT7B: 610 dong; 182 dong/>5tr; 144 nhom -> BT7B_HD_tren_5tr_DAP_AN.xlsx
- BT8: matched=407, BKMV_NA=529, HDDT_NA=17, thueBK=3153428532.0, no133=3153909044.0 -> BT8_Doi_chieu_mua_vao_HDDT_BKMV_NKC_DAP_AN.xlsx
- BT9: matched=119, NA=0, thueBK=178228430.0, no133=178228430.0 -> BT9_DN_thuong_mai_BKMV_NKC_DAP_AN.xlsx

## Cong thuc chot

### BT5
- `=COUNTIFS($F$9:$F$last;F9;$I$9:$I$last;I9;$K$9:$K$last;K9)`
- Khoa: So HD + MST + Gia tri chua thue

### BT6
- `=LEN(A2)`
- `=IF(OR(B2=10;B2=14);A2;IF(B2=9;"0"&A2;IF(B2=12;"0"&LEFT(A2;9)&"-"&RIGHT(A2;3);IF(MID(A2;10;1)="-";"0"&A2;LEFT(A2;10)&"-"&RIGHT(A2;3)))))`

### BT7A (>20tr)
- `J=G+I`
- `=SUMIFS($J$4:$J$last;$F$4:$F$last;F4;$D$4:$D$last;D4)`
- Flag neu tong nhom > 20000000

### BT7B (>5tr)
- `J=G+I`
- `=SUMIFS($J$4:$J$last;$E$4:$E$last;E4;$C$4:$C$last;C4)`
- Flag neu tong nhom > 5000000

### BT8
- Khoa: `SoHD(bo 0 dau) & "|" & MST_chuan`
- `KTRA`: KHOP / #N/A khi doi chieu HDDT <-> BKMV
- DOI CHIEU thang: PS No TK 133* (NKC) vs Tien thue BKMV

### BT9
- Khoa: `SoHD|GiaTri` (va thu `SoHD|TienThue` voi dong 133)
- DOI CHIEU thang: PS No 133 vs Tien thue BKMV nam 2020

## Luu y
- File HV khong co cong thuc VLOOKUP san; BT8/BT9 dung khoa tuong duong logic slide.
- May dung Excel tieng Viet co the can doi dau phay trong cong thuc thanh dau cham phay `;`.
- Day la dap an tu giai de doi chieu slide, khong phai file GV chinh thuc.
