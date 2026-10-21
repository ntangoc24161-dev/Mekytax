import json
import os
import re

with open('scan_result.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Company definitions
companies_def = [
    {
        "folder_name": "1. SON TIEN HAI",
        "norm_name": "SON_TIEN_HAI_3603859012",
        "ten_nnt": "CÔNG TY TNHH THƯƠNG MẠI SẢN XUẤT SƠN TIỀN HẢI",
        "mst": "3603859012",
        "has_tokhai_sub": None
    },
    {
        "folder_name": "2. T&T BHOA",
        "norm_name": "TT_BIEN_HOA_3603883505",
        "ten_nnt": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ T&T BIÊN HÒA",
        "mst": "3603883505",
        "has_tokhai_sub": ["3603883505", "3603883505-TT"]
    },
    {
        "folder_name": "3. PHUC NGUYEN TIEN",
        "norm_name": "PHUC_NGUYEN_TIEN_3603698904",
        "ten_nnt": "CÔNG TY TNHH PHÚC NGUYÊN TIẾN",
        "mst": "3603698904",
        "has_tokhai_sub": ["3603698904", "3603698904-PHUC NGUYEN TIEN"]
    },
    {
        "folder_name": "4. TAM HOANG PHAT",
        "norm_name": "TAM_HOANG_PHAT_3603721832",
        "ten_nnt": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ TÂM HOÀNG PHÁT",
        "mst": "3603721832",
        "has_tokhai_sub": None
    },
    {
        "folder_name": "5. HUNG ANH PHAT",
        "norm_name": "HUNG_ANH_PHAT_3602451559",
        "ten_nnt": "CÔNG TY TNHH MTV SX-TM-DV HÙNG ANH PHÁT",
        "mst": "3602451559",
        "has_tokhai_sub": None
    },
    {
        "folder_name": "6. DAI HUY HOANG",
        "norm_name": "DAI_HUY_HOANG_3602179712",
        "ten_nnt": "CÔNG TY TNHH ĐẠI HUY HOÀNG",
        "mst": "3602179712",
        "has_tokhai_sub": ["3602179712"]
    },
    {
        "folder_name": "7. NHAT VINH-PASS LAM",
        "norm_name": "NHAT_VINH_3601946069",
        "ten_nnt": "CÔNG TY TNHH MTV NHẤT VINH",
        "mst": "3601946069",
        "has_tokhai_sub": ["3601946069"]
    },
    {
        "folder_name": "8. HUY THANH DAT",
        "norm_name": "HUY_THANH_PHAT_3603820015",
        "ten_nnt": "CÔNG TY TNHH TM DV VÀ SX HUY THÀNH PHÁT",
        "mst": "3603820015",
        "has_tokhai_sub": ["3603820015", "3603820015-HUY THANH PHAT"]
    },
    {
        "folder_name": "9. TRIQUYEN",
        "norm_name": "TRI_QUYEN_3603157158",
        "ten_nnt": "CÔNG TY TNHH TM VÀ DV TRÍ QUYÊN",
        "mst": "3603157158",
        "has_tokhai_sub": None
    },
    {
        "folder_name": ". Phat hoang long",
        "norm_name": "PHAT_HOANG_LONG_3600974507",
        "ten_nnt": "CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI PHÁT HOÀNG LONG",
        "mst": "3600974507",
        "has_tokhai_sub": None
    },
    {
        "folder_name": "HO SO KE KOACH\\2. An Phuong Tam",
        "norm_name": "AN_PHUONG_TAM_3602694223",
        "ten_nnt": "CÔNG TY TNHH MỘT THÀNH VIÊN AN PHƯƠNG TÂM",
        "mst": "3602694223",
        "has_tokhai_sub": None
    },
    {
        "folder_name": "To khai\\3600840461",
        "norm_name": "THANH_THANH_PHUONG_3600840461",
        "ten_nnt": "CÔNG TY TNHH THANH THANH PHƯƠNG",
        "mst": "3600840461",
        "has_tokhai_sub": ["3600840461", "3600840461-CÔNG TY TNHH THANH THANH PHƯƠNG"]
    },
    {
        "folder_name": "To khai\\3600854351",
        "norm_name": "VINH_TRUONG_PHAT_3600854351",
        "ten_nnt": "CÔNG TY TNHH SX - TM VĨNH TRƯỜNG PHÁT",
        "mst": "3600854351",
        "has_tokhai_sub": ["3600854351"]
    }
]

# Let's inspect files for each
analysis_results = []

for c in companies_def:
    c_files = []
    # files in primary folder
    if c['folder_name'] in data:
        c_files.extend([f['rel'] for f in data[c['folder_name']]['files']])
    elif "\\" in c['folder_name']:
        parts = c['folder_name'].split("\\")
        parent = parts[0]
        sub = parts[1]
        if parent in data:
            c_files.extend([f['rel'] for f in data[parent]['files'] if f['rel'].startswith(sub)])
    
    # files in tokhai subfolders
    if c['has_tokhai_sub']:
        for tk_sub in c['has_tokhai_sub']:
            for f in data.get('To khai', {}).get('files', []):
                if f['rel'].startswith(tk_sub):
                    c_files.append(f"To khai\\{f['rel']}")

    # Categorize into 4 steps
    step1_xml = [f for f in c_files if f.lower().endswith('.xml')]
    step1_bangke = [f for f in c_files if ('bang ke' in f.lower() or 'bảng kê' in f.lower() or 'hddt' in f.lower() or 'br' in f.lower() or 'mv' in f.lower() or 'mua vao' in f.lower() or 'ban ra' in f.lower()) and f.lower().endswith(('.xlsx', '.xls', '.csv', '.zip'))]
    step1_nghiavu = [f for f in c_files if 'etax' in f.lower() or 'ptdn' in f.lower() or 'thue' in f.lower()]
    
    step2_soke_toan = [f for f in c_files if ('so chi tiet' in f.lower() or 'sổ chi tiết' in f.lower() or 'nkc' in f.lower() or 'nhat ky chung' in f.lower() or 'so cai' in f.lower() or 'sổ cái' in f.lower() or 'sctk' in f.lower() or 'cdps' in f.lower() or 'can doi' in f.lower()) and f.lower().endswith(('.xlsx', '.xls'))]
    step2_kho_giathanh = [f for f in c_files if ('nxt' in f.lower() or 'kho' in f.lower() or 'ton kho' in f.lower() or 'gia thanh' in f.lower() or 'vat tu' in f.lower())]
    step2_nganhang = [f for f in c_files if ('sao ke' in f.lower() or 'sao kê' in f.lower() or 'so phu' in f.lower() or 'sổ phụ' in f.lower() or 'so ngan hang' in f.lower() or 'sổ ngân hàng' in f.lower() or 'vcb' in f.lower() or 'bidv' in f.lower() or 'eximbank' in f.lower() or 'hdbank' in f.lower())]
    step2_luong_tscd = [f for f in c_files if ('luong' in f.lower() or 'lương' in f.lower() or 'tscd' in f.lower() or 'khau hao' in f.lower() or 'khấu hao' in f.lower() or 'ccdc' in f.lower() or 'bang pb' in f.lower() or 'bảng pb' in f.lower())]

    step3_cbrr = [f for f in c_files if ('cbrr' in f.lower() or 'cong an' in f.lower() or 'f0' in f.lower() or 'f1' in f.lower())]
    step3_danhba = [f for f in c_files if ('danh ba' in f.lower() or 'danh bạ' in f.lower())]
    step3_tncn_danguon = [f for f in c_files if ('tncn' in f.lower() and ('da nguon' in f.lower() or 'tra cuu' in f.lower()))]
    step3_giaitrinh = [f for f in c_files if ('bien ban' in f.lower() or 'biên bản' in f.lower() or 'qd' in f.lower() or 'quyet dinh' in f.lower() or 'tong hop' in f.lower() or 'giai trinh' in f.lower())]

    step4_ketqua = [f for f in c_files if 'ket-qua' in f.lower() or 'phan_tich_ho_so' in f.lower() or 'cau-thanh' in f.lower()]

    res = {
        "mst": c['mst'],
        "ten_nnt": c['ten_nnt'],
        "folder_name": c['folder_name'],
        "norm_name": c['norm_name'],
        "total_files": len(c_files),
        "step1_xml": len(step1_xml),
        "step1_bangke": len(step1_bangke),
        "step1_nghiavu": len(step1_nghiavu),
        "step2_soke_toan": len(step2_soke_toan),
        "step2_kho_giathanh": len(step2_kho_giathanh),
        "step2_nganhang": len(step2_nganhang),
        "step2_luong_tscd": len(step2_luong_tscd),
        "step3_cbrr": len(step3_cbrr),
        "step3_danhba": len(step3_danhba),
        "step3_tncn_danguon": len(step3_tncn_danguon),
        "step3_giaitrinh": len(step3_giaitrinh),
        "step4_ketqua": len(step4_ketqua)
    }
    analysis_results.append(res)

with open('company_analysis_results.json', 'w', encoding='utf-8') as f:
    json.dump(analysis_results, f, ensure_ascii=False, indent=2)

print("ANALYSIS SUMMARY FOR ALL COMPANIES:")
for r in analysis_results:
    print(f"\n[{r['mst']}] {r['ten_nnt']} (Folder: {r['folder_name']})")
    print(f"  Total files mapped: {r['total_files']}")
    print(f"  Step 1: XML={r['step1_xml']}, BangKe={r['step1_bangke']}, NghiaVu={r['step1_nghiavu']}")
    print(f"  Step 2: SoKT={r['step2_soke_toan']}, Kho-Z={r['step2_kho_giathanh']}, NganHang={r['step2_nganhang']}, Luong-TSCD={r['step2_luong_tscd']}")
    print(f"  Step 3: CBRR={r['step3_cbrr']}, DanhBa={r['step3_danhba']}, TNCN-DN={r['step3_tncn_danguon']}, GiaiTrinh={r['step3_giaitrinh']}")
    print(f"  Step 4: Master/CauThanh={r['step4_ketqua']}")
