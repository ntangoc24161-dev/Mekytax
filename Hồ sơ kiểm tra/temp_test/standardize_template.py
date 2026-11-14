import docx
import re
import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_field_mapping(instr, context_loc=""):
    raw = instr.strip()
    clean = re.sub(r'\s+', ' ', raw)
    
    # Group 1: Thong tin NNT & CQT
    if 'Tên_công_ty' in clean:
        return 'MERGEFIELD Ten_NNT', '«Ten_NNT»'
    if 'MST' in clean:
        return 'MERGEFIELD MST', '«MST»'
    if 'địa_chỉ' in clean.lower() or 'địa_chỉ' in clean:
        return 'MERGEFIELD Dia_chi', '«Dia_chi»'
    if 'NNKD' in clean:
        return 'MERGEFIELD Nganh_nghe', '«Nganh_nghe»'
    if 'GĐ' in clean:
        return 'MERGEFIELD Giam_doc', '«Giam_doc»'
    if 'CCCDKT' in clean:
        return 'MERGEFIELD CCCD_KT', '«CCCD_KT»'
    if 'CCCD' in clean:
        return 'MERGEFIELD CCCD_GD', '«CCCD_GD»'
    if 'ngày_cấp2' in clean.lower():
        return 'MERGEFIELD Ngay_cap_DKKD \\@ "dd/MM/yyyy"', '«Ngay_cap_DKKD»'
    if 'ngày_cấp' in clean.lower():
        return 'MERGEFIELD Ngay_cap_CCCD \\@ "dd/MM/yyyy"', '«Ngay_cap_CCCD»'
    if 'Nơi_cấp' in clean:
        return 'MERGEFIELD Noi_cap_CCCD', '«Noi_cap_CCCD»'
    if 'KT' in clean:
        return 'MERGEFIELD Ke_toan_truong', '«Ke_toan_truong»'
    if 'GGCNĐKKD' in clean:
        return 'MERGEFIELD So_DKKD', '«So_DKKD»'
    if 'Tài_khoản_số' in clean:
        return 'MERGEFIELD Tai_khoan_NH', '«Tai_khoan_NH»'
    if 'Chi_nhánh' in clean:
        return 'MERGEFIELD Chi_nhanh', '«Chi_nhanh»'
    if 'Ma_CQT' in clean:
        return 'MERGEFIELD Ma_CQT', '«Ma_CQT»'
    if 'CQT' in clean:
        return 'MERGEFIELD Ten_CQT', '«Ten_CQT»'
        
    # Group 2: Doan kiem tra & Phan cong
    if 'CVTĐ' in clean:
        return 'MERGEFIELD Chuc_vu_TD', '«Chuc_vu_TD»'
    if 'TĐ' in clean:
        return 'MERGEFIELD Truong_doan', '«Truong_doan»'
    if 'CVTV' in clean:
        return 'MERGEFIELD Chuc_vu_TV', '«Chuc_vu_TV»'
    if 'TV' in clean:
        return 'MERGEFIELD Thanh_vien', '«Thanh_vien»'
    if 'Nội_dung_kiểm_tra' in clean:
        return 'MERGEFIELD Noi_dung_KT', '«Noi_dung_KT»'
    if 'Phân_công_trưởng_đoàn' in clean:
        return 'MERGEFIELD Phan_cong_TD', '«Phan_cong_TD»'
    if 'Phân_công_thành_viên' in clean:
        return 'MERGEFIELD Phan_cong_TV', '«Phan_cong_TV»'
        
    # Group 3: Tien trinh & Moc van ban
    if 'QĐKT' in clean:
        return 'MERGEFIELD So_QDKT', '«So_QDKT»'
    if 'ngày_qđ' in clean.lower() or 'ngay_qđ' in clean.lower():
        return 'MERGEFIELD Ngay_QDKT \\@ "dd/MM/yyyy"', '«Ngay_QDKT»'
    if 'QĐGS' in clean:
        return 'MERGEFIELD So_QDGS', '«So_QDGS»'
    if 'Thời_hạn_kiểm_tra' in clean:
        return 'MERGEFIELD Thoi_han_KT', '«Thoi_han_KT»'
    if 'Thời_kỳ_kiểm_tra' in clean:
        return 'MERGEFIELD Thoi_ky_KT', '«Thoi_ky_KT»'
    if 'Ngày_công_bố' in clean:
        return 'MERGEFIELD Ngay_cong_bo \\@ "dd/MM/yyyy"', '«Ngay_cong_bo»'
    if 'Từ_ngày' in clean:
        return 'MERGEFIELD Ngay_bat_dau \\@ "dd/MM/yyyy"', '«Ngay_bat_dau»'
    if 'Đến_ngày' in clean:
        return 'MERGEFIELD Ngay_ket_thuc \\@ "dd/MM/yyyy"', '«Ngay_ket_thuc»'
    if 'Ngày_ký_Biên_bản_kiểm_tra' in clean:
        return 'MERGEFIELD Ngay_ky_BBKT \\@ "dd/MM/yyyy"', '«Ngay_ky_BBKT»'
    if 'Ngày_báo_cáo' in clean:
        if '\\@ "MM"' in clean or '@ "MM"' in clean:
            return 'MERGEFIELD Ngay_bao_cao \\@ "MM"', '«Ngay_bao_cao»'
        if '\\@ "dd"' in clean or '@ "dd"' in clean:
            return 'MERGEFIELD Ngay_bao_cao \\@ "dd"', '«Ngay_bao_cao»'
        return 'MERGEFIELD Ngay_bao_cao \\@ "dd/MM/yyyy"', '«Ngay_bao_cao»'

    # Group 4: So lieu tai chinh
    if 'Thuế_TNDN' in clean:
        return 'MERGEFIELD Truy_thu_TNDN \\# "#,##0"', '«Truy_thu_TNDN»'
    if 'Thuế_GTGT' in clean:
        return 'MERGEFIELD Truy_thu_GTGT \\# "#,##0"', '«Truy_thu_GTGT»'
    if 'Tổng_cộng_truy_thu_BC' in clean:
        return 'MERGEFIELD Tong_truy_thu_chu', '«Tong_truy_thu_chu»'
    if 'Tổng_cộng_truy_thu' in clean or ('Tổng' in clean and 'phạt' not in clean):
        return 'MERGEFIELD Tong_truy_thu \\# "#,##0"', '«Tong_truy_thu»'
    if 'Số_tiền_phạt_20_BC' in clean:
        return 'MERGEFIELD Tong_tien_phat_chu', '«Tong_tien_phat_chu»'
    if 'Tổng_phạt_kk_sai' in clean:
        return 'MERGEFIELD Tong_tien_phat \\# "#,##0"', '«Tong_tien_phat»'
    if 'Tổng_cộng_số_phạt__chậm_nộp' in clean:
        return 'MERGEFIELD Tong_tien_phat \\# "#,##0"', '«Tong_tien_phat»'
    if 'Số_tiền_phạt_003_TNDN' in clean:
        return 'MERGEFIELD Cham_nop_TNDN \\# "#,##0"', '«Cham_nop_TNDN»'
    if 'Số_tiền_phạt_003_GTGT' in clean:
        return 'MERGEFIELD Cham_nop_GTGT \\# "#,##0"', '«Cham_nop_GTGT»'
    if 'Tổng_phạt' in clean:
        return 'MERGEFIELD Tong_cham_nop \\# "#,##0"', '«Tong_cham_nop»'
    if 'Phạt_KK_sai_VPHC' in clean:
        return 'MERGEFIELD Phat_thu_tuc \\# "#,##0"', '«Phat_thu_tuc»'
    if 'Số_tiền_phạt_20' in clean:
        return 'MERGEFIELD Phat_20_phan_tram \\# "#,##0"', '«Phat_20_phan_tram»'
    if 'Tổng_cộng_phạt_và_truy_thu_BC' in clean:
        return 'MERGEFIELD Tong_phat_sinh_chu', '«Tong_phat_sinh_chu»'
    if 'Tổng_cộng_phạt_và_truy_thu' in clean:
        return 'MERGEFIELD Tong_phat_sinh \\# "#,##0"', '«Tong_phat_sinh»'
    if 'Giảm_lỗ' in clean:
        return 'MERGEFIELD Giam_lo \\# "#,##0"', '«Giam_lo»'
    if 'giảm_khấu_trừ' in clean:
        return 'MERGEFIELD Giam_khau_tru \\# "#,##0"', '«Giam_khau_tru»'

    return None, None

def make_mergefield_elements(instr_text, display_text, rPr_elem=None):
    runs = []
    
    # 1. begin
    r1 = OxmlElement('w:r')
    if rPr_elem is not None:
        r1.append(copy.deepcopy(rPr_elem))
    fc1 = OxmlElement('w:fldChar')
    fc1.set(qn('w:fldCharType'), 'begin')
    r1.append(fc1)
    runs.append(r1)
    
    # 2. instrText
    r2 = OxmlElement('w:r')
    if rPr_elem is not None:
        r2.append(copy.deepcopy(rPr_elem))
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = f' {instr_text} '
    r2.append(it)
    runs.append(r2)
    
    # 3. separate
    r3 = OxmlElement('w:r')
    if rPr_elem is not None:
        r3.append(copy.deepcopy(rPr_elem))
    fc2 = OxmlElement('w:fldChar')
    fc2.set(qn('w:fldCharType'), 'separate')
    r3.append(fc2)
    runs.append(r3)
    
    # 4. display text
    r4 = OxmlElement('w:r')
    if rPr_elem is not None:
        r4.append(copy.deepcopy(rPr_elem))
    t = OxmlElement('w:t')
    t.text = display_text
    r4.append(t)
    runs.append(r4)
    
    # 5. end
    r5 = OxmlElement('w:r')
    if rPr_elem is not None:
        r5.append(copy.deepcopy(rPr_elem))
    fc3 = OxmlElement('w:fldChar')
    fc3.set(qn('w:fldCharType'), 'end')
    r5.append(fc3)
    runs.append(r5)
    
    return runs

def process_paragraph_existing_fields(p_elem):
    children = list(p_elem)
    
    fields = []
    current_field = None
    
    for i, child in enumerate(children):
        for elem in child.iter():
            tag = elem.tag.split('}')[-1]
            if tag == 'fldChar':
                fc_type = elem.get(f"{{{W_NS}}}fldCharType") or elem.get("w:fldCharType") or elem.get("fldCharType")
                if fc_type == 'begin':
                    current_field = {'begin_idx': i, 'instr_elements': [], 'separate_idx': None, 'text_elements': [], 'end_idx': None}
                elif fc_type == 'separate' and current_field is not None:
                    current_field['separate_idx'] = i
                elif fc_type == 'end' and current_field is not None:
                    current_field['end_idx'] = i
                    fields.append(current_field)
                    current_field = None
                    
        if current_field is not None:
            if current_field['separate_idx'] is None:
                for elem in child.iter():
                    if elem.tag.split('}')[-1] == 'instrText':
                        current_field['instr_elements'].append(elem)
            else:
                for elem in child.iter():
                    if elem.tag.split('}')[-1] == 't':
                        current_field['text_elements'].append(elem)
                        
    updated_count = 0
    for f in fields:
        full_instr = ''.join(it.text or '' for it in f['instr_elements']).strip()
        target_instr, target_display = get_field_mapping(full_instr)
        if target_instr:
            if f['instr_elements']:
                f['instr_elements'][0].text = f" {target_instr} "
                f['instr_elements'][0].set(qn('xml:space'), 'preserve')
                for it in f['instr_elements'][1:]:
                    it.text = ''
            if f['text_elements']:
                f['text_elements'][0].text = target_display
                for t in f['text_elements'][1:]:
                    t.text = ''
            updated_count += 1
            
    return updated_count
