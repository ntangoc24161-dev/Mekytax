import os
import shutil
import zipfile
import re
import copy
from lxml import etree
import docx
import win32com.client

bak_docx = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP GIAI THE.bak-truoc-chuan-hoa-20261113.docx"
target_docx = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP GIAI THE.docx"
excel_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"

# 1. Restore from clean backup
shutil.copy2(bak_docx, target_docx)
print("Restored from clean backup.")

# 2. Clean table rows using python-docx
doc = docx.Document(target_docx)
tables_to_clean = [1, 3, 5, 6, 37, 39, 49, 55]

for t_idx in tables_to_clean:
    if t_idx < len(doc.tables):
        t = doc.tables[t_idx]
        if len(t.rows) > 1:
            row_to_del = t.rows[1]
            tr = row_to_del._tr
            tr.getparent().remove(tr)
            print(f"Deleted redundant row 1 in Table {t_idx} (now {len(t.rows)} rows)")

doc.save(target_docx)
print("Saved table cleanup.")

# 3. XML modifications with lxml
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
w = f"{{{W_NS}}}"

def create_field_runs(field_name, rpr_template=None):
    runs = []
    
    # 1. begin
    r1 = etree.Element(f"{w}r")
    if rpr_template is not None:
        r1.append(copy.deepcopy(rpr_template))
    fc1 = etree.SubElement(r1, f"{w}fldChar")
    fc1.set(f"{w}fldCharType", "begin")
    runs.append(r1)
    
    # 2. instrText
    r2 = etree.Element(f"{w}r")
    if rpr_template is not None:
        r2.append(copy.deepcopy(rpr_template))
    it = etree.SubElement(r2, f"{w}instrText")
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" MERGEFIELD {field_name} "
    runs.append(r2)
    
    # 3. separate
    r3 = etree.Element(f"{w}r")
    if rpr_template is not None:
        r3.append(copy.deepcopy(rpr_template))
    fc2 = etree.SubElement(r3, f"{w}fldChar")
    fc2.set(f"{w}fldCharType", "separate")
    runs.append(r3)
    
    # 4. display text
    r4 = etree.Element(f"{w}r")
    if rpr_template is not None:
        r4.append(copy.deepcopy(rpr_template))
    t = etree.SubElement(r4, f"{w}t")
    t.text = f"«{field_name}»"
    runs.append(r4)
    
    # 5. end
    r5 = etree.Element(f"{w}r")
    if rpr_template is not None:
        r5.append(copy.deepcopy(rpr_template))
    fc3 = etree.SubElement(r5, f"{w}fldChar")
    fc3.set(f"{w}fldCharType", "end")
    runs.append(r5)
    
    return runs

def replace_text_with_field(root, target_text, field_name):
    count = 0
    # Search for all w:t containing target_text
    for t_elem in root.xpath(f".//w:t[contains(text(), '{target_text}')]", namespaces={"w": W_NS}):
        text = t_elem.text
        if not text or target_text not in text:
            continue
        
        r_elem = t_elem.getparent()
        if r_elem.tag != f"{w}r":
            continue
        p_elem = r_elem.getparent()
        if p_elem is None:
            continue
            
        rpr = r_elem.find(f"{w}rPr")
        parts = text.split(target_text)
        
        # Keep the first part in t_elem
        t_elem.text = parts[0]
        curr_r = r_elem
        
        for part in parts[1:]:
            count += 1
            # insert field runs after curr_r
            field_runs = create_field_runs(field_name, rpr)
            for fr in field_runs:
                curr_r.addnext(fr)
                curr_r = fr
            
            # If there's subsequent text in this part
            if part:
                next_r = etree.Element(f"{w}r")
                if rpr is not None:
                    next_r.append(copy.deepcopy(rpr))
                next_t = etree.SubElement(next_r, f"{w}t")
                if part.startswith(" ") or part.endswith(" "):
                    next_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                next_t.text = part
                curr_r.addnext(next_r)
                curr_r = next_r
                
    return count

tmp_zip = target_docx + ".tmp.zip"
shutil.copy2(target_docx, tmp_zip)

with zipfile.ZipFile(tmp_zip, 'r') as zin:
    files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

root = etree.fromstring(files['word/document.xml'])

# 3.1 Rename existing instrText tags
rename_map = {
    'TĐ': 'Truong_doan',
    'CVTĐ': 'Chuc_vu_TD',
    'TV': 'Thanh_vien',
    'CVTV': 'Chuc_vu_TV',
    'Phân_công_trưởng_đoàn': 'Phan_cong_TD',
    'Phân_công_thành_viên': 'Phan_cong_TV',
    'Thời_hạn_kiểm_tra_': 'Thoi_han_KT',
    'Nội_dung_kiểm_tra': 'Noi_dung_KT',
    'CCCD': 'CCCD_GD',
    'NNKD': 'Nganh_nghe',
    'CQT': 'Ten_CQT',
}

renamed_count = 0
for it in root.xpath(".//w:instrText", namespaces={"w": W_NS}):
    text = it.text
    if not text:
        continue
    # check date switches
    if 'Ngày_báo_cáo' in text:
        if '@"dd"' in text or '@"dd"' in text:
            it.text = ' MERGEFIELD Ngay_bao_cao \\@ "dd" '
            renamed_count += 1
        elif '@"M"' in text or '@"MM"' in text:
            it.text = ' MERGEFIELD Ngay_bao_cao \\@ "MM" '
            renamed_count += 1
        continue
    for old_k, new_k in rename_map.items():
        if f'"{old_k}"' in text or f' {old_k} ' in text:
            it.text = f' MERGEFIELD {new_k} '
            renamed_count += 1
            break

print(f"Renamed {renamed_count} existing instrText tags.")

# 3.2 Replace hardcoded company texts with proper field runs
c_ten = replace_text_with_field(root, 'Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải', 'Ten_NNT')
print(f"Replaced {c_ten} occurrences of Ten_NNT.")

c_qdkt = replace_text_with_field(root, '7092/QĐ-TCS3', 'So_QDKT')
print(f"Replaced {c_qdkt} occurrences of So_QDKT.")

c_qdgs = replace_text_with_field(root, '7093/QĐ-TCS3', 'So_QDGS')
print(f"Replaced {c_qdgs} occurrences of So_QDGS.")

c_gd = replace_text_with_field(root, 'Phan Thế Hanh', 'Giam_doc')
print(f"Replaced {c_gd} occurrences of Giam_doc.")

c_kt = replace_text_with_field(root, 'Lê Hoài Phương', 'Ke_toan_truong')
print(f"Replaced {c_kt} occurrences of Ke_toan_truong.")

# Branch MST check before MST
for t in root.xpath(".//w:t[contains(text(), '3603859012-001')]", namespaces={"w": W_NS}):
    t.text = t.text.replace('3603859012-001', '___BRANCH_MST___')

c_mst = replace_text_with_field(root, '3603859012', 'MST')
print(f"Replaced {c_mst} occurrences of MST.")

for t in root.xpath(".//w:t[contains(text(), '___BRANCH_MST___')]", namespaces={"w": W_NS}):
    t.text = t.text.replace('___BRANCH_MST___', '3603859012-001')

files['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='utf-8', standalone=True)

# 4. Configure settings.xml
settings_xml = files['word/settings.xml'].decode('utf-8')
settings_xml = re.sub(r'<w:mailMerge\b.*?</w:mailMerge>', '', settings_xml, flags=re.DOTALL)

mailmerge_block = (
    '<w:mailMerge xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:mainDocumentType w:val="formLetters"/>'
    '<w:linkToQuery/>'
    '<w:dataType w:val="native"/>'
    f'<w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source={excel_path};Mode=Read;Extended Properties=&quot;HDR=YES;IMEX=1;&quot;;"/>'
    '<w:query w:val="SELECT * FROM `01_DA_BAN_HANH$`"/>'
    '<w:dataSource r:id="rIdMailMerge1"/>'
    '<w:activeRecord w:val="1"/>'
    '<w:odso>'
    f'<w:udl w:val="Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source={excel_path};Mode=Read;Extended Properties=&quot;HDR=YES;IMEX=1;&quot;;"/>'
    '<w:table w:val="`01_DA_BAN_HANH$`"/>'
    '<w:src r:id="rIdMailMerge2"/>'
    '<w:colDelim w:val="9"/>'
    '<w:type w:val="database"/>'
    '<w:fHdr/>'
    '</w:odso>'
    '</w:mailMerge>'
)

settings_xml = settings_xml.replace('</w:settings>', mailmerge_block + '</w:settings>')
files['word/settings.xml'] = settings_xml.encode('utf-8')

# 5. Configure settings.xml.rels
excel_target_url = "file:///" + excel_path.replace("\\", "/").replace(" ", "%20")
rels_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rIdMailMerge1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_target_url}" TargetMode="External"/>'
    f'<Relationship Id="rIdMailMerge2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/mailMergeSource" Target="{excel_target_url}" TargetMode="External"/>'
    '</Relationships>'
)
files['word/_rels/settings.xml.rels'] = rels_xml.encode('utf-8')

# Write back zip
with zipfile.ZipFile(target_docx, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for filename, content in files.items():
        zout.writestr(filename, content)

os.remove(tmp_zip)
print(f"Successfully generated {target_docx}!")

# 6. Verify with Word COM
print("\n--- Verifying with Microsoft Word COM ---")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

try:
    doc_word = word.Documents.Open(os.path.abspath(target_docx))
    print("SUCCESS: Microsoft Word opened the document without any error!")
    print(f"Page Count: {doc_word.ComputeStatistics(2)}") # wdStatisticPages = 2
    print(f"Paragraph Count: {doc_word.Paragraphs.Count}")
    print(f"Field Count: {doc_word.Fields.Count}")
    doc_word.Close(False)
    print("Document closed cleanly.")
finally:
    word.Quit()
