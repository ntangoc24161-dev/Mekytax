import xml.etree.ElementTree as ET
from pathlib import Path
import json
import openpyxl

base_dir = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505")
xml_dir = base_dir / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"

def try_num(s):
    if s is None:
        return 0
    s_str = str(s).strip()
    if not s_str:
        return 0
    # check if integer or float
    clean = s_str.replace(',', '').strip()
    try:
        if '.' in clean:
            return float(clean)
        return int(clean)
    except:
        return s_str

def parse_xml_files():
    data = {
        'BCTC': {},
        'GTGT': {},
        'TNDN': {},
        'TNCN': {}
    }
    
    for f in xml_dir.glob('*.xml'):
        try:
            tree = ET.parse(f)
            root = tree.getroot()
            
            # Find general info
            ten_tkhai = ''
            ky_kkhai = ''
            nam_kkhai = ''
            loai_tkhai = ''
            so_lan = '0'
            for e in root.iter():
                t = e.tag.split('}')[-1]
                if t in ['tenTKhai', 'tenToKhai']: ten_tkhai = e.text or ''
                elif t in ['kyKKhai']: ky_kkhai = e.text or ''
                elif t in ['namKKhai', 'namBCTC']: nam_kkhai = e.text or ''
                elif t in ['loaiTKhai']: loai_tkhai = e.text or ''
                elif t in ['soLan']: so_lan = e.text or '0'
            
            # 1. BCTC
            if 'BCTC' in f.name or 'B01A' in f.name:
                yr = ky_kkhai if ky_kkhai in ['2022', '2023', '2024', '2025'] else nam_kkhai
                if not yr:
                    for y in ['2022', '2023', '2024', '2025']:
                        if y in f.name: yr = y; break
                
                bctc_item = {'file': f.name, 'CDKT': {}, 'KQKD': {}, 'CDTK': {}, 'LCTT': {}}
                
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1]
                    if tag == 'CTieuTKhaiChinh':
                        for child in elem:
                            c_tag = child.tag.split('}')[-1]
                            d = {leaf.tag.split('}')[-1]: try_num(leaf.text) for leaf in child}
                            bctc_item['CDKT'][c_tag] = d
                    elif tag == 'PL_KQHDSXKD':
                        for child in elem:
                            c_tag = child.tag.split('}')[-1]
                            d = {leaf.tag.split('}')[-1]: try_num(leaf.text) for leaf in child}
                            bctc_item['KQKD'][c_tag] = d
                    elif tag == 'PL_CDTK':
                        for child in elem:
                            c_tag = child.tag.split('}')[-1]
                            d = {}
                            for sub in child:
                                s_tag = sub.tag.split('}')[-1]
                                sub_d = {leaf.tag.split('}')[-1]: try_num(leaf.text) for leaf in sub}
                                d[s_tag] = sub_d
                            bctc_item['CDTK'][c_tag] = d
                data['BCTC'][yr] = bctc_item
                
            # 2. TNDN
            elif 'TNDN' in f.name or '03_TNDN' in f.name:
                yr = ky_kkhai if ky_kkhai in ['2022', '2023', '2024', '2025'] else nam_kkhai
                if not yr:
                    for y in ['2022', '2023', '2024', '2025']:
                        if y in f.name: yr = y; break
                tndn_dict = {'file': f.name, 'loai': loai_tkhai, 'soLan': so_lan}
                for elem in root.iter():
                    t = elem.tag.split('}')[-1]
                    if elem.text and elem.text.strip():
                        tndn_dict[t] = try_num(elem.text)
                if yr not in data['TNDN'] or loai_tkhai == 'C' or int(so_lan) > int(data['TNDN'][yr].get('soLan', 0)):
                    data['TNDN'][yr] = tndn_dict
                    
            # 3. GTGT
            elif 'GTGT' in f.name or '01_GTGT' in f.name:
                ky_str = ''
                for q in ['Q1', 'Q2', 'Q3', 'Q4']:
                    for y in ['2022', '2023', '2024', '2025']:
                        if f"{q}.{y}" in f.name or f"{q}{y}" in f.name or f"{q}_{y}" in f.name:
                            ky_str = f"{q}/{y}"
                            break
                        if f"{q}." in f.name and y in f.name:
                            ky_str = f"{q}/{y}"
                            break
                        if f"{q} " in f.name and y in f.name:
                            ky_str = f"{q}/{y}"
                            break
                if not ky_str:
                    tu_ngay = ''
                    for elem in root.iter():
                        t = elem.tag.split('}')[-1]
                        if t == 'kyKKhaiTuNgay' and elem.text:
                            tu_ngay = elem.text.strip()
                    if '01/01/' in tu_ngay: ky_str = f"Q1/{tu_ngay[-4:]}"
                    elif '01/04/' in tu_ngay: ky_str = f"Q2/{tu_ngay[-4:]}"
                    elif '01/07/' in tu_ngay: ky_str = f"Q3/{tu_ngay[-4:]}"
                    elif '01/10/' in tu_ngay: ky_str = f"Q4/{tu_ngay[-4:]}"
                
                gtgt_dict = {'file': f.name, 'ky': ky_str, 'loai': loai_tkhai, 'soLan': so_lan}
                for elem in root.iter():
                    t = elem.tag.split('}')[-1]
                    if elem.text and elem.text.strip():
                        gtgt_dict[t] = try_num(elem.text)
                if ky_str:
                    if ky_str not in data['GTGT'] or loai_tkhai == 'C' or int(so_lan) > int(data['GTGT'][ky_str].get('soLan', 0)):
                        data['GTGT'][ky_str] = gtgt_dict

            # 4. TNCN
            elif 'TNCN' in f.name:
                is_qtt = False
                yr = None
                for y in ['2022', '2023', '2024', '2025']:
                    if f"{y} chinh thuc" in f.name or f"Y{y}" in f.name or f.name == 'QTT-TNCN.xml':
                        is_qtt = True
                        if f.name == 'QTT-TNCN.xml': yr = '2025'
                        else: yr = y
                        break
                if is_qtt and yr:
                    tncn_dict = {'file': f.name, 'nam': yr}
                    for elem in root.iter():
                        t = elem.tag.split('}')[-1]
                        if elem.text and elem.text.strip():
                            tncn_dict[t] = try_num(elem.text)
                    data['TNCN'][yr] = tncn_dict
        except Exception as ex:
            print(f"Error parsing {f.name}: {ex}")
            
    return data

parsed_data = parse_xml_files()
print("Parsed XML data summary:")
print("BCTC years:", sorted(list(parsed_data['BCTC'].keys())))
print("TNDN years:", sorted(list(parsed_data['TNDN'].keys())))
print("GTGT periods:", sorted(list(parsed_data['GTGT'].keys())))
print("TNCN QTT years:", sorted(list(parsed_data['TNCN'].keys())))

with open('scratch_all_parsed_xml.json', 'w', encoding='utf-8') as f_out:
    json.dump(parsed_data, f_out, ensure_ascii=False, indent=2)
