import xml.etree.ElementTree as ET
from pathlib import Path
import json
import openpyxl

xml_dir = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\1. DU LIEU THO NGANH THUE\1.1. To khai XML")

def parse_bctc(fpath):
    tree = ET.parse(fpath)
    root = tree.getroot()
    res = {'CDKT': {}, 'KQKD': {}, 'CDTK': {}, 'LCTT': {}}
    for elem in root.iter():
        tag = elem.tag.split('}')[-1]
        if tag == 'CTieuTKhaiChinh':
            for child in elem:
                c_tag = child.tag.split('}')[-1]
                if c_tag in ['SoCuoiNam', 'SoDauNam']:
                    d = {}
                    for leaf in child:
                        lt = leaf.tag.split('}')[-1]
                        if leaf.text and leaf.text.strip():
                            try:
                                v = int(leaf.text.strip()) if '.' not in leaf.text.strip() else float(leaf.text.strip())
                            except:
                                v = leaf.text.strip()
                            d[lt] = v
                    res['CDKT'][c_tag] = d
        elif tag == 'PL_KQHDSXKD':
            for child in elem:
                c_tag = child.tag.split('}')[-1]
                if c_tag in ['NamNay', 'NamTruoc']:
                    d = {}
                    for leaf in child:
                        lt = leaf.tag.split('}')[-1]
                        if leaf.text and leaf.text.strip():
                            try:
                                v = int(leaf.text.strip()) if '.' not in leaf.text.strip() else float(leaf.text.strip())
                            except:
                                v = leaf.text.strip()
                            d[lt] = v
                    res['KQKD'][c_tag] = d
        elif tag == 'PL_CDTK':
            for child in elem:
                c_tag = child.tag.split('}')[-1]
                d = {}
                for sub in child:
                    s_tag = sub.tag.split('}')[-1]
                    sub_d = {}
                    for leaf in sub:
                        lt = leaf.tag.split('}')[-1]
                        if leaf.text and leaf.text.strip():
                            try:
                                v = int(leaf.text.strip()) if '.' not in leaf.text.strip() else float(leaf.text.strip())
                            except:
                                v = leaf.text.strip()
                            sub_d[lt] = v
                    d[s_tag] = sub_d
                res['CDTK'][c_tag] = d
    return res

years = {
    '2022': list(xml_dir.glob('*BCTC*2022*'))[0],
    '2023': list(xml_dir.glob('*BCTC*2023*'))[0],
    '2024': list(xml_dir.glob('*BCTC*2024*'))[0],
    '2025': list(xml_dir.glob('*BCTC*Y2025*'))[0]
}

all_bctc = {}
for yr, f in years.items():
    all_bctc[yr] = parse_bctc(f)

print("=== BCTC EXTRACTED NUMBERS ===")
for yr in ['2022', '2023', '2024', '2025']:
    print(f"\n--- YEAR {yr} ---")
    kq = all_bctc[yr]['KQKD'].get('NamNay', {})
    cd = all_bctc[yr]['CDKT'].get('SoCuoiNam', {})
    cdtk_no = all_bctc[yr]['CDTK'].get('SoDuCuoiKy', {}).get('No', {})
    cdtk_co = all_bctc[yr]['CDTK'].get('SoDuCuoiKy', {}).get('Co', {})
    print(f"  Doanh thu [01/10]: {kq.get('ct01') or kq.get('ct10')}")
    print(f"  Giá vốn [11]: {kq.get('ct11')}")
    print(f"  Lợi nhuận gộp [20]: {kq.get('ct20')}")
    print(f"  Chi phí QLDN [24/25]: {kq.get('ct24') or kq.get('ct25')}")
    print(f"  Chi phí tài chính / Lãi vay [22/23]: {kq.get('ct22')} / {kq.get('ct23')}")
    print(f"  Lợi nhuận kế toán trước thuế [50]: {kq.get('ct50')}")
    print(f"  Chi phí thuế TNDN [51]: {kq.get('ct51')}")
    print(f"  Lợi nhuận sau thuế [60]: {kq.get('ct60')}")
    print(f"  Tiền và TĐT [110]: {cd.get('ct110')} (TK 111={cdtk_no.get('ct111')}, TK 112={cdtk_no.get('ct112')})")
    print(f"  Phải thu KH [131]: {cd.get('ct131')} (TK 131={cdtk_no.get('ct131')})")
    print(f"  Hàng tồn kho [140/141]: {cd.get('ct140') or cd.get('ct141')} (TK 156={cdtk_no.get('ct156')})")
    print(f"  Thuế GTGT khấu trừ [152]: {cd.get('ct152')} (TK 133={cdtk_no.get('ct133')})")
    print(f"  TSCĐ [211] / Hao mòn [214]: {cd.get('ct211')} / {cd.get('ct214')}")
    print(f"  Chi phí trả trước [242]: {cd.get('ct242')} (TK 242={cdtk_no.get('ct242')})")
    print(f"  Nợ phải trả [300]: {cd.get('ct300')}")
    print(f"  Phải trả người bán [311]: {cd.get('ct311')} (TK 331={cdtk_co.get('ct331')})")
    print(f"  Người mua trả tiền trước [312]: {cd.get('ct312')}")
    print(f"  Thuế phải nộp [313]: {cd.get('ct313')} (TK 333={cdtk_co.get('ct333')})")
    print(f"  Phải trả NLĐ [314]: {cd.get('ct314')} (TK 334={cdtk_co.get('ct334')})")
    print(f"  Chi phí phải trả trích trước [315]: {cd.get('ct315')} (TK 335={cdtk_co.get('ct335')})")
    print(f"  Doanh thu chưa thực hiện [318]: {cd.get('ct318')}")
    print(f"  Vay ngắn hạn [320]: {cd.get('ct320')}")
    print(f"  Vốn chủ sở hữu [400] / Vốn góp [411]: {cd.get('ct400')} / {cd.get('ct411')}")
    print(f"  Lợi nhuận chưa phân phối [421]: {cd.get('ct421')} (TK 421={cdtk_co.get('ct421')})")

with open('scratch_bctc_extracted.json', 'w', encoding='utf-8') as f_out:
    json.dump(all_bctc, f_out, ensure_ascii=False, indent=2)
