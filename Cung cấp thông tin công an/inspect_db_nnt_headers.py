import xml.etree.ElementTree as ET

with open(r'I:/My Drive/DB NNT.xls', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read(100000)

pos = text.find('Tên cơ quan thuế')
r_start = text.rfind('<Row', 0, pos)
r_end = text.find('</Row>', pos) + 6
row_xml = text[r_start:r_end]

wrapper = f'<root xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">{row_xml}</root>'
root = ET.fromstring(wrapper)
ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
cells = root.findall('.//ss:Cell', ns)
col_idx = 1
for c in cells:
    if '{urn:schemas-microsoft-com:office:spreadsheet}Index' in c.attrib:
        col_idx = int(c.attrib['{urn:schemas-microsoft-com:office:spreadsheet}Index'])
    d = c.find('ss:Data', ns)
    txt = d.text if d is not None else ''
    print(f'Col {col_idx}: {txt}')
    col_idx += 1
