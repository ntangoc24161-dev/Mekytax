import zipfile

with zipfile.ZipFile(r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm') as z:
    bin_data = z.read('xl/vbaProject.bin')

# Let's extract strings or text from vbaProject.bin
import re
# Find all printable ASCII / UTF-8 strings longer than 4 chars
strings = re.findall(rb'[\x20-\x7e\r\n]{4,}', bin_data)
for s in strings:
    s_dec = s.decode('latin1', errors='ignore')
    if any(k in s_dec for k in ['Sub ', 'Function ', '01_DA_BAN_HANH', '02_DANG_THUC_HIEN', 'Worksheet_Change']):
        print(s_dec[:200])
