import zipfile

with zipfile.ZipFile(r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm') as z:
    bin_data = z.read('xl/vbaProject.bin')

# Dump all printable text
import string
printable = set(bytes(string.printable, 'ascii'))
res = []
cur = bytearray()
for b in bin_data:
    if b in printable and b not in b'\r\n\t':
        cur.append(b)
    else:
        if len(cur) >= 4:
            res.append(cur.decode('ascii', errors='ignore'))
        cur = bytearray()

for s in res:
    if any(k in s.lower() for k in ['target', 'range', 'cells', 'sheet', 'row', 'column', 'copy', 'paste', 'ban_hanh', 'thuc_hien']):
        print(s)
