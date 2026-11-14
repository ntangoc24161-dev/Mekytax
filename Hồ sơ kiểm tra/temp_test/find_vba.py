import zipfile
import re

with zipfile.ZipFile(r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm') as z:
    bin_data = z.read('xl/vbaProject.bin')

# Search for DocSo or text
print("Length of vbaProject.bin:", len(bin_data))
# let's look for any occurrences of 'DocSo'
idx = 0
while True:
    idx = bin_data.find(b'DocSo', idx)
    if idx == -1:
        break
    print("Found DocSo at offset:", idx)
    start = max(0, idx - 50)
    end = min(len(bin_data), idx + 200)
    print("Context bytes:", bin_data[start:end])
    idx += 5
