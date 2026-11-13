import zipfile
import re

with zipfile.ZipFile(r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm') as z:
    bin_data = z.read('xl/vbaProject.bin')

# MS-OVBA decompression function
def decompress_vba(compressed_container):
    if not compressed_container or compressed_container[0] != 0x01:
        return b""
    decompressed = bytearray()
    pos = 1
    sig = compressed_container[pos:pos+2]
    # signature check omitted or standard
    pos = 3
    while pos < len(compressed_container):
        header = int.from_bytes(compressed_container[pos:pos+2], 'little')
        pos += 2
        chunk_size = (header & 0x0FFF) + 3
        chunk_compressed = (header & 0x8000) != 0
        chunk_end = pos + chunk_size - 2
        chunk_data = compressed_container[pos:chunk_end]
        pos = chunk_end
        
        if not chunk_compressed:
            decompressed.extend(chunk_data)
        else:
            c_pos = 0
            while c_pos < len(chunk_data):
                flag_byte = chunk_data[c_pos]
                c_pos += 1
                for bit in range(8):
                    if c_pos >= len(chunk_data):
                        break
                    if not (flag_byte & (1 << bit)):
                        decompressed.append(chunk_data[c_pos])
                        c_pos += 1
                    else:
                        if c_pos + 1 >= len(chunk_data):
                            break
                        token = int.from_bytes(chunk_data[c_pos:c_pos+2], 'little')
                        c_pos += 2
                        # compute length and offset
                        d_len = len(decompressed)
                        # mask calculation
                        bit_count = 4
                        while (1 << bit_count) <= d_len and bit_count < 12:
                            bit_count += 1
                        length_mask = 0xFFFF >> bit_count
                        offset_mask = ~length_mask & 0xFFFF
                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bit_count)) + 1
                        start_copy = len(decompressed) - offset
                        for k in range(length):
                            decompressed.append(decompressed[start_copy + k])
    return bytes(decompressed)

# Let's search for 'Attribut' in bin_data
for m in re.finditer(b'Attribut', bin_data):
    start = max(0, m.start() - 20)
    chunk = bin_data[start:start+4000]
    dec = decompress_vba(chunk)
    if dec:
        print("Decompressed chunk:")
        print(dec.decode('utf-8', errors='ignore')[:500])
        print("="*40)
