import sys
from paddleocr import PaddleOCR

ocr = PaddleOCR(lang='vi')
print("--- PAGE 1 ---")
res = ocr.predict(r"D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an\p_1.png")
for r in res:
    if hasattr(r, 'keys'):
        print("keys:", r.keys())
        for k in ['rec_text', 'rec_texts', 'texts', 'text', 'dt_polys', 'rec_boxes']:
            if k in r:
                print(f"{k}:", r[k])
    else:
        print(r)

print("--- PAGE 2 ---")
res2 = ocr.predict(r"D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an\p_2.png")
for r in res2:
    if hasattr(r, 'keys'):
        for k in ['rec_text', 'rec_texts', 'texts', 'text']:
            if k in r:
                print(f"{k}:", r[k])
