import docx
import zipfile
import re
import shutil
from pathlib import Path

SRC = Path(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")
DST = Path(r"temp_test\TONG_HOP_test.docx")

shutil.copy2(SRC, DST)
print(f"Copied {SRC} to {DST}, size: {DST.stat().st_size}")
