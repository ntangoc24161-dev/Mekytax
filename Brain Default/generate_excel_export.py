import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json, os, datetime

# Target file paths
target_dir = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026. CUNG CAP CONG AN NAM 2026\CV MOI'
os.makedirs(target_dir, exist_ok=True)

excel_file_path1 = os.path.join(target_dir, 'Danh_sach_tong_hop_ho_so_cung_cap_CV_MOI.xlsx')
excel_file_path2 = os.path.join(target_dir, 'CV MOI.xlsx')
excel_file_path3 = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026. CUNG CAP CONG AN NAM 2026\CV MOI.xlsx'

print('Target paths prepared')
