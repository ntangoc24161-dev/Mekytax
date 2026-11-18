import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    path = os.path.abspath(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
    wb = excel.Workbooks.Open(path, False, True)
    ws = wb.Sheets("NHAP_HO_SO")
    
    # find last non-empty row in column A
    last_a = ws.Cells(ws.Rows.Count, 1).End(-4162).Row # xlUp = -4162
    print("Last row in Col A:", last_a)
    for r in range(max(1, last_a - 10), last_a + 2):
        row_vals = [ws.Cells(r, c).Text for c in range(1, 20)]
        print(f"Row {r}:", row_vals[:15])
    wb.Close(False)
finally:
    excel.Quit()
