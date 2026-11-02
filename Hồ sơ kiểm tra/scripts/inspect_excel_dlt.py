import openpyxl

wb = openpyxl.load_workbook(r"output\3603859012_HO-SO-VE_son-tien-hai\du-lieu-tron_son-tien-hai_3603859012.xlsx")
for sheet in wb.sheetnames:
    print(f"=== SHEET: {sheet} ===")
    ws = wb[sheet]
    for r in range(1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(row_vals):
            print(f"r{r}: {row_vals}")
