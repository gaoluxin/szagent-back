from openpyxl import load_workbook

wb = load_workbook("output/储能收资表_20260225_165905.xlsx", data_only=True)
sheet = wb["6-电池组"]

print("6-电池组 sheet页前10行数据:")
print()

for row_idx in range(3, 13):
    print(f"第{row_idx}行:")
    for i, cell in enumerate(sheet[row_idx]):
        if cell.value:
            print(f"  列{i+1}: {cell.value}")
    print()

wb.close()
