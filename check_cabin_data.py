from openpyxl import load_workbook

wb = load_workbook("output/储能收资表_20260225_152331.xlsx", data_only=True)
sheet = wb["5-舱"]

print("5-舱 sheet页第5-25行数据:")
print()

for row_idx in range(5, 26):
    print(f"第{row_idx}行:")
    for i, cell in enumerate(sheet[row_idx]):
        if cell.value:
            print(f"  列{i+1}: {cell.value}")
    print()

wb.close()
