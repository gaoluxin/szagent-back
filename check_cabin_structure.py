from openpyxl import load_workbook

wb = load_workbook("public/储能电站收资表-客户版.xlsx", data_only=True)
sheet = wb["海博_部件信息"]

print("海博_部件信息 sheet页第1-10行（显示所有列）:")
print()

for row_idx in range(1, 11):
    print(f"第{row_idx}行:")
    for i, cell in enumerate(sheet[row_idx]):
        if cell.value and i <= 6:
            print(f"  列{i+1} (索引{i}): {cell.value}")
    print()

wb.close()
