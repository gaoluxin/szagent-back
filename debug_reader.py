from openpyxl import load_workbook

wb = load_workbook("public/储能电站收资表-客户版.xlsx", data_only=True)
sheet = wb["场站信息"]

start_row = 16
header_row = start_row + 2

print(f"子系统信息从第{start_row}行开始")
print(f"表头在第{header_row}行")
print()

headers = []
for i, cell in enumerate(sheet[header_row]):
    if cell.value:
        headers.append(str(cell.value).strip())
        print(f"列{i+1}: {cell.value}")

print()
print("第一行数据:")
data_start_row = header_row + 1
row = sheet[data_start_row]
for i, cell in enumerate(row):
    if i < len(headers):
        print(f"{headers[i]}: {cell.value} (类型: {type(cell.value).__name__})")

wb.close()
