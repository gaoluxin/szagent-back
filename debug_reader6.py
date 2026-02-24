from openpyxl import load_workbook

wb = load_workbook("public/储能电站收资表-客户版.xlsx", data_only=True)
sheet = wb["场站信息"]

start_row = 16
header_row = start_row + 2
data_start_row = header_row + 1

headers = []
for i, cell in enumerate(sheet[header_row]):
    if cell.value:
        headers.append(str(cell.value).strip())

row = sheet[data_start_row]
row_data = {}
for i, cell in enumerate(row):
    if i < len(headers):
        row_data[headers[i]] = str(cell.value).strip() if cell.value else ""

print("测试字段名（包含中文括号）:")
print(f"  '额定容量': {row_data.get('额定容量', 'NOT FOUND')}")
print(f"  '额定容量': {row_data.get('额定容量', 'NOT FOUND')}")
print(f"  '额定容量': {row_data.get('额定容量', 'NOT FOUND')}")
print(f"  '额定容量': {row_data.get('额定容量', 'NOT FOUND')}")
print()
print("所有字段（显示原始值）:")
for i, (key, value) in enumerate(row_data.items()):
    if i < 10:
        print(f"  {repr(key)}: {repr(value)}")

wb.close()
