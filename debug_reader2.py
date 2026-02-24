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

print(f"表头字段 ({len(headers)}个):")
for i, h in enumerate(headers):
    print(f"  {i}: {h}")
print()

row = sheet[data_start_row]
row_data = {}
for i, cell in enumerate(row):
    if i < len(headers):
        row_data[headers[i]] = str(cell.value).strip() if cell.value else ""

print(f"row_data中的字段 ({len(row_data)}个):")
for key, value in row_data.items():
    print(f"  {key}: {value}")
print()

print("测试字段读取:")
print(f"  row_data.get('额定容量', ''): {row_data.get('额定容量', '')}")
print(f"  row_data.get('额定功率', ''): {row_data.get('额定功率', '')}")
print(f"  row_data.get('型号', ''): {row_data.get('型号', '')}")
print(f"  row_data.get('储能系统名称', ''): {row_data.get('储能系统名称', '')}")

wb.close()
