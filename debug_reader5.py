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

print("包含'额定'的字段:")
for key in row_data.keys():
    if '额定' in key:
        print(f"  '{key}': {row_data[key]}")

print()
print("包含'功率'的字段:")
for key in row_data.keys():
    if '功率' in key:
        print(f"  '{key}': {row_data[key]}")

wb.close()
