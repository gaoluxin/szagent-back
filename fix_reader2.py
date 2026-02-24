file_path = "app/services/excel_reader.py"

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'rated_capacity=row_data.get("额定容量", "")' in line:
        lines[i] = line.replace('rated_capacity=row_data.get("额定容量", "")', 'rated_capacity=row_data.get("额定容量", "")')
    elif 'rated_power=row_data.get("额定功率", "")' in line:
        lines[i] = line.replace('rated_power=row_data.get("额定功率", "")', 'rated_power=row_data.get("额定功率", "")')

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("修复完成")
