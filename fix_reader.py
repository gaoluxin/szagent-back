import re

file_path = "app/services/excel_reader.py"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace(
    'rated_capacity=row_data.get("额定容量", "")',
    'rated_capacity=row_data.get("额定容量", "")'
)
content = content.replace(
    'rated_power=row_data.get("额定功率", "")',
    'rated_power=row_data.get("额定功率", "")'
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("修复完成")
