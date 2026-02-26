from app.services.excel_reader import ExcelReader

reader = ExcelReader("public/储能电站收资表-客户版.xlsx")
customer_data = reader.read_customer_data()

print("子系统舱数量:")
print()

for subsystem in customer_data.subsystems[:5]:
    print(f"子系统: {subsystem.name}")
    print(f"  舱数量: '{subsystem.cabin_count}'")
    print(f"  舱数量类型: {type(subsystem.cabin_count)}")
    try:
        cabin_count = int(subsystem.cabin_count)
        print(f"  舱数量(int): {cabin_count}")
    except (ValueError, TypeError) as e:
        print(f"  舱数量转换失败: {e}")
    print()

reader.close()
