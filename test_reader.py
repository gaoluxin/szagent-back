from app.services.excel_reader import ExcelReader

reader = ExcelReader("public/储能电站收资表-客户版.xlsx")
customer_data = reader.read_customer_data()

print(f"场站名称: {customer_data.station_info.name}")
print(f"子系统数量: {len(customer_data.subsystems)}")
print()

for i, subsystem in enumerate(customer_data.subsystems[:3]):
    print(f"子系统 {i+1}:")
    print(f"  名称: {subsystem.name}")
    print(f"  制造厂家: {subsystem.manufacturer}")
    print(f"  型号: {subsystem.model}")
    print(f"  额定容量: {subsystem.rated_capacity}")
    print(f"  额定功率: {subsystem.rated_power}")
    print()

reader.close()
