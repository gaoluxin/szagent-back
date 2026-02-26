from app.services.excel_reader import ExcelReader
from app.services.excel_writer import ExcelWriter

reader = ExcelReader("public/储能电站收资表-客户版.xlsx")
customer_data = reader.read_customer_data()

writer = ExcelWriter("public/储模板/储能收资模板.xlsx", customer_data)
writer.write_all()
writer.close()

print("舱信息:")
print()

for manufacturer, components in customer_data.component_data.items():
    if "舱信息" in components:
        print(f"制造厂家: {manufacturer}")
        cabin_info = components["舱信息"]
        print(f"  cabin_model: '{cabin_info.cabin_model}'")
        print(f"  cabin_manufacturer: '{cabin_info.cabin_manufacturer}'")

reader.close()
