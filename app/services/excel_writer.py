from typing import Dict, List
from openpyxl import load_workbook
from app.models.schemas import CustomerData
from app.core.config import OUTPUT_DIR
import logging

logger = logging.getLogger(__name__)


class ExcelWriter:
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.wb = load_workbook(template_path)
        self.sheets = {
            "1-场站": None,
            "1.1 物理场站": None,
            "2-储能系统": None,
            "3-箱变": None,
            "4-变流器": None,
            "5-舱": None,
            "6-电池组": None,
            "7-电池簇": None,
            "8-空调": None,
            "9-电表": None,
            "10-消防设备": None,
            "11-其他设备": None
        }
        self._load_sheets()

    def _load_sheets(self):
        for sheet_name in self.sheets.keys():
            if sheet_name in self.wb.sheetnames:
                self.sheets[sheet_name] = self.wb[sheet_name]
                logger.info(f"加载sheet页: {sheet_name}")
            else:
                logger.warning(f"模板中未找到sheet页: {sheet_name}")

    def write_customer_data(self, customer_data: CustomerData, output_filename: str) -> str:
        self._write_station_sheet(customer_data)
        self._write_physical_station_sheet(customer_data)
        self._write_energy_storage_system_sheet(customer_data)

        output_path = OUTPUT_DIR / output_filename
        self.wb.save(output_path)
        logger.info(f"文件已保存到: {output_path}")
        return str(output_path)

    def _write_station_sheet(self, customer_data: CustomerData):
        sheet = self.sheets.get("1-场站")
        if not sheet:
            logger.warning("未找到1-场站sheet页")
            return

        station = customer_data.station_info
        subsystem_count = len(customer_data.subsystems)

        mapping = {
            "名称*": station.name,
            "时区*": station.timezone,
            "语言*": station.language,
            "场站详细地址*": station.address,
            "场站额定容量*": station.rated_capacity_mwh,
            "场站额定功率*": station.rated_power_mw,
            "经度*": station.longitude,
            "纬度*": station.latitude,
            "储能系统数量*": str(subsystem_count),
            "所属物理场站*": f"{station.name}物理场站",
            "场站类型*": station.station_type,
            "Scada别名": "",
            "升压等级": "",
            "电网线路名称": ""
        }

        header_row = 3
        headers = []
        for cell in sheet[header_row]:
            if cell.value:
                headers.append(str(cell.value).strip())

        for col_idx, header in enumerate(headers, start=1):
            if header in mapping:
                sheet.cell(row=4, column=col_idx, value=mapping[header])
                logger.debug(f"写入{header}: {mapping[header]}")

    def _write_physical_station_sheet(self, customer_data: CustomerData):
        sheet = self.sheets.get("1.1 物理场站")
        if not sheet:
            logger.warning("未找到1.1 物理场站sheet页")
            return

        station = customer_data.station_info
        physical_station_name = f"{station.name}物理场站"

        header_row = 3
        headers = []
        for cell in sheet[header_row]:
            if cell.value:
                headers.append(str(cell.value).strip())

        data_row = 4
        for col_idx, header in enumerate(headers, start=1):
            if header == "名称*":
                sheet.cell(row=data_row, column=col_idx, value=physical_station_name)
                logger.debug(f"写入物理场站名称: {physical_station_name}")
            elif header == "场站详细地址*":
                sheet.cell(row=data_row, column=col_idx, value=station.address)
            elif header == "经度*":
                sheet.cell(row=data_row, column=col_idx, value=station.longitude)
            elif header == "纬度*":
                sheet.cell(row=data_row, column=col_idx, value=station.latitude)

    def _write_energy_storage_system_sheet(self, customer_data: CustomerData):
        sheet = self.sheets.get("2-储能系统")
        if not sheet:
            logger.warning("未找到2-储能系统sheet页")
            return

        header_row = 3
        headers = []
        for cell in sheet[header_row]:
            if cell.value:
                headers.append(str(cell.value).strip())

        data_start_row = 4
        for subsystem in customer_data.subsystems:
            pcs_connection_type = self._get_pcs_connection_type(subsystem.equipment_structure)
            system_composition = self._get_system_composition(subsystem)

            mapping = {
                "名称*": subsystem.name,
                "制造厂家*": subsystem.manufacturer,
                "型号*": subsystem.model,
                "额定容量*": subsystem.rated_capacity,
                "额定功率*": subsystem.rated_power,
                "所属升压站进线名称*": subsystem.incoming_line_name,
                "PCS连接形式*": pcs_connection_type,
                "序号*": str(subsystem.serial_number),
                "Scada别名": "",
                "已接入系统构成*": system_composition,
                "模型ID": ""
            }

            for col_idx, header in enumerate(headers, start=1):
                if header in mapping:
                    sheet.cell(row=data_start_row, column=col_idx, value=mapping[header])
                    logger.debug(f"写入{header}: {mapping[header]}")

            data_start_row += 1

    def _get_pcs_connection_type(self, equipment_structure: str) -> str:
        structure_map = {
            "系统模式1": "集中式PCS",
            "系统模式2": "组串式PCS",
            "系统模式3": "分散式"
        }
        return structure_map.get(equipment_structure, "")

    def _get_system_composition(self, subsystem) -> str:
        composition_parts = ["ESSVIEW"]

        if subsystem.battery_cluster_count and subsystem.battery_cluster_count.strip() and subsystem.battery_cluster_count.strip() != "0":
            composition_parts.append("DC")

        if subsystem.pcs_count and subsystem.pcs_count.strip() and subsystem.pcs_count.strip() != "0":
            composition_parts.append("AC")

        has_air_cooler = subsystem.air_cooler_count and subsystem.air_cooler_count.strip() and subsystem.air_cooler_count.strip() != "0"
        has_liquid_cooler = subsystem.liquid_cooler_count and subsystem.liquid_cooler_count.strip() and subsystem.liquid_cooler_count.strip() != "0"
        if has_air_cooler or has_liquid_cooler:
            composition_parts.append("ThermalSystem")

        has_fire_host = subsystem.fire_host_count and subsystem.fire_host_count.strip() and subsystem.fire_host_count.strip() != "0"
        has_fire_detector = subsystem.fire_detector_count and subsystem.fire_detector_count.strip() and subsystem.fire_detector_count.strip() != "0"
        has_fire_suppressor = subsystem.fire_suppressor_count and subsystem.fire_suppressor_count.strip() and subsystem.fire_suppressor_count.strip() != "0"
        if has_fire_host or has_fire_detector or has_fire_suppressor:
            composition_parts.append("FireSuppressionSystem")

        if subsystem.battery_bank_count and subsystem.battery_bank_count.strip() and subsystem.battery_bank_count.strip() != "0":
            composition_parts.append("BatteryBankView")

        return ",".join(composition_parts)

    def close(self):
        self.wb.close()
