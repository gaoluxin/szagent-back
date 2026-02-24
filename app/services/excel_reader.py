from typing import Dict, List, Optional, Any
from openpyxl import load_workbook
from app.models.schemas import CustomerData, StationInfo, MeterInfo, SubsystemInfo, ComponentInfo
import logging

logger = logging.getLogger(__name__)


class ExcelReader:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.wb = load_workbook(file_path, data_only=True)
        self.station_sheet = None

    def read_customer_data(self) -> CustomerData:
        self._find_station_sheet()
        if not self.station_sheet:
            raise ValueError("未找到'场站信息'sheet页")

        station_info = self._extract_station_info()
        meter_info = self._extract_meter_info()
        subsystems = self._extract_subsystems()
        component_data = self._extract_component_data(subsystems)

        return CustomerData(
            station_info=station_info,
            meter_info=meter_info,
            subsystems=subsystems,
            component_data=component_data
        )

    def _find_station_sheet(self):
        for sheet_name in self.wb.sheetnames:
            if sheet_name == "场站信息":
                self.station_sheet = self.wb[sheet_name]
                logger.info(f"找到场站信息sheet页: {sheet_name}")
                return
        logger.warning("未找到'场站信息'sheet页")

    def _find_section_start(self, section_name: str) -> Optional[tuple]:
        for row in self.station_sheet.iter_rows(min_row=1):
            for cell in row:
                if cell.value and section_name in str(cell.value):
                    logger.info(f"找到'{section_name}'在第{cell.row}行第{cell.column}列")
                    return (cell.row, cell.column)
        logger.warning(f"未找到'{section_name}'")
        return None

    def _extract_station_info(self) -> StationInfo:
        start_info = self._find_section_start("场站基础信息")
        if not start_info:
            raise ValueError("未找到场站基础信息")

        start_row, start_col = start_info

        data = {}
        for row in self.station_sheet.iter_rows(min_row=start_row + 1, max_col=start_col + 1):
            if not row[start_col - 1].value or str(row[start_col - 1].value).strip() == "":
                break
            key = str(row[start_col - 1].value).strip()
            value = str(row[start_col].value).strip() if row[start_col].value else ""
            data[key] = value
            logger.debug(f"场站信息: {key} = {value}")

        return StationInfo(
            name=data.get("名称", ""),
            timezone=data.get("时区", ""),
            language=data.get("语言", ""),
            address=data.get("场站地址", ""),
            rated_capacity_mwh=data.get("额定容量MWh", ""),
            rated_power_mw=data.get("额定功率MW", ""),
            longitude=data.get("经度", ""),
            latitude=data.get("纬度", ""),
            station_type=data.get("场站类型", "")
        )

    def _extract_meter_info(self) -> Optional[MeterInfo]:
        start_info = self._find_section_start("关口表信息")
        if not start_info:
            logger.info("未找到关口表信息")
            return None

        start_row, start_col = start_info

        data = {}
        for row in self.station_sheet.iter_rows(min_row=start_row + 1, max_col=start_col + 1):
            if not row[start_col - 1].value or str(row[start_col - 1].value).strip() == "":
                break
            key = str(row[start_col - 1].value).strip()
            value = str(row[start_col].value).strip() if row[start_col].value else ""
            data[key] = value
            logger.debug(f"关口表信息: {key} = {value}")

        if not data:
            return None

        return MeterInfo(
            name=data.get("名称", ""),
            meter_type=data.get("类型", ""),
            rated_capacity=data.get("额定容量", ""),
            rated_power=data.get("额定功率", "")
        )

    def _extract_subsystems(self) -> List[SubsystemInfo]:
        start_info = self._find_section_start("子系统信息")
        if not start_info:
            logger.info("未找到子系统信息")
            return []

        start_row, start_col = start_info
        subsystems = []

        header_row = start_row + 2
        headers = []
        for cell in self.station_sheet[header_row]:
            if cell.value:
                headers.append(str(cell.value).strip())

        data_start_row = header_row + 1
        for row in self.station_sheet.iter_rows(min_row=data_start_row, max_col=len(headers)):
            if not row[0].value or str(row[0].value).strip() == "":
                break

            row_data = {}
            for i, cell in enumerate(row):
                if i < len(headers):
                    row_data[headers[i]] = str(cell.value).strip() if cell.value else ""

            try:
                serial_number = int(row_data.get("序号", 0))
            except (ValueError, TypeError):
                serial_number = len(subsystems) + 1

            subsystem = SubsystemInfo(
                serial_number=serial_number,
                name=row_data.get("名称", ""),
                manufacturer=row_data.get("制造厂家", ""),
                rated_capacity=row_data.get("额定容量", ""),
                rated_power=row_data.get("额定功率", "")
            )
            subsystems.append(subsystem)
            logger.debug(f"子系统: {subsystem.name}, 制造厂家: {subsystem.manufacturer}")

        subsystems.sort(key=lambda x: x.serial_number)
        return subsystems

    def _extract_component_data(self, subsystems: List[SubsystemInfo]) -> Dict[str, Dict[str, ComponentInfo]]:
        manufacturers = set(sub.manufacturer for sub in subsystems if sub.manufacturer)
        component_data = {}

        for manufacturer in manufacturers:
            sheet_name = f"{manufacturer}_部件信息"
            if sheet_name not in self.wb.sheetnames:
                logger.warning(f"未找到部件信息sheet页: {sheet_name}")
                continue

            sheet = self.wb[sheet_name]
            component_data[manufacturer] = self._extract_components_from_sheet(sheet)

        return component_data

    def _extract_components_from_sheet(self, sheet) -> Dict[str, ComponentInfo]:
        component_types = [
            "箱变信息", "变流器信息", "电池组信息", "电池簇信息",
            "储能表信息", "风冷空调信息", "液冷空调信息", "消防设备信息"
        ]

        components = {}
        for comp_type in component_types:
            start_info = self._find_section_start_in_sheet(sheet, comp_type)
            if not start_info:
                continue

            start_row, start_col = start_info

            data = {}
            for row in sheet.iter_rows(min_row=start_row + 1, max_col=start_col + 1):
                if not row[start_col - 1].value or str(row[start_col - 1].value).strip() == "":
                    break
                key = str(row[start_col - 1].value).strip()
                value = str(row[start_col].value).strip() if row[start_col].value else ""
                data[key] = value

            components[comp_type] = ComponentInfo(component_type=comp_type, data=data)
            logger.debug(f"提取{comp_type}: {len(data)}个字段")

        return components

    def _find_section_start_in_sheet(self, sheet, section_name: str) -> Optional[tuple]:
        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                if cell.value and section_name in str(cell.value):
                    return (cell.row, cell.column)
        return None

    def get_all_sheet_names(self) -> List[str]:
        return self.wb.sheetnames

    def close(self):
        self.wb.close()
