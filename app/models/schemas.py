from typing import Dict, List, Optional, Any
from pydantic import BaseModel


class StationInfo(BaseModel):
    name: str
    short_name: str = ""
    timezone: str
    language: str
    address: str
    rated_capacity_mwh: str
    rated_power_mw: str
    longitude: str
    latitude: str
    station_type: str


class MeterInfo(BaseModel):
    name: str
    meter_type: str
    rated_capacity: str
    rated_power: str


class SubsystemInfo(BaseModel):
    serial_number: int
    name: str
    manufacturer: str
    model: str = ""
    rated_capacity: str
    rated_power: str
    equipment_structure: str = ""
    incoming_line_name: str = ""
    transformer_count: str = ""
    pcs_count: str = ""
    battery_bank_count: str = ""
    battery_cluster_count: str = ""
    energy_meter_count: str = ""
    auxiliary_meter_count: str = ""
    air_conditioner_structure: str = ""
    air_cooler_count: str = ""
    liquid_cooler_count: str = ""
    fire_suppression_structure: str = ""
    fire_host_count: str = ""
    fire_detector_count: str = ""
    fire_suppressor_count: str = ""


class ComponentInfo(BaseModel):
    component_type: str
    data: Dict[str, str]
    box_transformer_type: str = ""
    cooling_system_type: str = ""


class CustomerData(BaseModel):
    station_info: StationInfo
    meter_info: Optional[MeterInfo] = None
    subsystems: List[SubsystemInfo] = []
    component_data: Dict[str, Dict[str, ComponentInfo]] = {}
