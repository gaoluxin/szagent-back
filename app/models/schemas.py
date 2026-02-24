from typing import Dict, List, Optional, Any
from pydantic import BaseModel


class StationInfo(BaseModel):
    name: str
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
    rated_capacity: str
    rated_power: str


class ComponentInfo(BaseModel):
    component_type: str
    data: Dict[str, str]


class CustomerData(BaseModel):
    station_info: StationInfo
    meter_info: Optional[MeterInfo] = None
    subsystems: List[SubsystemInfo] = []
    component_data: Dict[str, Dict[str, ComponentInfo]] = {}
