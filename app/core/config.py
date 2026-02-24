from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent.parent
PUBLIC_DIR = BASE_DIR / "public"
TEMPLATE_DIR = PUBLIC_DIR / "储模板"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CUSTOMER_TEMPLATE = PUBLIC_DIR / "储能电站收资表-客户版.xlsx"
ENERGY_STORAGE_TEMPLATE = TEMPLATE_DIR / "储能收资模板.xlsx"
POINT_TABLE_TEMPLATE = TEMPLATE_DIR / "储能点表模板.xlsx"
