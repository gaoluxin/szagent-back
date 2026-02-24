import argparse
import sys
from datetime import datetime
from app.services.excel_reader import ExcelReader
from app.services.excel_writer import ExcelWriter
from app.core.config import ENERGY_STORAGE_TEMPLATE
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def energy_storage_info_collection(input_file: str, output_file: str = None):
    try:
        logger.info(f"开始处理储能收资表: {input_file}")

        reader = ExcelReader(input_file)
        customer_data = reader.read_customer_data()
        reader.close()

        logger.info(f"读取成功: 场站名称={customer_data.station_info.name}, 子系统数量={len(customer_data.subsystems)}")

        writer = ExcelWriter(str(ENERGY_STORAGE_TEMPLATE))
        if not output_file:
            output_file = f"储能收资表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = writer.write_customer_data(customer_data, output_file)
        writer.close()

        logger.info(f"处理完成，输出文件: {output_path}")
        return output_path

    except Exception as e:
        logger.error(f"处理失败: {str(e)}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="收资工具-命令行工具")
    subparsers = parser.add_subparsers(dest="module", help="选择功能模块")

    energy_storage_parser = subparsers.add_parser("energy-storage", help="储能收资")
    energy_storage_parser.add_argument("input", help="客户版收资表文件路径")
    energy_storage_parser.add_argument("-o", "--output", help="输出文件路径（可选）")

    pv_parser = subparsers.add_parser("pv", help="光伏收资")
    pv_parser.add_argument("input", help="客户版收资表文件路径")

    args = parser.parse_args()

    if args.module == "energy-storage":
        energy_storage_info_collection(args.input, args.output)
    elif args.module == "pv":
        logger.error("光伏收资功能暂未开发")
        sys.exit(1)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
