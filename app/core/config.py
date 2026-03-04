from pathlib import Path
import sys


def _get_runtime_base_dir() -> Path:
    """
    获取运行时可读取资源的根目录：
    - 开发环境：项目根目录（szagent-back）
    - PyInstaller 打包后：临时解压目录 sys._MEIPASS
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent.parent


def _get_exec_base_dir() -> Path:
    """
    获取可写输出目录的根目录：
    - 开发环境：项目根目录（szagent-back）
    - PyInstaller 打包后：main.exe 所在目录
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent.parent


RUNTIME_BASE_DIR = _get_runtime_base_dir()
EXEC_BASE_DIR = _get_exec_base_dir()

# 所有需要外部可编辑/替换的资源，都从可执行文件所在目录读取
PUBLIC_DIR = EXEC_BASE_DIR / "public"
TEMPLATE_DIR = PUBLIC_DIR / "储模板"

# 输出目录同样放在可执行文件所在目录下，便于用户查找
OUTPUT_DIR = EXEC_BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

CUSTOMER_TEMPLATE = PUBLIC_DIR / "储能电站收资表-客户版.xlsx"
ENERGY_STORAGE_TEMPLATE = TEMPLATE_DIR / "储能收资模板.xlsx"
POINT_TABLE_TEMPLATE = TEMPLATE_DIR / "储能点表模板.xlsx"
