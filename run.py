#!/usr/bin/env python
"""
收资工具-后台
模块总入口
"""

import argparse
import sys
import uvicorn
from app.cli.main import main as cli_main


def start_server(host: str = "0.0.0.0", port: int = 8000, reload: bool = False):
    uvicorn.run(
        "main:app",
        host=host,
        port=port,
        reload=reload
    )


def main():
    parser = argparse.ArgumentParser(
        description="收资工具-后台 - 模块总入口",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  启动Web服务:
    python run.py server
    python run.py server --host 127.0.0.1 --port 8080
    python run.py server --reload

  使用命令行工具:
    python run.py cli energy-storage 客户收资表.xlsx
    python run.py cli energy-storage 客户收资表.xlsx -o 输出文件.xlsx
    python run.py cli pv 客户收资表.xlsx
        """
    )

    subparsers = parser.add_subparsers(dest="mode", help="运行模式")

    server_parser = subparsers.add_parser("server", help="启动Web API服务")
    server_parser.add_argument("--host", default="0.0.0.0", help="监听地址（默认: 0.0.0.0）")
    server_parser.add_argument("--port", type=int, default=8000, help="监听端口（默认: 8000）")
    server_parser.add_argument("--reload", action="store_true", help="启用热重载（开发模式）")

    cli_parser = subparsers.add_parser("cli", help="使用命令行工具")
    cli_parser.add_argument("module", choices=["energy-storage", "pv"], help="功能模块")
    cli_parser.add_argument("input", help="客户版收资表文件路径")
    cli_parser.add_argument("-o", "--output", help="输出文件路径（可选）")

    args = parser.parse_args()

    if args.mode == "server":
        print(f"启动收资工具-后台服务...")
        print(f"访问地址: http://{args.host}:{args.port}")
        print(f"API文档: http://{args.host}:{args.port}/docs")
        start_server(args.host, args.port, args.reload)
    elif args.mode == "cli":
        sys.argv = ["run.py", args.module, args.input]
        if args.output:
            sys.argv.extend(["-o", args.output])
        cli_main()
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
