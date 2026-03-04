import os
import sys

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from app.api.routers import energy_storage, pv

app = FastAPI(
    title="收资工具-后台",
    description="提供光伏收资和储能收资功能的API接口",
    version="1.0.0",
)


def get_static_dir() -> str:
    """
    获取前端静态资源目录。
    - 开发环境：使用 main.py 同级的 dist 目录
    - PyInstaller 打包后：使用临时解压目录中的 dist 目录
    """
    if hasattr(sys, "_MEIPASS"):
        base_dir = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_dir, "dist")


STATIC_DIR = get_static_dir()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(energy_storage.router)
app.include_router(pv.router)


@app.get("/health")
async def health():
    return {"status": "healthy"}


# 将前端静态资源挂载到根路径：
#   - /            -> index.html
#   - /assets/...  -> 对应 js/css 等静态文件（包含 index-*.js）
if os.path.isdir(STATIC_DIR):
    app.mount(
        "/",
        StaticFiles(directory=STATIC_DIR, html=True),
        name="frontend",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
