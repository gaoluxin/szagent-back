from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api.routers import energy_storage, pv

app = FastAPI(
    title="收资工具-后台",
    description="提供光伏收资和储能收资功能的API接口",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(energy_storage.router)
app.include_router(pv.router)


@app.get("/")
async def root():
    return {
        "message": "收资工具-后台服务",
        "version": "1.0.0",
        "modules": {
            "光伏收资": "暂未开发",
            "储能收资": "已开发"
        }
    }


@app.get("/health")
async def health():
    return {"status": "healthy"}
