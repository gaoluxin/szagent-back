from fastapi import APIRouter

router = APIRouter(prefix="/api/pv", tags=["光伏收资"])


@router.get("/health")
async def health_check():
    return {"status": "ok", "service": "光伏收资", "message": "功能暂未开发"}
