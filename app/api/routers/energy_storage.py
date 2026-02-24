from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from app.services.excel_reader import ExcelReader
from app.services.excel_writer import ExcelWriter
from app.core.config import ENERGY_STORAGE_TEMPLATE
import os
import shutil
from datetime import datetime

router = APIRouter(prefix="/api/energy-storage", tags=["储能收资"])


@router.post("/info-collection")
async def energy_storage_info_collection(file: UploadFile = File(...)):
    try:
        temp_input_path = f"temp_input_{datetime.now().timestamp()}.xlsx"
        with open(temp_input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        reader = ExcelReader(temp_input_path)
        customer_data = reader.read_customer_data()
        reader.close()

        writer = ExcelWriter(str(ENERGY_STORAGE_TEMPLATE))
        output_filename = f"储能收资表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = writer.write_customer_data(customer_data, output_filename)
        writer.close()

        os.remove(temp_input_path)

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_filename
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")


@router.get("/health")
async def health_check():
    return {"status": "ok", "service": "储能收资"}
