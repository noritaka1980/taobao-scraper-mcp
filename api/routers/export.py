"""Excel 文件导出与下载路由"""
import os

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from api.config import settings

router = APIRouter(prefix="/api/v1/export", tags=["文件导出"])


@router.get("/files")
def list_files():
    """列出输出目录中的 Excel 文件"""
    files = []
    for f in os.listdir(settings.output_dir):
        if f.endswith((".xlsx", ".xls")):
            fp = os.path.join(settings.output_dir, f)
            files.append({
                "filename": f,
                "size_kb": round(os.path.getsize(fp) / 1024, 1),
                "modified": os.path.getmtime(fp),
            })
    files.sort(key=lambda x: x["modified"], reverse=True)
    return files


@router.get("/download/{filename}")
def download_file(filename: str):
    """下载指定文件"""
    # 安全检查：防止路径遍历
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="非法文件名")

    filepath = os.path.join(settings.output_dir, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail=f"文件不存在: {filename}")

    return FileResponse(
        filepath,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
