"""系统管理路由"""
from fastapi import APIRouter

from api.models import HealthResponse, ChromeStatusResponse

router = APIRouter(prefix="/api/v1/system", tags=["系统管理"])


@router.get("/health", response_model=HealthResponse)
def health_check():
    """健康检查"""
    return HealthResponse()


@router.get("/chrome", response_model=ChromeStatusResponse)
def chrome_status():
    """查询 Chrome 状态"""
    from api.main import chrome_manager
    return chrome_manager.status()


@router.post("/chrome/restart")
def chrome_restart():
    """重启 Chrome"""
    from api.main import chrome_manager
    if chrome_manager.is_busy:
        return {"error": "Chrome 正在执行任务，无法重启", "current_task": chrome_manager.current_task_id}
    try:
        chrome_manager.restart()
        return {"message": "Chrome 已重启"}
    except Exception as e:
        return {"error": f"重启失败: {e}"}
