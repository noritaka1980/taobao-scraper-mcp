"""任务管理路由"""
from fastapi import APIRouter, HTTPException, Query

from api.models import TaskStatus

router = APIRouter(prefix="/api/v1/tasks", tags=["任务管理"])


@router.get("")
def list_tasks(limit: int = Query(20, ge=1, le=100)):
    """列出所有任务"""
    from api.main import task_manager
    return task_manager.list_tasks(limit=limit)


@router.get("/{task_id}")
def get_task(task_id: str):
    """查询单个任务状态"""
    from api.main import task_manager
    task = task_manager.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail=f"任务不存在: {task_id}")
    return task.to_response()


@router.post("/{task_id}/cancel")
def cancel_task(task_id: str):
    """取消任务"""
    from api.main import task_manager
    ok = task_manager.cancel(task_id)
    if not ok:
        raise HTTPException(status_code=400, detail="任务无法取消（可能已完成或不存在）")
    return {"message": "任务已取消"}


@router.get("/{task_id}/result")
def get_task_result(task_id: str):
    """获取任务结果数据"""
    from api.main import task_manager
    task = task_manager.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail=f"任务不存在: {task_id}")
    if task.status != TaskStatus.COMPLETED:
        raise HTTPException(
            status_code=400,
            detail=f"任务状态为 {task.status.value}，尚无结果",
        )
    return {
        "task_id": task_id,
        "result": task.result,
        "result_file": task.result_file,
    }
