import httpx

from mcp_server.server import server, api, _json_text, _error_text


@server.tool()
async def list_tasks(
    limit: int = 20,
) -> str:
    """列出最近的采集任务及其状态。

    返回任务列表，包含 task_id、类型、状态、进度等信息。
    状态: pending(等待) / running(运行中) / completed(完成) / failed(失败) / cancelled(已取消)
    """
    try:
        result = await api.get("/api/v1/tasks", params={"limit": limit})
        return _json_text(result)
    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)


@server.tool()
async def get_task(
    task_id: str,
    include_result: bool = False,
) -> str:
    """查询单个任务的详细状态。

    设置 include_result=True 可同时获取任务的完整结果数据（仅在任务完成时有效）。
    用于查看采集任务的进度或获取最终结果。
    """
    try:
        task_info = await api.get(f"/api/v1/tasks/{task_id}")

        if include_result and task_info.get("status") == "completed":
            task_result = await api.get(f"/api/v1/tasks/{task_id}/result")
            task_info["result"] = task_result.get("result")
            task_info["result_file"] = task_result.get("result_file")

        return _json_text(task_info)

    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 404:
            return _error_text(f"任务不存在: {task_id}")
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)


@server.tool()
async def cancel_task(
    task_id: str,
) -> str:
    """取消一个正在运行或等待中的采集任务。"""
    try:
        result = await api.post(f"/api/v1/tasks/{task_id}/cancel")
        return _json_text(result)
    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 400:
            return _error_text("任务无法取消（可能已完成或不存在）")
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)
