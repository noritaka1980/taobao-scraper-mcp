import httpx

from mcp_server.server import server, api, _json_text, _error_text


@server.tool()
async def system_status(
    action: str = "check",
) -> str:
    """查看系统状态或执行系统操作。

    - action="check": 返回 API 健康状态和 Chrome 浏览器状态（是否存活、是否忙碌）
    - action="restart_chrome": 重启 Chrome 浏览器（采集出错或卡死时使用）
    """
    try:
        if action == "check":
            health = await api.get("/api/v1/system/health")
            chrome = await api.get("/api/v1/system/chrome")
            return _json_text({"health": health, "chrome": chrome})

        elif action == "restart_chrome":
            result = await api.post("/api/v1/system/chrome/restart")
            return _json_text(result)

        else:
            return _error_text(
                f"不支持的操作: {action}",
                "请选择 check / restart_chrome",
            )

    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)
