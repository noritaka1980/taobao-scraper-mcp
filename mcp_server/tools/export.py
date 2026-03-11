import httpx

from mcp_server.server import server, api, _json_text, _error_text


@server.tool()
async def list_files() -> str:
    """列出所有已导出的 Excel 数据文件。

    返回文件名、大小(KB)、修改时间等信息。
    这些文件是采集任务自动生成的结果。
    用户可通过 FastAPI 接口下载: GET /api/v1/export/download/{filename}
    """
    try:
        result = await api.get("/api/v1/export/files")
        return _json_text(result)
    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)
