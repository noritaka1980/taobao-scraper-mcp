import json

from mcp.server.fastmcp import FastMCP

from mcp_server.config import MCPConfig
from mcp_server.api_client import ScraperAPIClient

# 创建 MCP Server 实例
server = FastMCP("taobao-scraper")

# 全局配置和 API 客户端
config = MCPConfig()
api = ScraperAPIClient(
    base_url=config.API_BASE_URL,
    api_key=config.API_KEY,
    timeout=config.HTTP_TIMEOUT,
)


def _json_text(data) -> str:
    """将数据格式化为 JSON 文本返回"""
    return json.dumps(data, ensure_ascii=False, indent=2)


def _error_text(error: str, suggestion: str = "") -> str:
    """返回结构化错误信息"""
    result = {"error": error}
    if suggestion:
        result["suggestion"] = suggestion
    return _json_text(result)


# 导入注册所有 tool 模块
from mcp_server.tools import scraping   # noqa: E402, F401
from mcp_server.tools import tasks      # noqa: E402, F401
from mcp_server.tools import export     # noqa: E402, F401
from mcp_server.tools import notify     # noqa: E402, F401
from mcp_server.tools import system     # noqa: E402, F401
