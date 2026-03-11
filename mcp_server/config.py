import os


class MCPConfig:
    """MCP Server 配置，全部通过环境变量控制"""

    # FastAPI 后端地址
    API_BASE_URL: str = os.environ.get("SCRAPER_API_URL", "http://localhost:8000")

    # FastAPI 认证 Key（须与 FastAPI 的 SCRAPER_API_KEY 一致）
    API_KEY: str = os.environ.get("SCRAPER_API_KEY", "changeme-your-secret-key")

    # HTTP 单次请求超时（秒）
    HTTP_TIMEOUT: float = float(os.environ.get("MCP_HTTP_TIMEOUT", "30"))

    # 异步任务轮询默认超时（秒）
    TASK_POLL_TIMEOUT: int = int(os.environ.get("MCP_TASK_TIMEOUT", "300"))

    # 异步任务轮询间隔（秒）
    TASK_POLL_INTERVAL: float = float(os.environ.get("MCP_POLL_INTERVAL", "3"))

    # SSE 模式监听配置
    SSE_HOST: str = os.environ.get("MCP_SSE_HOST", "0.0.0.0")
    SSE_PORT: int = int(os.environ.get("MCP_SSE_PORT", "8001"))

    # SSE 认证开关（true=启用 Bearer Token 认证，false=关闭）
    SSE_AUTH_ENABLED: bool = os.environ.get("MCP_SSE_AUTH", "true").lower() == "true"
