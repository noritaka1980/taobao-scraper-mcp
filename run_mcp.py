"""TaobaoScraper MCP Server 启动脚本

用法:
    # stdio 模式（Claude Desktop / Claude Code 推荐）
    python run_mcp.py

    # SSE 模式（Web 客户端 / 远程调用）
    python run_mcp.py --transport sse
    python run_mcp.py --transport sse --port 8001

环境变量:
    SCRAPER_API_URL   - FastAPI 地址（默认 http://localhost:8000）
    SCRAPER_API_KEY   - API 认证密钥（默认 changeme-your-secret-key）
    MCP_SSE_HOST      - SSE 监听地址（默认 0.0.0.0）
    MCP_SSE_PORT      - SSE 监听端口（默认 8001）

注意: 启动前请先确保 FastAPI 服务已运行 (python run_api.py)
"""
import sys
import os
import argparse

# 确保项目目录在 sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    parser = argparse.ArgumentParser(description="TaobaoScraper MCP Server")
    parser.add_argument(
        "--transport", choices=["stdio", "sse"], default="stdio",
        help="传输协议 (默认 stdio)",
    )
    parser.add_argument("--host", default=None, help="SSE 监听地址")
    parser.add_argument("--port", type=int, default=None, help="SSE 监听端口")
    args = parser.parse_args()

    from mcp_server.server import server
    from mcp_server.config import MCPConfig
    import asyncio

    cfg = MCPConfig()

    if args.transport == "stdio":
        asyncio.run(server.run_stdio_async())

    elif args.transport == "sse":
        host = args.host or cfg.SSE_HOST
        port = args.port or cfg.SSE_PORT
        # 新版 MCP SDK: host/port 在 FastMCP 构造函数中设置
        server._host = host
        server._port = port
        print(f"MCP SSE Server 启动: http://{host}:{port}/sse")
        asyncio.run(server.run_sse_async())


if __name__ == "__main__":
    main()
