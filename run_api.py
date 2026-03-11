"""TaobaoScraper API 启动脚本

用法:
    python run_api.py
    python run_api.py --port 8080
    python run_api.py --host 0.0.0.0 --port 8000

启动后访问:
    API 文档: http://localhost:8000/docs
    健康检查: http://localhost:8000/api/v1/system/health
"""
import sys
import os
import argparse

# 确保项目根目录在 sys.path 中，以便 import scraper_core
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="TaobaoScraper API Server")
    parser.add_argument("--host", default="0.0.0.0", help="监听地址 (默认 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8000, help="监听端口 (默认 8000)")
    parser.add_argument("--reload", action="store_true", help="开发模式（热重载）")
    args = parser.parse_args()

    import uvicorn
    uvicorn.run(
        "api.main:app",
        host=args.host,
        port=args.port,
        reload=args.reload,
        log_level="info",
    )
