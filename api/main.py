"""FastAPI 应用入口"""
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Depends
from fastapi.responses import JSONResponse
from selenium.common.exceptions import WebDriverException

from api.config import settings
from api.auth import verify_api_key
from api.chrome_manager import ChromeManager
from api.task_manager import TaskManager

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# 全局实例
chrome_manager = ChromeManager(settings)
task_manager = TaskManager(chrome_manager, max_history=settings.task_history_max)


@asynccontextmanager
async def lifespan(app: FastAPI):
    logging.getLogger("api").info("TaobaoScraper API 启动中...")
    yield
    logging.getLogger("api").info("TaobaoScraper API 关闭中...")
    chrome_manager.shutdown()


app = FastAPI(
    title="TaobaoScraper API",
    description="淘宝/京东数据采集工具 API，支持多平台采集、评价分析、热词采集等功能。",
    version="1.0.0",
    lifespan=lifespan,
    dependencies=[Depends(verify_api_key)],
)


# ── 全局异常处理 ──

@app.exception_handler(WebDriverException)
async def selenium_exception_handler(request, exc):
    return JSONResponse(status_code=503, content={
        "error": "Chrome 连接异常",
        "detail": str(exc),
        "suggestion": "请检查 Chrome 是否运行，或调用 POST /api/v1/system/chrome/restart",
    })


@app.exception_handler(TimeoutError)
async def timeout_handler(request, exc):
    return JSONResponse(status_code=429, content={
        "error": "系统繁忙",
        "detail": str(exc),
        "suggestion": "当前有任务正在运行，请稍后再试",
    })


# ── 注册路由 ──

from api.routers import scrape, reviews, hotwords, export, notify, tasks, system

app.include_router(system.router)
app.include_router(tasks.router)
app.include_router(scrape.router)
app.include_router(reviews.router)
app.include_router(hotwords.router)
app.include_router(export.router)
app.include_router(notify.router)
