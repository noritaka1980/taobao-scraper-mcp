"""采集路由（淘宝、京东、多平台）"""
import re
import logging

from fastapi import APIRouter, HTTPException

from api.models import (
    TaobaoScrapeRequest,
    JDScrapeRequest,
    MultiPlatformRequest,
    TaskSubmitResponse,
)
from api.config import settings

router = APIRouter(prefix="/api/v1/scrape", tags=["数据采集"])


def _make_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(f"api.scrape.{name}")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        fh = logging.FileHandler("scraper_log.txt", encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            f"[API:{name}] %(asctime)s %(message)s", datefmt="%H:%M:%S",
        ))
        logger.addHandler(fh)
    return logger


def _safe_filename(keyword: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", keyword)


@router.post("/taobao", response_model=TaskSubmitResponse)
def scrape_taobao(req: TaobaoScrapeRequest):
    """提交淘宝采集任务（异步）"""
    from api.main import task_manager

    output_file = req.output_file or f"{_safe_filename(req.keyword)}_data.xlsx"
    cfg = {
        "keyword": req.keyword,
        "keywords": [req.keyword],
        "pages": req.pages,
        "sort_by": req.sort_by,
        "tmall_only": req.tmall_only,
        "exact": req.exact_match,
        "price_min": req.price_min,
        "price_max": req.price_max,
        "output_file": output_file,
        "own_shop": req.own_shop or settings.own_shop,
        "img_size": req.img_size,
    }

    def executor(driver, task):
        from scraper_core import run_round, keep_minimized
        log = _make_logger(task.task_id)
        keep_minimized(driver)
        count = run_round(driver, cfg, 1, log, task.stop_event)
        task.result_file = output_file
        task.progress = f"采集完成，共 {count} 条"
        return count

    task = task_manager.submit("taobao", cfg, executor)
    return TaskSubmitResponse(task_id=task.task_id)


@router.post("/jd", response_model=TaskSubmitResponse)
def scrape_jd(req: JDScrapeRequest):
    """提交京东采集任务（异步）"""
    from api.main import task_manager

    output_file = req.output_file or f"{_safe_filename(req.keyword)}_jd_data.xlsx"

    def executor(driver, task):
        from scraper_core import scrape_jd, keep_minimized
        log = _make_logger(task.task_id)
        keep_minimized(driver)
        items = scrape_jd(
            driver, req.keyword, req.pages, log, task.stop_event,
        )
        # 保存结果
        if items:
            from scraper_core import save_excel
            save_excel(items, output_file, log)
        task.result_file = output_file
        task.progress = f"采集完成，共 {len(items)} 条"
        return items

    task = task_manager.submit("jd", {"keyword": req.keyword, "pages": req.pages}, executor)
    return TaskSubmitResponse(task_id=task.task_id)


@router.post("/multiplatform", response_model=TaskSubmitResponse)
def scrape_multiplatform(req: MultiPlatformRequest):
    """提交多平台采集任务（异步）：先淘宝再京东+评价+热词"""
    from api.main import task_manager

    output_file = req.output_file or f"{_safe_filename(req.keyword)}_data.xlsx"
    cfg = {
        "keyword": req.keyword,
        "keywords": [req.keyword],
        "pages": req.pages,
        "sort_by": req.sort_by,
        "tmall_only": req.tmall_only,
        "exact": req.exact_match,
        "output_file": output_file,
        "own_shop": req.own_shop or settings.own_shop,
        "img_size": req.img_size,
        "platforms": req.platforms,
        "scrape_reviews": req.scrape_reviews,
        "scrape_hotwords": req.scrape_hotwords,
    }

    def executor(driver, task):
        from scraper_core import run_round, run_multiplatform, keep_minimized
        log = _make_logger(task.task_id)
        keep_minimized(driver)

        # 第一步：淘宝采集
        task.progress = "正在采集淘宝..."
        tb_count = run_round(driver, cfg, 1, log, task.stop_event)
        if task.stop_event.is_set():
            return {"taobao": tb_count}

        # 第二步：多平台采集
        task.progress = "正在采集其他平台..."
        mp_cfg = dict(cfg)
        mp_cfg["platforms"] = []
        if "jd" in req.platforms:
            mp_cfg["platforms"].append("jd")
        if req.scrape_reviews:
            mp_cfg["platforms"].append("review")
        if req.scrape_hotwords:
            mp_cfg["platforms"].append("hotwords")

        result = run_multiplatform(driver, mp_cfg, log, task.stop_event)
        task.result_file = output_file
        task.progress = "多平台采集完成"
        return {"taobao": tb_count, "multiplatform": str(result)}

    task = task_manager.submit("multiplatform", cfg, executor)
    return TaskSubmitResponse(task_id=task.task_id)
