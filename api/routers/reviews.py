"""评价采集与分析路由"""
import logging

from fastapi import APIRouter

from api.models import (
    ReviewScrapeRequest,
    ReviewAnalyseRequest,
    ReviewAnalyseResponse,
    TaskSubmitResponse,
)

router = APIRouter(prefix="/api/v1/reviews", tags=["评价分析"])


@router.post("/scrape", response_model=TaskSubmitResponse)
def scrape_reviews(req: ReviewScrapeRequest):
    """采集单商品评价（异步，需要 Chrome）"""
    from api.main import task_manager

    def executor(driver, task):
        from scraper_core import scrape_reviews as _scrape, keep_minimized
        log = logging.getLogger(f"api.reviews.{task.task_id}")
        keep_minimized(driver)
        reviews = _scrape(
            driver, req.item_url, req.platform,
            req.max_reviews, log, task.stop_event,
        )
        task.progress = f"采集到 {len(reviews)} 条评价"
        return reviews

    task = task_manager.submit(
        "reviews",
        {"item_url": req.item_url, "platform": req.platform},
        executor,
    )
    return TaskSubmitResponse(task_id=task.task_id)


@router.post("/analyse", response_model=ReviewAnalyseResponse)
def analyse_reviews(req: ReviewAnalyseRequest):
    """分析评价文本（同步，纯计算，不需要 Chrome）"""
    from scraper_core import analyse_reviews as _analyse
    log = logging.getLogger("api.reviews.analyse")
    result = _analyse(req.reviews, log)
    return ReviewAnalyseResponse(
        total=len(req.reviews),
        keywords=result,
    )
