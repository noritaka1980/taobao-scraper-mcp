"""热词采集路由"""
import logging

from fastapi import APIRouter

from api.models import HotwordsRequest, TaskSubmitResponse

router = APIRouter(prefix="/api/v1/hotwords", tags=["热词采集"])


@router.post("/scrape", response_model=TaskSubmitResponse)
def scrape_hotwords(req: HotwordsRequest):
    """采集热门关键词（异步，需要 Chrome）"""
    from api.main import task_manager

    def executor(driver, task):
        from scraper_core import scrape_hot_keywords, keep_minimized
        log = logging.getLogger(f"api.hotwords.{task.task_id}")
        keep_minimized(driver)
        task.progress = "正在采集热词..."
        result = scrape_hot_keywords(driver, req.keyword, log, task.stop_event)
        task.progress = "热词采集完成"
        return result

    task = task_manager.submit("hotwords", {"keyword": req.keyword}, executor)
    return TaskSubmitResponse(task_id=task.task_id)
