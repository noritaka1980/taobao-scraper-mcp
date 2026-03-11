"""异步任务队列管理"""
import threading
import uuid
import logging
from datetime import datetime
from typing import Optional, Callable, Any
from collections import OrderedDict

from api.models import TaskStatus

log = logging.getLogger("task_manager")


class Task:

    def __init__(self, task_id: str, task_type: str, params: dict):
        self.task_id = task_id
        self.task_type = task_type
        self.params = params
        self.status = TaskStatus.PENDING
        self.created_at = datetime.now()
        self.started_at: Optional[datetime] = None
        self.finished_at: Optional[datetime] = None
        self.progress: Optional[str] = None
        self.result: Any = None
        self.result_file: Optional[str] = None
        self.error: Optional[str] = None
        self.stop_event = threading.Event()

    def to_response(self) -> dict:
        return {
            "task_id": self.task_id,
            "task_type": self.task_type,
            "status": self.status,
            "created_at": self.created_at.isoformat(),
            "started_at": self.started_at.isoformat() if self.started_at else None,
            "finished_at": self.finished_at.isoformat() if self.finished_at else None,
            "progress": self.progress,
            "result_summary": self._make_summary(),
            "result_file": self.result_file,
            "error": self.error,
        }

    def _make_summary(self) -> Optional[str]:
        if self.status != TaskStatus.COMPLETED or self.result is None:
            return None
        if self.task_type == "taobao":
            return f"采集到 {self.result} 条淘宝数据"
        if self.task_type == "jd":
            if isinstance(self.result, list):
                return f"采集到 {len(self.result)} 条京东数据"
            return str(self.result)
        if self.task_type == "multiplatform":
            return "多平台采集完成"
        if self.task_type == "reviews":
            if isinstance(self.result, list):
                return f"采集到 {len(self.result)} 条评价"
            return str(self.result)
        if self.task_type == "hotwords":
            return "热词采集完成"
        return str(self.result)[:200]


class TaskManager:

    def __init__(self, chrome_manager, max_history: int = 100):
        self._tasks: OrderedDict[str, Task] = OrderedDict()
        self._chrome_manager = chrome_manager
        self._max_history = max_history
        self._lock = threading.Lock()

    def submit(self, task_type: str, params: dict, executor: Callable) -> Task:
        """提交一个异步任务，executor 签名: (driver, task) -> result"""
        task_id = (
            f"{task_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            f"_{uuid.uuid4().hex[:6]}"
        )
        task = Task(task_id, task_type, params)

        with self._lock:
            self._tasks[task_id] = task
            while len(self._tasks) > self._max_history:
                self._tasks.popitem(last=False)

        def _run():
            task.status = TaskStatus.RUNNING
            task.started_at = datetime.now()
            try:
                driver = self._chrome_manager.acquire(task_id)
                try:
                    result = executor(driver, task)
                    task.result = result
                    task.status = TaskStatus.COMPLETED
                finally:
                    self._chrome_manager.release()
            except Exception as e:
                task.error = str(e)
                task.status = TaskStatus.FAILED
                log.error(f"任务 {task_id} 失败: {e}", exc_info=True)
            finally:
                task.finished_at = datetime.now()

        thread = threading.Thread(target=_run, daemon=True, name=f"task-{task_id}")
        thread.start()
        return task

    def get(self, task_id: str) -> Optional[Task]:
        return self._tasks.get(task_id)

    def cancel(self, task_id: str) -> bool:
        task = self._tasks.get(task_id)
        if not task:
            return False
        if task.status in (TaskStatus.COMPLETED, TaskStatus.FAILED, TaskStatus.CANCELLED):
            return False
        task.stop_event.set()
        task.status = TaskStatus.CANCELLED
        task.finished_at = datetime.now()
        return True

    def list_tasks(self, limit: int = 20) -> list[dict]:
        tasks = list(self._tasks.values())
        return [t.to_response() for t in reversed(tasks[-limit:])]
