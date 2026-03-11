import asyncio
import logging

import httpx

log = logging.getLogger("mcp.api_client")


class ScraperAPIClient:
    """FastAPI 后端的 httpx 异步客户端"""

    def __init__(self, base_url: str, api_key: str, timeout: float = 30.0):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.timeout = timeout

    def _headers(self) -> dict:
        return {"X-API-Key": self.api_key}

    async def _request(self, method: str, path: str, **kwargs) -> dict:
        async with httpx.AsyncClient(timeout=self.timeout) as client:
            url = f"{self.base_url}{path}"
            resp = await client.request(method, url, headers=self._headers(), **kwargs)
            resp.raise_for_status()
            return resp.json()

    async def get(self, path: str, params: dict | None = None) -> dict:
        return await self._request("GET", path, params=params)

    async def post(self, path: str, data: dict | None = None) -> dict:
        return await self._request("POST", path, json=data)

    async def submit_and_wait(
        self,
        submit_path: str,
        submit_data: dict,
        timeout: int = 300,
        poll_interval: float = 3.0,
    ) -> dict:
        """提交异步任务并轮询等待结果。

        1. POST submit_path 提交任务 → task_id
        2. 轮询 GET /tasks/{task_id} 直到完成/失败/超时
        3. 完成时自动获取结果返回
        """
        # 提交任务
        result = await self.post(submit_path, data=submit_data)
        task_id = result["task_id"]
        log.info(f"任务已提交: {task_id}")

        # 轮询等待
        elapsed = 0.0
        last_progress = ""
        while elapsed < timeout:
            await asyncio.sleep(poll_interval)
            elapsed += poll_interval

            task_info = await self.get(f"/api/v1/tasks/{task_id}")
            status = task_info.get("status")
            progress = task_info.get("progress", "")

            if progress and progress != last_progress:
                log.info(f"[{task_id}] {progress}")
                last_progress = progress

            if status == "completed":
                task_result = await self.get(f"/api/v1/tasks/{task_id}/result")
                return {
                    "status": "completed",
                    "task_id": task_id,
                    "summary": task_info.get("result_summary", ""),
                    "result_file": task_info.get("result_file"),
                    "result": task_result.get("result"),
                }

            if status in ("failed", "cancelled"):
                return {
                    "status": status,
                    "task_id": task_id,
                    "error": task_info.get("error", "未知错误"),
                }

        # 超时
        return {
            "status": "timeout",
            "task_id": task_id,
            "message": f"任务在 {timeout} 秒内未完成，可通过 get_task('{task_id}') 继续查询",
            "last_progress": last_progress,
        }
