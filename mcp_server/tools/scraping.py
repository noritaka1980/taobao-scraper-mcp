import httpx

from mcp_server.server import server, api, config, _json_text, _error_text


@server.tool()
async def scrape_products(
    keyword: str,
    platform: str = "taobao",
    pages: int = 3,
    sort_by: str = "sale",
    tmall_only: bool = False,
    exact_match: bool = False,
    price_min: float = 0,
    price_max: float = 0,
    jd_platforms: list[str] | None = None,
    scrape_hotwords: bool = True,
    wait_for_result: bool = True,
    timeout: int = 300,
) -> str:
    """采集电商平台商品数据。

    支持三种模式：
    - platform="taobao": 淘宝/天猫商品采集
    - platform="jd": 京东商品采集
    - platform="multi": 多平台采集（淘宝+京东+评价+热词）

    默认自动等待采集完成并返回结果摘要。设置 wait_for_result=False 可立即返回 task_id。
    采集结果会自动保存为 Excel 文件，可通过 list_files 查看。
    """
    try:
        if platform == "taobao":
            path = "/api/v1/scrape/taobao"
            data = {
                "keyword": keyword,
                "pages": pages,
                "sort_by": sort_by,
                "tmall_only": tmall_only,
                "exact_match": exact_match,
                "price_min": price_min,
                "price_max": price_max,
            }
        elif platform == "jd":
            path = "/api/v1/scrape/jd"
            data = {"keyword": keyword, "pages": pages}
        elif platform == "multi":
            path = "/api/v1/scrape/multiplatform"
            data = {
                "keyword": keyword,
                "pages": pages,
                "sort_by": sort_by,
                "tmall_only": tmall_only,
                "exact_match": exact_match,
                "platforms": jd_platforms or ["jd"],
                "scrape_hotwords": scrape_hotwords,
            }
        else:
            return _error_text(f"不支持的平台: {platform}", "请选择 taobao / jd / multi")

        if wait_for_result:
            result = await api.submit_and_wait(
                path, data,
                timeout=timeout,
                poll_interval=config.TASK_POLL_INTERVAL,
            )
        else:
            result = await api.post(path, data=data)

        return _json_text(result)

    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动 (python run_api.py)")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)


@server.tool()
async def scrape_hotwords(
    keyword: str,
    wait_for_result: bool = True,
    timeout: int = 120,
) -> str:
    """采集指定关键词的热门搜索词/相关搜索词。

    用于分析市场热点、发现关联关键词。
    默认等待采集完成返回结果。
    """
    try:
        data = {"keyword": keyword}
        if wait_for_result:
            result = await api.submit_and_wait(
                "/api/v1/hotwords/scrape", data,
                timeout=timeout,
                poll_interval=config.TASK_POLL_INTERVAL,
            )
        else:
            result = await api.post("/api/v1/hotwords/scrape", data=data)
        return _json_text(result)

    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动 (python run_api.py)")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)
