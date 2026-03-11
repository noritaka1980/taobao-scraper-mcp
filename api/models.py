"""Pydantic 请求/响应模型"""
from pydantic import BaseModel, Field
from typing import Optional, Literal, Any
from datetime import datetime
from enum import Enum


# ── 任务状态 ──

class TaskStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class TaskSubmitResponse(BaseModel):
    task_id: str
    message: str = "任务已提交"


class TaskResponse(BaseModel):
    task_id: str
    task_type: str
    status: TaskStatus
    created_at: str
    started_at: Optional[str] = None
    finished_at: Optional[str] = None
    progress: Optional[str] = None
    result_summary: Optional[str] = None
    result_file: Optional[str] = None
    error: Optional[str] = None


class TaskResultResponse(BaseModel):
    task_id: str
    result: Any
    result_file: Optional[str] = None


# ── 采集请求 ──

class TaobaoScrapeRequest(BaseModel):
    keyword: str = Field(..., description="搜索关键词")
    pages: int = Field(3, ge=1, le=20, description="采集页数")
    sort_by: Literal["sale", "default", "price_asc", "price_desc"] = "sale"
    tmall_only: bool = False
    exact_match: bool = False
    price_min: float = 0
    price_max: float = 0
    output_file: Optional[str] = None
    own_shop: str = ""
    img_size: int = Field(60, ge=0, le=200, description="图片缩放尺寸，0=不采集图片")


class JDScrapeRequest(BaseModel):
    keyword: str = Field(..., description="搜索关键词")
    pages: int = Field(3, ge=1, le=20, description="采集页数")
    output_file: Optional[str] = None


class MultiPlatformRequest(BaseModel):
    keyword: str = Field(..., description="搜索关键词")
    pages: int = Field(3, ge=1, le=20, description="采集页数")
    sort_by: Literal["sale", "default", "price_asc", "price_desc"] = "sale"
    tmall_only: bool = False
    exact_match: bool = False
    platforms: list[Literal["jd"]] = Field(["jd"], description="要采集的平台列表")
    scrape_reviews: bool = True
    scrape_hotwords: bool = True
    output_file: Optional[str] = None
    own_shop: str = ""
    img_size: int = Field(60, ge=0, le=200)


# ── 评价 ──

class ReviewScrapeRequest(BaseModel):
    item_url: str = Field(..., description="商品链接")
    platform: Literal["taobao", "jd"] = "taobao"
    max_reviews: int = Field(20, ge=1, le=100, description="最大评价数")


class ReviewAnalyseRequest(BaseModel):
    reviews: list[str] = Field(..., description="评价文本列表")


class ReviewAnalyseResponse(BaseModel):
    total: int = 0
    keywords: Any = None


# ── 热词 ──

class HotwordsRequest(BaseModel):
    keyword: str = Field(..., description="搜索关键词")


# ── 通知 ──

class WechatNotifyRequest(BaseModel):
    title: str = Field(..., description="消息标题")
    content: str = Field(..., description="消息内容（支持 Markdown）")
    wx_type: Literal["server_chan", "pushplus"] = "server_chan"
    wx_key: Optional[str] = Field(None, description="推送 Key，为空则用配置文件中的")


class EmailNotifyRequest(BaseModel):
    subject: str = Field(..., description="邮件主题")
    body: str = Field(..., description="邮件正文")
    attachment_path: Optional[str] = Field(None, description="附件文件路径")
    email_from: Optional[str] = None
    email_pwd: Optional[str] = None
    email_to: Optional[str] = None


# ── 系统 ──

class ChromeStatusResponse(BaseModel):
    driver_alive: bool
    is_busy: bool
    current_task_id: Optional[str] = None


class HealthResponse(BaseModel):
    status: str = "ok"
    version: str = "1.0.0"
