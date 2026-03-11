"""通知推送路由（微信 / 邮件）"""
import json
import logging
import smtplib
import os
import urllib.request
import urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from fastapi import APIRouter, HTTPException

from api.models import WechatNotifyRequest, EmailNotifyRequest
from api.config import settings

router = APIRouter(prefix="/api/v1/notify", tags=["通知推送"])

log = logging.getLogger("api.notify")

# SMTP 配置（提取自 scraper_gui.py:16-27）
SMTP_SERVERS = {
    "qq.com":      ("smtp.qq.com",      465, True),
    "foxmail.com": ("smtp.qq.com",      465, True),
    "163.com":     ("smtp.163.com",     465, True),
    "126.com":     ("smtp.126.com",     465, True),
    "yeah.net":    ("smtp.yeah.net",    465, True),
    "gmail.com":   ("smtp.gmail.com",   465, True),
    "outlook.com": ("smtp-mail.outlook.com", 587, False),
    "hotmail.com": ("smtp-mail.outlook.com", 587, False),
    "sina.com":    ("smtp.sina.com",    465, True),
    "sohu.com":    ("smtp.sohu.com",    465, True),
}


def _get_smtp_config(email_addr: str):
    domain = email_addr.strip().split("@")[-1].lower()
    if domain in SMTP_SERVERS:
        return SMTP_SERVERS[domain]
    return (f"smtp.{domain}", 465, True)


def _smtp_connect(host: str, port: int, use_ssl: bool):
    if use_ssl:
        return smtplib.SMTP_SSL(host, port)
    else:
        srv = smtplib.SMTP(host, port)
        srv.starttls()
        return srv


@router.post("/wechat")
def send_wechat(req: WechatNotifyRequest):
    """发送微信推送（Server酱 / PushPlus）"""
    key = req.wx_key or settings.wx_key
    if not key:
        raise HTTPException(status_code=400, detail="缺少微信推送 Key")

    try:
        if req.wx_type == "pushplus":
            url = "https://www.pushplus.plus/send"
            payload = json.dumps({
                "token": key,
                "title": req.title,
                "content": req.content,
                "template": "markdown",
            }).encode("utf-8")
            http_req = urllib.request.Request(url, data=payload, method="POST")
            http_req.add_header("Content-Type", "application/json")
            with urllib.request.urlopen(http_req, timeout=10) as r:
                resp = json.loads(r.read().decode("utf-8"))
                ok = resp.get("code") == 200
        else:
            url = f"https://sctapi.ftqq.com/{key}.send"
            data = urllib.parse.urlencode({
                "title": req.title,
                "desp": req.content,
            }).encode("utf-8")
            http_req = urllib.request.Request(url, data=data, method="POST")
            http_req.add_header("Content-Type", "application/x-www-form-urlencoded")
            with urllib.request.urlopen(http_req, timeout=10) as r:
                resp = json.loads(r.read().decode("utf-8"))
                ok = (
                    resp.get("data", {}).get("errno", -1) == 0
                    or resp.get("code") == 0
                )

        if ok:
            return {"message": "推送成功", "type": req.wx_type}
        else:
            return {"message": "推送失败", "type": req.wx_type, "response": resp}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"推送异常: {e}")


@router.post("/email")
def send_email(req: EmailNotifyRequest):
    """发送邮件通知"""
    sender = req.email_from or settings.email_from
    pwd = req.email_pwd or settings.email_pwd
    to = req.email_to or settings.email_to

    if not sender or not pwd or not to:
        raise HTTPException(status_code=400, detail="缺少邮箱配置（发件人/密码/收件人）")

    try:
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = to
        msg["Subject"] = req.subject
        msg.attach(MIMEText(req.body, "plain", "utf-8"))

        # 添加附件
        if req.attachment_path and os.path.exists(req.attachment_path):
            with open(req.attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                fname = os.path.basename(req.attachment_path)
                from email.header import Header
                part.add_header(
                    "Content-Disposition", "attachment",
                    filename=("utf-8", "", fname),
                )
                msg.attach(part)

        host, port, use_ssl = _get_smtp_config(sender)
        log.info(f"SMTP: {host}:{port} (SSL={use_ssl})")
        with _smtp_connect(host, port, use_ssl) as srv:
            srv.login(sender, pwd)
            srv.sendmail(sender, [to], msg.as_string())

        return {"message": f"邮件已发送至 {to}"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"邮件发送失败: {e}")


@router.post("/test")
def test_notify():
    """测试通知配置（微信+邮件）"""
    results = {}

    # 测试微信
    if settings.wx_key:
        try:
            resp = send_wechat(WechatNotifyRequest(
                title="API 测试消息",
                content="**通知配置测试成功！**",
                wx_type=settings.wx_type,
                wx_key=settings.wx_key,
            ))
            results["wechat"] = "成功"
        except Exception as e:
            results["wechat"] = f"失败: {e}"
    else:
        results["wechat"] = "未配置"

    # 测试邮件
    if settings.email_from and settings.email_pwd and settings.email_to:
        try:
            resp = send_email(EmailNotifyRequest(
                subject="API 测试邮件",
                body="这是 TaobaoScraper API 的测试邮件。",
            ))
            results["email"] = "成功"
        except Exception as e:
            results["email"] = f"失败: {e}"
    else:
        results["email"] = "未配置"

    return results
