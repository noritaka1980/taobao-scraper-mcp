import httpx

from mcp_server.server import server, api, _json_text, _error_text


@server.tool()
async def send_notification(
    channel: str,
    title: str = "",
    content: str = "",
    wx_type: str = "server_chan",
    wx_key: str | None = None,
    email_from: str | None = None,
    email_pwd: str | None = None,
    email_to: str | None = None,
    attachment_path: str | None = None,
) -> str:
    """发送通知推送。

    三种模式：
    - channel="wechat": 微信推送（需要 title + content）
      wx_type 可选 "server_chan" 或 "pushplus"
    - channel="email": 邮件通知（需要 title 作为主题 + content 作为正文）
    - channel="test": 测试所有已配置的通知渠道

    未传入的凭证参数会使用服务器配置文件中的默认值。
    """
    try:
        if channel == "wechat":
            data = {"title": title, "content": content, "wx_type": wx_type}
            if wx_key:
                data["wx_key"] = wx_key
            result = await api.post("/api/v1/notify/wechat", data=data)

        elif channel == "email":
            data = {"subject": title, "body": content}
            if email_from:
                data["email_from"] = email_from
            if email_pwd:
                data["email_pwd"] = email_pwd
            if email_to:
                data["email_to"] = email_to
            if attachment_path:
                data["attachment_path"] = attachment_path
            result = await api.post("/api/v1/notify/email", data=data)

        elif channel == "test":
            result = await api.post("/api/v1/notify/test")

        else:
            return _error_text(
                f"不支持的通知渠道: {channel}",
                "请选择 wechat / email / test",
            )

        return _json_text(result)

    except httpx.ConnectError:
        return _error_text("无法连接 FastAPI 后端", "请确认 FastAPI 服务已启动")
    except httpx.HTTPStatusError as e:
        return _error_text(f"API 返回错误: {e.response.status_code}", e.response.text)
