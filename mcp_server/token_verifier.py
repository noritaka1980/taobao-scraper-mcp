"""简单的 API Key Token 验证器，用于 MCP SSE 模式认证"""

from mcp.server.auth.provider import TokenVerifier, AccessToken


class SimpleApiKeyTokenVerifier(TokenVerifier):
    """基于 API Key 的 Bearer Token 验证器

    客户端连接 SSE 时需要在 Header 中带上:
        Authorization: Bearer <api_key>
    """

    def __init__(self, api_key: str):
        self.api_key = api_key

    async def verify_token(self, token: str) -> AccessToken | None:
        """验证 Bearer Token 是否匹配 API Key"""
        if token == self.api_key:
            return AccessToken(
                token=token,
                client_id="mcp-client",
                scopes=["scrape"],
            )
        return None
