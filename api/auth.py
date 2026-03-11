"""API Key 认证"""
from fastapi import Security, HTTPException
from fastapi.security import APIKeyHeader, APIKeyQuery

from api.config import settings

api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)
api_key_query = APIKeyQuery(name="api_key", auto_error=False)


def verify_api_key(
    header_key: str = Security(api_key_header),
    query_key: str = Security(api_key_query),
) -> str:
    """验证 API Key（Header 或 Query 参数均可）"""
    key = header_key or query_key
    if not key or key != settings.api_key:
        raise HTTPException(status_code=401, detail="无效或缺少 API Key")
    return key
