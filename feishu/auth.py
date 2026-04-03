"""
feishu/auth.py
==============
飞书应用鉴权管理 —— 整个项目的「门禁钥匙」。

职责：
  1. 管理 tenant_access_token 的获取 / 缓存 / 自动刷新
  2. 为所有飞书 API 调用提供统一的请求头（Authorization + Content-Type）
  3. 封装通用 HTTP 工具（GET / POST / PATCH），统一异常处理

使用方式：
  auth = FeishuAuth(app_id="cli_xxx", app_secret="xxx")
  headers = auth.get_headers()          # 适用于 requests.get/post
  resp = auth.get("/open-apis/xxx")     # 带自动鉴权的 GET
  resp = auth.post("/open-apis/xxx", json={})  # 带自动鉴权的 POST

依赖：requests
"""

from __future__ import annotations

import os
import time
import logging
from dataclasses import dataclass, field
from typing import Any

import requests

logger = logging.getLogger(__name__)

# 飞书 Token 接口
_TOKEN_URL = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"

# Token 默认提前刷新秒数（2 小时有效期 - 300 秒 = 安全刷新窗口）
_REFRESH_BUFFER = 300


# ---------------------------------------------------------------------------
# 异常定义
# ---------------------------------------------------------------------------

class FeishuAuthError(Exception):
    """鉴权相关异常。"""


class FeishuAPIError(Exception):
    """飞书 API 通用业务异常。"""

    def __init__(self, code: int, msg: str, request_id: str = ""):
        self.code = code
        self.msg = msg
        self.request_id = request_id
        super().__init__(f"[code={code}] {msg} (request_id={request_id})")


# ---------------------------------------------------------------------------
# Token 缓存数据
# ---------------------------------------------------------------------------

@dataclass
class _TokenCache:
    token: str = ""
    expire_at: float = 0.0   # Unix 时间戳，token 过期时刻


# ---------------------------------------------------------------------------
# 核心类
# ---------------------------------------------------------------------------

class FeishuAuth:
    """
    飞书应用鉴权管理器。

    Parameters
    ----------
    app_id : str
        飞书应用 App ID（即 "cli_xxx" 开头的凭证）。
    app_secret : str
        飞书应用 App Secret。
    timeout : int
        HTTP 请求超时秒数。
    """

    def __init__(
        self,
        app_id: str | None = None,
        app_secret: str | None = None,
        timeout: int = 30,
    ):
        self.app_id = app_id or os.getenv("FEISHU_APP_ID", "")
        self.app_secret = app_secret or os.getenv("FEISHU_APP_SECRET", "")
        self.timeout = timeout
        self._cache = _TokenCache()

        if not self.app_id or not self.app_secret:
            raise FeishuAuthError(
                "缺少飞书应用凭证。请通过参数传入或在环境变量中设置：\n"
                "  FEISHU_APP_ID=cli_xxx  FEISHU_APP_SECRET=xxx"
            )

    # ------------------------------------------------------------------
    # Token 管理
    # ------------------------------------------------------------------

    def _fetch_token(self) -> str:
        """向飞书服务器请求新的 tenant_access_token。"""
        payload = {
            "app_id": self.app_id,
            "app_secret": self.app_secret,
        }
        try:
            resp = requests.post(
                _TOKEN_URL,
                json=payload,
                timeout=self.timeout,
            )
            resp.raise_for_status()
        except requests.RequestException as exc:
            raise FeishuAuthError(f"请求 Token 失败（网络层）：{exc}") from exc

        data = resp.json()
        if data.get("code") != 0:
            raise FeishuAuthError(
                f"获取 Token 失败：code={data.get('code')}, msg={data.get('msg')}"
            )

        token: str = data["tenant_access_token"]
        expire: int = data.get("expire", 7200)
        self._cache.token = token
        self._cache.expire_at = time.time() + expire - _REFRESH_BUFFER

        logger.info("tenant_access_token 已刷新，有效期 %d 秒", expire)
        return token

    def get_token(self) -> str:
        """获取有效的 token，过期自动刷新。"""
        if not self._cache.token or time.time() >= self._cache.expire_at:
            return self._fetch_token()
        return self._cache.token

    def get_headers(self, content_type: str = "application/json; charset=utf-8") -> dict[str, str]:
        """构造统一的鉴权请求头。"""
        return {
            "Authorization": f"Bearer {self.get_token()}",
            "Content-Type": content_type,
        }

    # ------------------------------------------------------------------
    # 通用 HTTP 封装
    # ------------------------------------------------------------------

    @staticmethod
    def _check_response(resp: requests.Response) -> dict[str, Any]:
        """校验飞书 API 业务响应，成功返回 data 字典，失败抛 FeishuAPIError。"""
        body = resp.json()
        if body.get("code") != 0:
            raise FeishuAPIError(
                code=body.get("code", -1),
                msg=body.get("msg", "未知错误"),
                request_id=body.get("request_id", ""),
            )
        return body.get("data", {})

    def get(self, path: str, params: dict | None = None) -> dict[str, Any]:
        """带鉴权的 GET 请求，返回 data 字段。"""
        url = f"https://open.feishu.cn{path}"
        resp = requests.get(
            url,
            headers=self.get_headers(),
            params=params,
            timeout=self.timeout,
        )
        return self._check_response(resp)

    def post(self, path: str, json: dict | None = None) -> dict[str, Any]:
        """带鉴权的 POST（JSON），返回 data 字段。"""
        url = f"https://open.feishu.cn{path}"
        resp = requests.post(
            url,
            headers=self.get_headers(),
            json=json,
            timeout=self.timeout,
        )
        return self._check_response(resp)

    def post_multipart(
        self,
        path: str,
        files: dict | list | None = None,
        data: dict | None = None,
    ) -> dict[str, Any]:
        """带鉴权的 POST（multipart/form-data），用于文件上传场景。"""
        url = f"https://open.feishu.cn{path}"
        # multipart/form-data 由 requests 自动设置，只保留 Authorization
        headers = {"Authorization": f"Bearer {self.get_token()}"}
        resp = requests.post(
            url,
            headers=headers,
            files=files,
            data=data,
            timeout=self.timeout,
        )
        return self._check_response(resp)

    def patch(self, path: str, json: dict | None = None) -> dict[str, Any]:
        """带鉴权的 PATCH（JSON），用于更新消息卡片等场景。"""
        url = f"https://open.feishu.cn{path}"
        resp = requests.patch(
            url,
            headers=self.get_headers(),
            json=json,
            timeout=self.timeout,
        )
        return self._check_response(resp)

    # ------------------------------------------------------------------
    # CLI 自验证
    # ------------------------------------------------------------------

    def ping(self) -> dict[str, Any]:
        """验证鉴权是否正常（调用一次 token 接口并返回结果）。"""
        token = self.get_token()
        return {"token_prefix": token[:10] + "...", "ok": True}


# ---------------------------------------------------------------------------
# 快捷入口（单例工厂）
# ---------------------------------------------------------------------------

_auth_instance: FeishuAuth | None = None


def get_auth() -> FeishuAuth:
    """获取全局 FeishuAuth 单例（需提前调用 init_auth 或设置环境变量）。"""
    global _auth_instance
    if _auth_instance is None:
        _auth_instance = FeishuAuth()
    return _auth_instance


def init_auth(app_id: str, app_secret: str, timeout: int = 30) -> FeishuAuth:
    """初始化全局单例并返回。"""
    global _auth_instance
    _auth_instance = FeishuAuth(app_id=app_id, app_secret=app_secret, timeout=timeout)
    return _auth_instance


if __name__ == "__main__":
    # 自验证：python feishu/auth.py
    import sys
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    try:
        auth = get_auth()
        result = auth.ping()
        print(f"[PASS] 鉴权验证成功，token 前缀：{result['token_prefix']}")
    except FeishuAuthError as e:
        print(f"[FAIL] {e}", file=sys.stderr)
        print("请设置环境变量后重试：", file=sys.stderr)
        print("  set FEISHU_APP_ID=cli_xxx", file=sys.stderr)
        print("  set FEISHU_APP_SECRET=xxx", file=sys.stderr)
        sys.exit(1)
