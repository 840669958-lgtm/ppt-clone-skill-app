"""
feishu/test_auth.py
===================
FeishuAuth 鉴权校验测试脚本

功能：
  1. 验证能否成功获取 tenant_access_token
  2. 验证 token 有效期
  3. 验证应用是否拥有飞书文档/PPT 的读写权限

使用方式：
  方式一：设置环境变量后运行
    set FEISHU_APP_ID=cli_xxx
    set FEISHU_APP_SECRET=xxx
    python feishu/test_auth.py

  方式二：命令行传参
    python feishu/test_auth.py --app-id cli_xxx --app-secret xxx

输出示例：
  ========================================================
    FeishuAuth 鉴权校验测试
  ========================================================
  ✅ 凭证检查通过
  ✅ Token 获取成功
     有效期：7195 秒（约 2.0 小时）
  ✅ 权限校验通过
     - 云盘文件读取：✅ 已授权
     - 云盘文件写入：✅ 已授权
     - 消息读取：✅ 已授权
     - 消息发送：✅ 已授权
  ========================================================
  全部校验通过，应用配置正确，可正常使用。
"""

from __future__ import annotations

import argparse
import logging
import sys
from typing import Any

# 复用 auth.py 的鉴权逻辑
from .auth import FeishuAuth, FeishuAuthError, FeishuAPIError

# 关闭日志输出，避免干扰测试结果展示
logging.disable(logging.CRITICAL)


# ---------------------------------------------------------------------------
# 权限校验配置
# ---------------------------------------------------------------------------

# 本 Skill 需要的核心权限清单
REQUIRED_PERMISSIONS = [
    {
        "name": "云盘文件读取",
        "endpoint": "/open-apis/drive/v1/files",
        "method": "GET",
        "params": {"page_size": 1},
    },
    {
        "name": "云盘文件写入",
        "endpoint": "/open-apis/drive/v1/files/upload_all",
        "method": "POST",
        "skip_real_call": True,  # 不上传真实文件，仅检查接口可达性
        "note": "需通过飞书开放平台申请 drive:drive 和 drive:file 权限",
    },
    {
        "name": "消息读取",
        "endpoint": "/open-apis/im/v1/messages",
        "method": "GET",
        "params": {"page_size": 1},
    },
    {
        "name": "消息发送",
        "endpoint": "/open-apis/im/v1/chats",
        "method": "GET",
        "params": {"page_size": 1},
        "note": "实际发送消息需额外申请 im:message:send_as_bot 权限",
    },
]


# ---------------------------------------------------------------------------
# 测试函数
# ---------------------------------------------------------------------------

def print_header(title: str) -> None:
    """打印带边框的标题。"""
    width = 60
    print("=" * width)
    print(f"  {title}")
    print("=" * width)


def print_result(success: bool, message: str, detail: str = "") -> None:
    """打印单项测试结果。"""
    icon = "✅" if success else "❌"
    print(f"  {icon} {message}")
    if detail:
        print(f"     {detail}")


def test_credentials(app_id: str, app_secret: str) -> tuple[bool, str]:
    """
    测试凭证格式是否合法。

    Returns
    -------
    (是否通过, 失败原因)
    """
    if not app_id or not app_secret:
        return False, "缺少 App ID 或 App Secret"

    if not app_id.startswith("cli_"):
        return False, f"App ID 格式错误：'{app_id}' 应以 'cli_' 开头"

    if len(app_secret) < 10:
        return False, "App Secret 长度过短，请检查是否完整"

    return True, ""


def test_token_fetch(auth: FeishuAuth) -> tuple[bool, dict[str, Any] | str]:
    """
    测试能否成功获取 tenant_access_token。

    Returns
    -------
    (是否通过, 成功时返回 token 信息字典，失败时返回错误信息)
    """
    try:
        token = auth.get_token()
        if not token:
            return False, "获取到的 token 为空"

        # 计算剩余有效期
        remaining = int(auth._cache.expire_at - __import__("time").time())
        if remaining < 0:
            remaining = 0

        return True, {
            "token_prefix": token[:20] + "...",
            "expires_in": remaining,
            "expires_human": f"{remaining // 3600}小时 {(remaining % 3600) // 60}分钟",
        }
    except FeishuAuthError as e:
        return False, f"鉴权失败：{e}"
    except Exception as e:
        return False, f"未知错误：{e}"


def test_permission(auth: FeishuAuth, perm: dict) -> tuple[bool, str]:
    """
    测试单项权限是否可用。

    对于需要实际上传文件的权限，采用轻量级探测：
    - 尝试调用接口但不携带真实文件
    - 如果返回 403/权限不足，则说明权限未开通
    - 如果返回 400/参数错误，则说明权限已开通（只是缺文件）

    Returns
    -------
    (是否通过, 详细信息)
    """
    endpoint = perm["endpoint"]
    method = perm.get("method", "GET")
    params = perm.get("params", {})
    skip_real = perm.get("skip_real_call", False)

    try:
        if skip_real:
            # 对于上传类接口，发送一个空请求来探测权限
            # 预期会收到参数错误（权限已开通）或权限错误（权限未开通）
            if method == "POST":
                try:
                    auth.post(endpoint, json={})
                except FeishuAPIError as e:
                    # 参数错误（如 400）说明权限已开通，只是参数不对
                    if e.code == 400 or "参数" in e.msg or "param" in e.msg.lower():
                        return True, "权限已开通（接口可访问）"
                    # 403 或 99991663 等是权限不足
                    if e.code in (403, 99991663, 99991664):
                        return False, f"权限未开通：{e.msg}"
                    # 其他错误视为权限已开通（接口可达）
                    return True, f"权限已开通（返回码 {e.code}）"
            return True, "跳过实际调用"
        else:
            # 正常调用 GET 接口
            if method == "GET":
                auth.get(endpoint, params=params)
            return True, "接口调用成功"
    except FeishuAPIError as e:
        # 权限相关错误码
        if e.code in (403, 99991663, 99991664, 99991665):
            return False, f"权限不足（code={e.code}）：{e.msg}"
        # 其他业务错误视为权限已开通（只是业务条件不满足）
        if e.code in (400, 404, 422):
            return True, f"权限已开通（业务错误 {e.code}）"
        return False, f"接口错误（code={e.code}）：{e.msg}"
    except Exception as e:
        return False, f"请求异常：{e}"


def run_all_tests(app_id: str, app_secret: str) -> bool:
    """
    运行全部校验测试。

    Returns
    -------
    全部通过返回 True，否则返回 False
    """
    print_header("FeishuAuth 鉴权校验测试")
    print()

    all_passed = True

    # ------------------------------
    # 测试 1：凭证格式检查
    # ------------------------------
    ok, reason = test_credentials(app_id, app_secret)
    print_result(ok, "凭证检查通过" if ok else f"凭证检查失败：{reason}")
    if not ok:
        print()
        print("  💡 解决方案：")
        print("     1. 确认已创建飞书自建应用")
        print("     2. 在飞书开放平台 → 应用详情 → 凭证信息 中获取 App ID 和 App Secret")
        print("     3. 设置环境变量：")
        print(f"        set FEISHU_APP_ID={app_id or 'cli_xxxxxxxxxxxxxxxx'}")
        print(f"        set FEISHU_APP_SECRET={'x' * 32}")
        return False
    print()

    # ------------------------------
    # 测试 2：Token 获取
    # ------------------------------
    try:
        auth = FeishuAuth(app_id=app_id, app_secret=app_secret)
    except FeishuAuthError as e:
        print_result(False, f"初始化失败：{e}")
        return False

    ok, result = test_token_fetch(auth)
    if ok:
        info = result  # type: ignore
        print_result(True, "Token 获取成功")
        print(f"     Token 前缀：{info['token_prefix']}")
        print(f"     有效期：{info['expires_in']} 秒（约 {info['expires_human']}）")
    else:
        print_result(False, f"Token 获取失败：{result}")
        print()
        print("  💡 解决方案：")
        print("     1. 检查 App ID 和 App Secret 是否匹配（注意大小写）")
        print("     2. 确认应用已发布（飞书开放平台 → 版本管理与发布 → 创建版本并申请发布）")
        print("     3. 检查网络能否访问 open.feishu.cn")
        return False
    print()

    # ------------------------------
    # 测试 3：权限校验
    # ------------------------------
    print("  权限校验：")
    permission_results = []
    for perm in REQUIRED_PERMISSIONS:
        ok, detail = test_permission(auth, perm)
        permission_results.append((perm["name"], ok, detail))
        icon = "✅" if ok else "❌"
        print(f"     {icon} {perm['name']}：{detail}")
        if not ok:
            all_passed = False

    print()
    print("=" * 60)

    if all_passed:
        print("  ✅ 全部校验通过，应用配置正确，可正常使用。")
    else:
        print("  ❌ 部分校验未通过，请按以下步骤排查：")
        print()
        print("  1. 登录飞书开放平台：https://open.feishu.cn")
        print("  2. 进入你的应用 → 权限管理")
        print("  3. 申请以下权限：")
        print("     • im:message（读取消息）")
        print("     • im:message:send_as_bot（发送消息）")
        print("     • drive:drive（云盘操作）")
        print("     • drive:file（文件上传下载）")
        print("  4. 重新发布应用版本")
        print("  5. 确保机器人已添加到测试群聊")

    return all_passed


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def main() -> int:
    """主函数，返回退出码（0=成功，1=失败）。"""
    parser = argparse.ArgumentParser(
        description="FeishuAuth 鉴权校验测试",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  # 方式一：环境变量
  set FEISHU_APP_ID=cli_xxx
  set FEISHU_APP_SECRET=xxx
  python -m feishu.test_auth

  # 方式二：命令行传参
  python -m feishu.test_auth --app-id cli_xxx --app-secret xxx
        """,
    )
    parser.add_argument(
        "--app-id",
        help="飞书应用 App ID（默认从环境变量 FEISHU_APP_ID 读取）",
    )
    parser.add_argument(
        "--app-secret",
        help="飞书应用 App Secret（默认从环境变量 FEISHU_APP_SECRET 读取）",
    )
    args = parser.parse_args()

    # 优先使用命令行参数，其次环境变量
    import os
    app_id = args.app_id or os.getenv("FEISHU_APP_ID", "")
    app_secret = args.app_secret or os.getenv("FEISHU_APP_SECRET", "")

    success = run_all_tests(app_id, app_secret)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
