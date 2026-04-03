# -*- coding: utf-8 -*-
"""
飞书消息处理模块 - Work Buddy Skill集成版

负责：
1. 接收Work Buddy转发的飞书对话消息
2. 解析用户指令、提取PPT链接和参数
3. 调用核心流程进行PPT复刻
4. 返回处理结果给Work Buddy

API文档：https://open.feishu.cn/document/server-docs/im-v1/message/events
"""

import re
import json
import time
import threading
from typing import Optional, Dict, Any, List, Callable, Tuple
from dataclasses import dataclass, field, asdict
from enum import Enum
from urllib.parse import urlparse
from pathlib import Path

# 导入核心模块
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from feishu.auth import get_auth, FeishuAPIError, init_auth
from feishu.file_manager import FeishuFileManager, FileManagerError
from core.ppt_analyzer import analyze, AnalysisError
from core.ppt_builder import PPTBuilder, BuildOptions, PPTBuilderError


# ==================== 异常定义 ====================

class MessageError(Exception):
    """消息处理基础异常"""
    pass


class InvalidMessageError(MessageError):
    """消息格式无效"""
    pass


class MessageSendError(MessageError):
    """消息发送失败"""
    pass


class ParameterParseError(MessageError):
    """参数解析失败"""
    pass


# ==================== 数据模型 ====================

class MessageType(Enum):
    """飞书消息类型"""
    TEXT = "text"
    POST = "post"           # 富文本
    IMAGE = "image"
    FILE = "file"
    MERGE_FORWARD = "merge_forward"
    CARD = "interactive"    # 卡片消息


@dataclass
class PPTShareMessage:
    """飞书PPT分享消息数据模型"""
    message_id: str
    chat_id: str
    chat_type: str          # group/p2p
    sender_open_id: str
    sender_name: str
    ppt_url: str
    file_name: Optional[str] = None
    file_token: Optional[str] = None
    receive_time: float = field(default_factory=time.time)
    
    # Work Buddy上下文
    session_id: Optional[str] = None
    user_id: Optional[str] = None
    
    def __repr__(self) -> str:
        return f"PPTShareMessage(id={self.message_id[:20]}..., sender={self.sender_name}, file={self.file_name or 'Unknown'})"


@dataclass
class CloneParameters:
    """PPT复刻参数"""
    slide_count: Optional[int] = None      # 页面数
    primary_color: Optional[str] = None    # 主色调（十六进制）
    replace_logo: Optional[str] = None     # 替换Logo路径
    keep_animation: bool = False           # 是否保留动画
    output_name: Optional[str] = None      # 输出文件名
    
    @classmethod
    def from_text(cls, text: str) -> "CloneParameters":
        """
        从用户文本中解析参数
        
        支持的参数格式：
        - 页面数: --pages 15 / -p 15 / 15页
        - 主色调: --color FF5500 / -c FF5500 / 橙色主题
        - 保留动画: --keep-animation / 保留动画
        
        Args:
            text: 用户输入文本
            
        Returns:
            CloneParameters对象
        """
        params = cls()
        text_lower = text.lower()
        
        # 解析页面数
        # 格式: --pages 15, -p 15, 15页, 15页左右
        page_patterns = [
            r'--pages?\s+(\d+)',
            r'-p\s+(\d+)',
            r'(\d+)\s*页',
        ]
        for pattern in page_patterns:
            match = re.search(pattern, text_lower)
            if match:
                params.slide_count = int(match.group(1))
                break
        
        # 解析主色调
        # 格式: --color FF5500, -c FF5500, 橙色主题
        color_patterns = [
            r'--colou?r?\s+([0-9a-fA-F]{6})',
            r'-c\s+([0-9a-fA-F]{6})',
        ]
        for pattern in color_patterns:
            match = re.search(pattern, text)
            if match:
                params.primary_color = match.group(1).upper()
                break
        
        # 颜色关键词映射
        color_keywords = {
            '红色': 'FF0000', '绿色': '00FF00', '蓝色': '0000FF',
            '橙色': 'FF8800', '黄色': 'FFFF00', '紫色': '8800FF',
            '粉色': 'FF88FF', '青色': '00FFFF', '黑色': '000000',
            '白色': 'FFFFFF', '灰色': '888888',
        }
        if not params.primary_color:
            for keyword, hex_color in color_keywords.items():
                if keyword in text:
                    params.primary_color = hex_color
                    break
        
        # 解析保留动画
        if '--keep-animation' in text_lower or '保留动画' in text:
            params.keep_animation = True
        
        # 解析输出文件名
        output_match = re.search(r'--output\s+["\']?([^"\'\s]+)', text)
        if output_match:
            params.output_name = output_match.group(1)
        
        return params


@dataclass
class CloneResult:
    """PPT复刻结果"""
    success: bool
    original_url: str
    original_name: str
    new_url: Optional[str] = None
    new_file_token: Optional[str] = None
    page_count: int = 0
    duration: float = 0.0
    error_message: Optional[str] = None
    suggestion: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return asdict(self)


# ==================== URL解析工具 ====================

# 飞书PPT链接正则表达式
FEISHU_PPT_URL_PATTERNS = [
    # 标准分享链接
    r'https?://[\w\-]+\.feishu\.cn/slides/([a-zA-Z0-9]+)',
    # 旧版链接
    r'https?://[\w\-]+\.feishu\.cn/file/([a-zA-Z0-9]+)',
    # 短链接
    r'https?://[\w\-]+\.feishu\.cn/s/([a-zA-Z0-9]+)',
]


def extract_ppt_urls(text: str) -> List[str]:
    """
    从文本中提取所有飞书PPT链接
    
    Args:
        text: 待解析的文本内容
        
    Returns:
        提取到的PPT链接列表（去重）
    """
    urls = []
    for pattern in FEISHU_PPT_URL_PATTERNS:
        matches = re.findall(pattern, text)
        for match in matches:
            # 重构完整URL
            if 'slides' in pattern:
                url = f"https://www.feishu.cn/slides/{match}"
            elif 'file' in pattern:
                url = f"https://www.feishu.cn/file/{match}"
            else:
                url = f"https://www.feishu.cn/s/{match}"
            urls.append(url)
    
    # 去重保持顺序
    seen = set()
    unique_urls = []
    for url in urls:
        if url not in seen:
            seen.add(url)
            unique_urls.append(url)
    
    return unique_urls


def is_feishu_ppt_url(url: str) -> bool:
    """
    检查URL是否为飞书PPT链接
    
    Args:
        url: 待检查的URL
        
    Returns:
        是否为飞书PPT链接
    """
    for pattern in FEISHU_PPT_URL_PATTERNS:
        if re.match(pattern, url):
            return True
    return False


# ==================== Work Buddy Skill处理器 ====================

class PPTCloneSkill:
    """
    PPT复刻Skill - Work Buddy集成版
    
    完整处理流程：
    1. 接收用户消息
    2. 解析PPT链接和参数
    3. 下载 → 解析 → 重建 → 上传
    4. 返回结果
    """
    
    def __init__(self, download_dir: str = "./downloads", output_dir: str = "./output"):
        """
        初始化Skill
        
        Args:
            download_dir: 下载文件保存目录
            output_dir: 生成文件输出目录
        """
        self.download_dir = download_dir
        self.output_dir = output_dir
        
        # 确保目录存在
        Path(download_dir).mkdir(parents=True, exist_ok=True)
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        # 初始化组件（延迟初始化认证）
        self._file_manager: Optional[FeishuFileManager] = None
        self._auth = None
    
    @property
    def file_manager(self) -> FeishuFileManager:
        """懒加载文件管理器"""
        if self._file_manager is None:
            self._file_manager = FeishuFileManager(download_dir=self.download_dir)
        return self._file_manager
    
    def handle_message(self, message_text: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """
        处理Work Buddy转发的消息
        
        Args:
            message_text: 用户消息文本
            context: Work Buddy上下文信息
                {
                    "chat_id": "oc_xxx",
                    "chat_type": "group",
                    "sender_id": "ou_xxx",
                    "sender_name": "用户名",
                    "message_id": "om_xxx",
                    "session_id": "session_xxx"
                }
        
        Returns:
            处理结果字典
        """
        start_time = time.time()
        
        # 1. 提取PPT链接
        ppt_urls = extract_ppt_urls(message_text)
        if not ppt_urls:
            return {
                "success": False,
                "error": "未检测到飞书PPT链接",
                "suggestion": "请分享一个飞书PPT链接，例如：https://xxx.feishu.cn/slides/xxx",
                "requires_user_action": True
            }
        
        ppt_url = ppt_urls[0]
        
        # 2. 解析参数
        params = CloneParameters.from_text(message_text)
        
        # 3. 创建消息对象
        ppt_msg = PPTShareMessage(
            message_id=context.get("message_id", f"wb_{int(time.time())}"),
            chat_id=context.get("chat_id", ""),
            chat_type=context.get("chat_type", "p2p"),
            sender_open_id=context.get("sender_id", ""),
            sender_name=context.get("sender_name", "用户"),
            ppt_url=ppt_url,
            session_id=context.get("session_id")
        )
        
        # 4. 执行复刻流程
        result = self._execute_clone(ppt_msg, params)
        result.duration = time.time() - start_time
        
        return result.to_dict()
    
    def _execute_clone(self, ppt_msg: PPTShareMessage, params: CloneParameters) -> CloneResult:
        """
        执行PPT复刻流程
        
        Args:
            ppt_msg: PPT消息对象
            params: 复刻参数
            
        Returns:
            CloneResult对象
        """
        result = CloneResult(
            success=False,
            original_url=ppt_msg.ppt_url,
            original_name="Unknown"
        )
        
        try:
            # 1. 获取文件信息
            print(f"[1/5] 获取文件信息...")
            file_info = self.file_manager.get_ppt_info(ppt_msg.ppt_url)
            result.original_name = file_info.file_name
            print(f"      文件名: {file_info.file_name}")
            
            # 2. 下载PPT
            print(f"[2/5] 下载PPT文件...")
            download_result = self.file_manager.download_ppt(ppt_msg.ppt_url)
            local_path = download_result.file_path
            print(f"      已下载到: {local_path}")
            
            # 3. 解析模板
            print(f"[3/5] 解析模板特征...")
            profile = analyze(local_path)
            print(f"      页面尺寸: {profile.geometry.width_cm:.1f} x {profile.geometry.height_cm:.1f} cm")
            print(f"      配色方案: {len(profile.theme_colors)} 种主题色")
            print(f"      可用版式: {len(profile.all_layouts)} 种")
            
            # 4. 重建PPT
            print(f"[4/5] 重建PPT...")
            builder = PPTBuilder(profile, output_dir=self.output_dir)
            
            # 确定页面数
            slide_count = params.slide_count or 10
            
            build_options = BuildOptions(
                slide_count=slide_count,
                primary_color=params.primary_color,
                replace_logo_path=params.replace_logo,
                preserve_animations=params.keep_animation
            )
            
            build_result = builder.build(build_options)
            print(f"      生成完成: {build_result.output_path}")
            
            # 5. 上传到飞书
            print(f"[5/5] 上传到飞书...")
            upload_result = self.file_manager.upload_ppt(build_result.output_path)
            print(f"      上传成功: {upload_result.feishu_url}")
            
            # 更新结果
            result.success = True
            result.new_url = upload_result.feishu_url
            result.new_file_token = upload_result.file_token
            result.page_count = slide_count
            
        except FileManagerError as e:
            result.error_message = str(e)
            if "权限" in str(e) or "Permission" in str(e):
                result.suggestion = "请检查飞书应用是否已开通所需权限: drive:file:download, drive:file:upload"
            elif "不存在" in str(e) or "NotFound" in str(e):
                result.suggestion = "请检查PPT链接是否有效，文件是否已被删除或移动"
            else:
                result.suggestion = "文件操作失败，请稍后重试或联系管理员"
                
        except AnalysisError as e:
            result.error_message = f"PPT解析失败: {str(e)}"
            result.suggestion = "请确保分享的是有效的PPT文件（.pptx格式）"
            
        except PPTBuilderError as e:
            result.error_message = f"PPT生成失败: {str(e)}"
            result.suggestion = "模板重建过程出错，请尝试使用其他PPT文件"
            
        except Exception as e:
            result.error_message = f"处理失败: {str(e)}"
            result.suggestion = "未知错误，请检查日志或联系管理员"
        
        return result
    
    def build_reply_message(self, result: CloneResult) -> str:
        """
        构建回复消息
        
        Args:
            result: 复刻结果
            
        Returns:
            回复文本
        """
        if result.success:
            return f"""✅ PPT复刻完成！

📄 原文件：{result.original_name}
📊 页面数：{result.page_count} 页
⏱️ 耗时：{result.duration:.1f} 秒

🔗 复刻后的PPT：{result.new_url}

💡 提示：点击链接即可在飞书中编辑新PPT"""
        else:
            reply = f"""❌ PPT复刻失败

错误信息：{result.error_message}"""
            if result.suggestion:
                reply += f"\n\n💡 建议：{result.suggestion}"
            return reply
    
    def build_reply_card(self, result: CloneResult) -> Dict[str, Any]:
        """
        构建回复卡片（富文本）
        
        Args:
            result: 复刻结果
            
        Returns:
            卡片JSON对象
        """
        if result.success:
            return {
                "config": {"wide_screen_mode": True},
                "header": {
                    "title": {"tag": "plain_text", "content": "✅ PPT复刻完成"},
                    "template": "green"
                },
                "elements": [
                    {"tag": "div", "text": {"tag": "lark_md", "content": f"**原文件：**{result.original_name}"}},
                    {"tag": "div", "text": {"tag": "lark_md", "content": f"**页面数：**{result.page_count} 页"}},
                    {"tag": "div", "text": {"tag": "lark_md", "content": f"**耗时：**{result.duration:.1f} 秒"}},
                    {"tag": "hr"},
                    {
                        "tag": "action",
                        "actions": [
                            {
                                "tag": "button",
                                "text": {"tag": "plain_text", "content": "📎 打开复刻后的PPT"},
                                "type": "primary",
                                "url": result.new_url
                            }
                        ]
                    }
                ]
            }
        else:
            elements = [
                {"tag": "div", "text": {"tag": "lark_md", "content": f"**错误：**{result.error_message}"}}
            ]
            if result.suggestion:
                elements.append({"tag": "hr"})
                elements.append({"tag": "div", "text": {"tag": "lark_md", "content": f"💡 **建议：**{result.suggestion}"}})
            
            return {
                "config": {"wide_screen_mode": True},
                "header": {
                    "title": {"tag": "plain_text", "content": "❌ PPT复刻失败"},
                    "template": "red"
                },
                "elements": elements
            }


# ==================== 便捷函数（供Work Buddy调用） ====================

_skill_instance: Optional[PPTCloneSkill] = None


def get_skill() -> PPTCloneSkill:
    """获取Skill单例"""
    global _skill_instance
    if _skill_instance is None:
        _skill_instance = PPTCloneSkill()
    return _skill_instance


def init_skill(download_dir: str = "./downloads", output_dir: str = "./output") -> PPTCloneSkill:
    """
    初始化Skill
    
    Args:
        download_dir: 下载目录
        output_dir: 输出目录
        
    Returns:
        PPTCloneSkill实例
    """
    global _skill_instance
    _skill_instance = PPTCloneSkill(download_dir=download_dir, output_dir=output_dir)
    return _skill_instance


def process_ppt_clone_request(message_text: str, context: Dict[str, Any]) -> Dict[str, Any]:
    """
    处理PPT复刻请求（Work Buddy入口函数）
    
    Args:
        message_text: 用户消息文本
        context: Work Buddy上下文
        
    Returns:
        处理结果
    """
    skill = get_skill()
    return skill.handle_message(message_text, context)


def build_success_reply(result: Dict[str, Any]) -> str:
    """
    构建成功回复（文本格式）
    
    Args:
        result: 处理结果字典
        
    Returns:
        回复文本
    """
    if result.get("success"):
        return f"""✅ PPT复刻完成！

📄 原文件：{result.get('original_name')}
📊 页面数：{result.get('page_count')} 页
⏱️ 耗时：{result.get('duration', 0):.1f} 秒

🔗 复刻后的PPT：{result.get('new_url')}

💡 提示：点击链接即可在飞书中编辑新PPT"""
    else:
        reply = f"""❌ PPT复刻失败

错误信息：{result.get('error_message', '未知错误')}"""
        if result.get('suggestion'):
            reply += f"\n\n💡 建议：{result.get('suggestion')}"
        return reply


# ==================== 测试入口 ====================

if __name__ == "__main__":
    """
    测试入口：验证Skill功能
    
    运行方式:
        # 1. 设置环境变量
        set FEISHU_APP_ID=cli_xxx
        set FEISHU_APP_SECRET=xxx
        
        # 2. 运行测试
        python -m feishu.message
    """
    print("=" * 60)
    print("PPT复刻Skill - 测试入口")
    print("=" * 60)
    
    # 测试1: 参数解析
    print("\n【测试1】参数解析")
    test_cases = [
        ("复刻这个PPT https://test.feishu.cn/slides/abc123 --pages 15", 15, None),
        ("复刻 https://test.feishu.cn/slides/abc123 20页 蓝色主题", 20, "0000FF"),
        ("复刻PPT -p 10 --color FF5500", 10, "FF5500"),
        ("复刻PPT https://test.feishu.cn/slides/abc123", None, None),
    ]
    
    for text, expected_pages, expected_color in test_cases:
        params = CloneParameters.from_text(text)
        pages_ok = params.slide_count == expected_pages
        color_ok = params.primary_color == expected_color
        status = "✅" if pages_ok and color_ok else "❌"
        print(f"{status} 输入: {text[:50]}...")
        print(f"   页面数: {params.slide_count} (期望: {expected_pages})")
        print(f"   主色调: {params.primary_color} (期望: {expected_color})")
    
    # 测试2: URL提取
    print("\n【测试2】URL提取")
    url_tests = [
        "分享PPT: https://abc.feishu.cn/slides/sldcnABC123",
        "多个链接: https://a.feishu.cn/slides/sldcn1 和 https://b.feishu.cn/slides/sldcn2",
        "没有链接的消息",
    ]
    
    for text in url_tests:
        urls = extract_ppt_urls(text)
        status = "✅" if urls else "⚠️"
        print(f"{status} 提取到 {len(urls)} 个链接: {urls}")
    
    # 测试3: Skill初始化（需要凭证）
    print("\n【测试3】Skill初始化")
    try:
        skill = init_skill()
        print("✅ Skill初始化成功")
    except Exception as e:
        print(f"⚠️ Skill初始化失败（可能需要配置凭证）: {e}")
    
    # 测试4: 完整流程（模拟）
    print("\n【测试4】消息处理流程（模拟）")
    
    test_message = "帮我复刻这个PPT模板 https://test.feishu.cn/slides/sldcnTest123 --pages 12 --color FF5500"
    test_context = {
        "chat_id": "test_chat",
        "chat_type": "p2p",
        "sender_id": "test_user",
        "sender_name": "测试用户",
        "message_id": "test_msg_001"
    }
    
    try:
        result = process_ppt_clone_request(test_message, test_context)
        print(f"处理结果: {json.dumps(result, ensure_ascii=False, indent=2)[:500]}...")
    except Exception as e:
        print(f"⚠️ 流程测试失败（可能需要配置凭证）: {e}")
    
    print("\n" + "=" * 60)
    print("测试完成！")
    print("如需完整测试，请配置飞书凭证后运行")
    print("=" * 60)
