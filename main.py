# -*- coding: utf-8 -*-
"""
PPT复刻助手 - 主入口模块

这是整个应用的串联中枢，负责：
1. 接收飞书消息事件（PPT分享链接）
2. 下载原始PPT文件
3. 解析模板特征（颜色、字体、版式）
4. 重建生成新PPT
5. 上传新PPT到飞书
6. 回复处理结果

运行方式:
    # 开发模式（本地测试）
    python main.py --mode local --url "https://xxx.feishu.cn/slides/xxx"
    
    # Webhook模式（接收飞书事件）
    python main.py --mode webhook --port 8080
    
    # 直接处理本地文件
    python main.py --mode file --input ./template.pptx --output ./result.pptx

环境变量:
    FEISHU_APP_ID: 飞书应用ID
    FEISHU_APP_SECRET: 飞书应用密钥
    FEISHU_ENCRYPT_KEY: 事件订阅加密密钥（可选）
    FEISHU_VERIFICATION_TOKEN: 验证Token（可选）
"""

import os
import sys
import time
import json
import argparse
import tempfile
import base64
import hashlib
from pathlib import Path
from typing import Optional, Dict, Any
from Crypto.Cipher import AES
from dataclasses import dataclass
from datetime import datetime

# 确保项目根目录在路径中
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from feishu.auth import init_auth, get_auth, FeishuAuthError
from feishu.file_manager import FeishuFileManager, FileManagerError
from feishu.message import (
    PPTShareMessage, 
    process_ppt_clone_request,
    MessageError,
    get_skill
)
from core.ppt_analyzer import analyze, AnalysisError
from core.ppt_builder import PPTBuilder, BuildOptions, BuildResult, PPTBuilderError


# ==================== 配置 ====================

@dataclass
class AppConfig:
    """应用配置"""
    app_id: str
    app_secret: str
    encrypt_key: Optional[str] = None
    verification_token: Optional[str] = None
    download_dir: str = "./downloads"
    output_dir: str = "./output"
    default_slide_count: int = 10
    
    @classmethod
    def from_env(cls) -> "AppConfig":
        """从环境变量加载配置"""
        app_id = os.getenv("FEISHU_APP_ID", "")
        app_secret = os.getenv("FEISHU_APP_SECRET", "")
        
        if not app_id or not app_secret:
            raise ValueError(
                "缺少飞书应用凭证。请设置环境变量:\n"
                "  FEISHU_APP_ID=你的应用ID\n"
                "  FEISHU_APP_SECRET=你的应用密钥"
            )
        
        return cls(
            app_id=app_id,
            app_secret=app_secret,
            encrypt_key=os.getenv("FEISHU_ENCRYPT_KEY"),
            verification_token=os.getenv("FEISHU_VERIFICATION_TOKEN"),
            download_dir=os.getenv("DOWNLOAD_DIR", "./downloads"),
            output_dir=os.getenv("OUTPUT_DIR", "./output"),
            default_slide_count=int(os.getenv("DEFAULT_SLIDE_COUNT", "10"))
        )


# ==================== 核心工作流 ====================

class PPTCloneWorkflow:
    """
    PPT复刻工作流
    
    完整的处理流程：下载 → 解析 → 重建 → 上传 → 通知
    """
    
    def __init__(self, config: AppConfig):
        self.config = config
        self.file_manager = FeishuFileManager(download_dir=config.download_dir)
        self.skill = get_skill()
        
        # 确保输出目录存在
        Path(config.output_dir).mkdir(parents=True, exist_ok=True)
        
    def _format_file_size(self, size_bytes: int) -> str:
        """将字节数转换为人类可读的格式。"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / 1024 / 1024:.1f} MB"
        else:
            return f"{size_bytes / 1024 / 1024 / 1024:.1f} GB"
    
    def process_ppt_share(self, ppt_msg: PPTShareMessage, 
                          slide_count: Optional[int] = None,
                          primary_color: Optional[str] = None,
                          replace_logo_path: Optional[str] = None) -> Dict[str, Any]:
        """
        处理PPT分享消息 - 完整工作流
        
        Args:
            ppt_msg: PPT分享消息
            slide_count: 指定页面数（默认使用配置值）
            primary_color: 指定主色调（可选）
            replace_logo_path: 替换Logo路径（可选）
            
        Returns:
            处理结果字典
        """
        start_time = time.time()
        result = {
            "success": False,
            "original_url": ppt_msg.ppt_url,
            "new_url": None,
            "error": None,
            "duration": 0,
            "page_count": 0
        }
        
        try:
            # 1. 发送开始处理通知
            print(f"[1/5] 开始处理: {ppt_msg.ppt_url}")
            
            # 2. 获取文件信息
            print(f"[2/5] 获取文件信息...")
            file_info = self.file_manager.get_ppt_info(ppt_msg.ppt_url)
            print(f"      文件名: {file_info.file_name}")
            print(f"      大小: {self._format_file_size(file_info.size)}")
            
            # 3. 下载PPT文件
            print(f"[3/5] 下载PPT文件...")
            
            download_result = self.file_manager.download_ppt(ppt_msg.ppt_url)
            local_path = download_result.file_path
            print(f"      已下载到: {local_path}")
            
            # 4. 解析模板
            print(f"[4/5] 解析模板特征...")
            
            profile = analyze(local_path)
            print(f"      页面尺寸: {profile.geometry.width_cm:.1f} x {profile.geometry.height_cm:.1f} cm")
            print(f"      配色方案: {len(profile.theme_colors)} 种主题色")
            print(f"      字体方案: {profile.font_scheme.major_latin or '默认'}")
            print(f"      可用版式: {len(profile.all_layouts)} 种")
            
            # 5. 重建PPT
            print(f"[5/5] 重建PPT...")
            
            builder = PPTBuilder(profile, output_dir=self.config.output_dir)
            build_options = BuildOptions(
                slide_count=slide_count or self.config.default_slide_count,
                primary_color=primary_color,
                replace_logo_path=replace_logo_path
            )
            
            build_result = builder.build(build_options)
            print(f"      生成完成: {build_result.output_path}")
            import os
            output_size = os.path.getsize(build_result.output_path)
            print(f"      文件大小: {self._format_file_size(output_size)}")
            
            # 6. 上传到飞书
            print(f"      上传到飞书...")
            
            upload_result = self.file_manager.upload_ppt(build_result.output_path)
            print(f"      上传成功: {upload_result.feishu_url}")
            
            # 7. 发送完成通知
            duration = time.time() - start_time
            
            # 更新结果
            result["success"] = True
            result["new_url"] = upload_result.feishu_url
            result["duration"] = duration
            result["page_count"] = build_options.slide_count
            
            print(f"\n✅ 处理完成！耗时 {duration:.1f} 秒")
            print(f"   新PPT链接: {upload_result.feishu_url}")
            
        except Exception as e:
            duration = time.time() - start_time
            result["duration"] = duration
            result["error"] = str(e)
            
            error_msg = str(e)
            suggestion = ""
            
            # 根据错误类型提供建议
            if "权限" in error_msg or "Permission" in error_msg:
                suggestion = "请检查飞书应用是否已开通所需权限: drive:file:download, drive:file:upload"
            elif "不存在" in error_msg or "NotFound" in error_msg:
                suggestion = "请检查PPT链接是否有效，文件是否已被删除"
            elif "下载" in error_msg:
                suggestion = "文件下载失败，请检查网络连接或稍后重试"
            elif "上传" in error_msg:
                suggestion = "文件上传失败，请检查文件大小是否超过限制（默认100MB）"
            else:
                suggestion = "请检查日志获取详细信息，或联系管理员"
            
            print(f"\n❌ 处理失败: {error_msg}")
            if suggestion:
                print(f"   建议: {suggestion}")
        
        return result
    
    def process_local_file(self, input_path: str, output_path: Optional[str] = None,
                           slide_count: Optional[int] = None,
                           primary_color: Optional[str] = None) -> BuildResult:
        """
        处理本地PPT文件（不经过飞书）
        
        Args:
            input_path: 输入PPT文件路径
            output_path: 输出路径（可选，默认自动生成）
            slide_count: 页面数
            primary_color: 主色调
            
        Returns:
            BuildResult对象
        """
        print(f"处理本地文件: {input_path}")
        
        # 1. 解析模板
        print("  解析模板...")
        profile = analyze(input_path)
        
        # 2. 重建
        print("  重建PPT...")
        builder = PPTBuilder(profile, output_dir=self.config.output_dir)
        
        build_options = BuildOptions(
            slide_count=slide_count or self.config.default_slide_count,
            primary_color=primary_color
        )
        
        result = builder.build(build_options)
        
        # 3. 如果指定了输出路径，复制过去
        if output_path:
            import shutil
            shutil.copy2(result.output_path, output_path)
            result.output_path = output_path
            print(f"  已保存到: {output_path}")
        
        return result


# ==================== Webhook服务器 ====================

def decrypt_feishu_msg(encrypt_msg: str, encrypt_key: str) -> str:
    """
    解密飞书加密消息（AES-256-CBC）
    
    Args:
        encrypt_msg: 加密的消息字符串（base64编码）
        encrypt_key: 加密密钥
        
    Returns:
        解密后的明文JSON字符串
    """
    # 飞书使用 SHA256 哈希密钥取前32字节作为AES密钥
    key = hashlib.sha256(encrypt_key.encode('utf-8')).digest()
    
    # Base64 解码
    encrypted_data = base64.b64decode(encrypt_msg)
    
    # 提取 IV (前 16 字节) 和密文
    iv = encrypted_data[:16]
    ciphertext = encrypted_data[16:]
    
    # AES-CBC 解密
    cipher = AES.new(key, AES.MODE_CBC, iv)
    plaintext = cipher.decrypt(ciphertext)
    
    # 去除 PKCS7 填充
    padding_len = plaintext[-1]
    plaintext = plaintext[:-padding_len]
    
    return plaintext.decode('utf-8')


def create_webhook_app(config: AppConfig) -> Any:
    """
    创建Webhook应用（用于接收飞书事件）
    
    需要安装: pip install flask
    
    Args:
        config: 应用配置
        
    Returns:
        Flask应用实例
    """
    try:
        from flask import Flask, request, jsonify
    except ImportError:
        print("错误: 运行Webhook模式需要安装Flask")
        print("  pip install flask")
        sys.exit(1)
    
    app = Flask(__name__)
    workflow = PPTCloneWorkflow(config)
    
    @app.route("/webhook", methods=["POST"])
    def webhook():
        """接收飞书事件推送"""
        try:
            data = request.get_json()
            print(f"[Webhook] 收到请求: {data}")
            
            # 处理加密数据
            if "encrypt" in data:
                # 有加密，需要解密
                encrypt_key = config.encrypt_key or ""
                if encrypt_key:
                    try:
                        decrypted = decrypt_feishu_msg(data["encrypt"], encrypt_key)
                        print(f"[Webhook] 解密后: {decrypted}")
                        data = json.loads(decrypted)
                    except Exception as e:
                        print(f"[Webhook] 解密失败: {e}")
                        return jsonify({"code": 1, "msg": "decrypt failed"}), 400
                else:
                    print("[Webhook] 收到加密数据但没有配置 encrypt_key")
                    return jsonify({"code": 1, "msg": "no encrypt_key"}), 400
            
            # 处理URL验证（首次配置事件订阅时）
            if data.get("type") == "url_verification":
                challenge = data.get("challenge")
                print(f"[Webhook] 收到URL验证请求, challenge: {challenge}")
                response = {"challenge": challenge}
                print(f"[Webhook] 返回响应: {response}")
                from flask import Response
                return Response(
                    json.dumps(response),
                    status=200,
                    mimetype='application/json'
                )
            
            # 处理消息事件
            event_type = data.get("header", {}).get("event_type", "")
            
            if event_type == "im.message.receive_v1":
                try:
                    # 解析消息内容
                    event_data = data.get("event", {})
                    message = event_data.get("message", {})
                    content = json.loads(message.get("content", "{}"))
                    
                    # 检查是否是PPT分享
                    msg_type = message.get("message_type", "")
                    if msg_type == "share_chat" or "url" in content:
                        print(f"\n[Webhook] 收到PPT分享消息")
                        print(f"  内容: {content}")
                        
                        # TODO: 解析PPT链接并处理
                        # 这里暂时返回处理中，实际应该启动异步任务
                        return jsonify({"code": 0, "msg": "ppt_share_received"})
                    else:
                        return jsonify({"code": 0, "msg": "not_ppt_share"})
                        
                except Exception as e:
                    print(f"[Webhook] 处理消息事件失败: {e}")
                    import traceback
                    traceback.print_exc()
                    return jsonify({"code": -1, "msg": str(e)}), 500
            
            return jsonify({"code": 0, "msg": "ignored"})
        except Exception as e:
            print(f"[Webhook] 处理请求失败: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({"code": -1, "msg": "internal_error"}), 500
    
    @app.errorhandler(404)
    def not_found(error):
        """处理404错误"""
        return jsonify({"code": 404, "msg": "not found"}), 404
    
    @app.errorhandler(500)
    def internal_error(error):
        """处理500错误"""
        return jsonify({"code": 500, "msg": "internal error"}), 500
    
    @app.route("/health", methods=["GET"])
    def health():
        """健康检查"""
        return jsonify({
            "status": "ok",
            "timestamp": datetime.now().isoformat()
        })
    
    return app


def create_webhook_app_from_env():
    """从环境变量创建Webhook应用（供Gunicorn使用）"""
    config = AppConfig.from_env()
    init_auth(app_id=config.app_id, app_secret=config.app_secret)
    return create_webhook_app(config)


# ==================== 命令行入口 ====================

def main():
    """主入口函数"""
    parser = argparse.ArgumentParser(
        description="PPT复刻助手 - 飞书PPT模板智能复刻工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 处理飞书PPT链接（开发测试）
  python main.py --mode local --url "https://xxx.feishu.cn/slides/xxx"
  
  # 启动Webhook服务器
  python main.py --mode webhook --port 8080
  
  # 处理本地PPT文件
  python main.py --mode file --input ./template.pptx --output ./result.pptx --pages 15
        """
    )
    
    parser.add_argument(
        "--mode", "-m",
        choices=["local", "webhook", "file"],
        default="local",
        help="运行模式: local=本地测试, webhook=Webhook服务器, file=本地文件处理 (默认: local)"
    )
    
    # Local模式参数
    parser.add_argument(
        "--url", "-u",
        help="飞书PPT分享链接（local模式）"
    )
    
    # File模式参数
    parser.add_argument(
        "--input", "-i",
        help="输入PPT文件路径（file模式）"
    )
    parser.add_argument(
        "--output", "-o",
        help="输出PPT文件路径（file模式）"
    )
    
    # 通用参数
    parser.add_argument(
        "--pages", "-p",
        type=int,
        help=f"生成PPT的页面数（默认从环境变量DEFAULT_SLIDE_COUNT读取，或10页）"
    )
    parser.add_argument(
        "--color", "-c",
        help="指定主色调（十六进制，如 FF5500）"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8080,
        help="Webhook服务器端口（默认: 8080）"
    )
    
    args = parser.parse_args()
    
    # File模式不需要飞书凭证，直接处理
    if args.mode == "file":
        if not args.input:
            print("错误: file模式需要指定 --input 参数")
            parser.print_help()
            sys.exit(1)
        
        # 确保输出目录存在
        output_dir = Path(os.getenv("OUTPUT_DIR", "./output"))
        output_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"处理本地文件: {args.input}")
        
        try:
            # 1. 解析模板
            print("  解析模板...")
            profile = analyze(args.input)
            print(f"     页面尺寸: {profile.geometry.width_cm:.1f} x {profile.geometry.height_cm:.1f} cm")
            
            # 2. 重建
            print("  重建PPT...")
            builder = PPTBuilder(profile, output_dir=str(output_dir))
            
            from core.ppt_builder import BuildOptions
            build_options = BuildOptions(
                slide_count=args.pages or 10,
                primary_color=args.color
            )
            
            result = builder.build(build_options)
            
            # 3. 如果指定了输出路径，复制过去
            if args.output:
                import shutil
                shutil.copy2(result.output_path, args.output)
                result.output_path = args.output
                print(f"  已保存到: {args.output}")
            
            print(f"\n🎉 处理完成！")
            print(f"   输出文件: {result.output_path}")
            print(f"   页面数: {result.slide_count}")
            
        except Exception as e:
            print(f"\n💥 处理失败: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        
        return  # file模式处理完成，直接返回
    
    # 其他模式需要飞书凭证
    try:
        config = AppConfig.from_env()
    except ValueError as e:
        print(f"配置错误: {e}")
        sys.exit(1)
    
    # 初始化飞书认证
    try:
        init_auth(app_id=config.app_id, app_secret=config.app_secret)
        print("✅ 飞书认证初始化成功")
    except FeishuAuthError as e:
        print(f"❌ 飞书认证失败: {e}")
        sys.exit(1)
    
    # 根据模式执行
    if args.mode == "local":
        if not args.url:
            print("错误: local模式需要指定 --url 参数")
            parser.print_help()
            sys.exit(1)
        
        # 创建模拟消息
        ppt_msg = PPTShareMessage(
            message_id=f"manual_{int(time.time())}",
            chat_id="manual_test",
            sender_open_id="user_manual",
            sender_name="手动测试",
            ppt_url=args.url
        )
        
        workflow = PPTCloneWorkflow(config)
        result = workflow.process_ppt_share(
            ppt_msg,
            slide_count=args.pages,
            primary_color=args.color
        )
        
        if result["success"]:
            print(f"\n🎉 复刻成功！")
            print(f"   新PPT链接: {result['new_url']}")
        else:
            print(f"\n💥 复刻失败: {result['error']}")
            sys.exit(1)
    
    elif args.mode == "webhook":
        app = create_webhook_app(config)
        port = int(os.getenv("PORT", args.port))
        print(f"\n🚀 启动Webhook服务器...")
        print(f"   监听端口: {port}")
        print(f"   Webhook URL: http://0.0.0.0:{port}/webhook")
        print(f"   健康检查: http://0.0.0.0:{port}/health")
        print(f"\n请在飞书开发者后台配置事件订阅地址:")
        print(f"   https://你的域名/webhook")
        print(f"\n按 Ctrl+C 停止服务\n")
        
        app.run(host="0.0.0.0", port=port, debug=False)


# ==================== 测试入口 ====================

if __name__ == "__main__":
    """
    测试入口
    
    运行方式:
        # 本地测试（需要设置环境变量）
        FEISHU_APP_ID=xxx FEISHU_APP_SECRET=yyy python main.py --mode local --url "PPT链接"
        
        # 本地文件处理（无需飞书凭证）
        python main.py --mode file --input ./template.pptx --output ./result.pptx --pages 15
        
        # 启动Webhook服务器
        python main.py --mode webhook --port 8080
    """
    main()
