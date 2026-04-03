"""
feishu/file_manager.py
======================
飞书文件管理器 —— 对接飞书云文档/PPT的核心桥梁。

职责：
  1. 通过飞书PPT分享链接解析文件token、获取文件基础信息
  2. 下载飞书PPT文件到本地（供ppt_analyzer.py解析）
  3. 上传本地PPT到飞书云文档，生成可编辑链接

依赖：requests

飞书应用需要开启的权限清单：
  - drive:drive:readonly     读取云盘文件
  - drive:drive:write        写入云盘文件
  - drive:file:download      下载文件
  - drive:file:upload        上传文件
  - drive:file:share         创建分享链接
  - im:message:send          发送消息（如需要投递到对话）

环境依赖：
  pip install requests

运行步骤：
  1. 设置环境变量：FEISHU_APP_ID 和 FEISHU_APP_SECRET
  2. 在 __main__ 中替换 TEST_FEISHU_PPT_URL 为你的飞书PPT链接
  3. 运行：python -m feishu.file_manager
"""

from __future__ import annotations

import json
import logging
import re
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse, parse_qs

import requests

from .auth import FeishuAuth, FeishuAPIError, get_auth, FeishuAuthError

logger = logging.getLogger(__name__)

# 飞书API基础URL
FEISHU_API_BASE = "https://open.feishu.cn/open-apis"

# 文件大小限制
MAX_DOWNLOAD_SIZE = 50 * 1024 * 1024   # 50MB
MAX_UPLOAD_SIZE = 200 * 1024 * 1024    # 200MB


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class PPTFileInfo:
    """
    飞书PPT文件基础信息。
    
    Attributes
    ----------
    file_token : str
        飞书文件唯一标识
    file_name : str
        文件名（含扩展名）
    file_type : str
        文件类型（如pptx）
    owner : str
        创建人/拥有者名称
    create_time : str
        创建时间（ISO格式）
    version : str
        版本号
    size : int
        文件大小（字节）
    url : str
        原始分享链接
    """
    file_token: str
    file_name: str
    file_type: str
    owner: str = ""
    create_time: str = ""
    version: str = ""
    size: int = 0
    url: str = ""


@dataclass
class DownloadResult:
    """文件下载结果。"""
    file_path: str           # 本地临时保存路径
    file_name: str           # 原始文件名
    file_size: int           # 字节数
    mime_type: str = ""


@dataclass
class UploadResult:
    """文件上传结果。"""
    file_token: str          # 飞书文件 token
    file_name: str
    file_type: str           # 文件 MIME 类型
    feishu_url: str = ""     # 飞书文档URL


@dataclass
class ShareResult:
    """分享链接创建结果。"""
    url: str                 # 分享链接
    file_token: str


# ---------------------------------------------------------------------------
# 异常定义
# ---------------------------------------------------------------------------

class FileManagerError(Exception):
    """文件管理器通用异常。"""
    
    def __init__(self, message: str, solution: str = ""):
        self.message = message
        self.solution = solution
        super().__init__(message)


class InvalidURLError(FileManagerError):
    """URL解析失败异常。"""
    pass


class PermissionDeniedError(FileManagerError):
    """权限不足异常。"""
    pass


class FeishuFileNotFoundError(FileManagerError):
    """飞书文件不存在异常。"""
    pass


# ---------------------------------------------------------------------------
# 辅助函数
# ---------------------------------------------------------------------------

def _parse_feishu_url(url: str) -> tuple[str, str] | None:
    """
    从飞书文档 URL 中提取 file_token 和 file_type。
    
    支持格式：
      - https://xxx.feishu.cn/docx/XXXXX
      - https://xxx.feishu.cn/sheets/XXXXX
      - https://xxx.feishu.cn/slides/XXXXX    (PPT)
      - https://xxx.feishu.cn/file/XXXXX
      - https://xxx.feishu.cn/drive/folder/XXXXX
      - https://open.feishu.cn/open-apis/drive/v1/files/XXXXX
    
    Parameters
    ----------
    url : str
        飞书文档分享链接
        
    Returns
    -------
    tuple[str, str] | None
        (file_token, file_type) 或 None（无法识别）
    """
    patterns = [
        # 标准飞书文档 URL（捕获 type 和 token）
        r"feishu\.cn/(?P<type>docx|sheets|bitable|mindnote|file|slides)/(?P<token>[a-zA-Z0-9_-]+)",
        # drive 文件夹 URL
        r"feishu\.cn/drive/folder/(?P<token>[a-zA-Z0-9_-]+)",
        # API 直链
        r"drive/v1/files/(?P<token>[a-zA-Z0-9_-]+)",
    ]
    for pat in patterns:
        m = re.search(pat, url)
        if m:
            token = m.group("token")
            file_type = m.groupdict().get("type", "file")
            return token, file_type
    return None


def _format_file_size(size_bytes: int) -> str:
    """将字节数转换为人类可读的格式。"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / 1024 / 1024:.1f} MB"
    else:
        return f"{size_bytes / 1024 / 1024 / 1024:.1f} GB"


# ---------------------------------------------------------------------------
# 核心类
# ---------------------------------------------------------------------------

class FeishuFileManager:
    """
    飞书文件管理器。
    
    提供PPT文件的解析、下载、上传全流程管理能力。
    
    Parameters
    ----------
    auth : FeishuAuth | None
        鉴权实例，为 None 时自动获取全局单例。
    download_dir : str | Path
        文件下载的本地临时目录。
        
    Examples
    --------
    >>> from feishu.file_manager import FeishuFileManager
    >>> fm = FeishuFileManager()
    >>> 
    >>> # 解析PPT链接获取信息
    >>> info = fm.get_ppt_info("https://xxx.feishu.cn/slides/xxx")
    >>> print(f"文件名: {info.file_name}, 大小: {info.size}")
    >>> 
    >>> # 下载PPT到本地
    >>> result = fm.download_ppt("https://xxx.feishu.cn/slides/xxx")
    >>> print(f"已下载到: {result.file_path}")
    >>> 
    >>> # 上传本地PPT到飞书
    >>> upload = fm.upload_ppt("/path/to/output.pptx")
    >>> print(f"飞书链接: {upload.feishu_url}")
    """

    def __init__(
        self,
        auth: FeishuAuth | None = None,
        download_dir: str | Path = "./downloads",
    ):
        self.auth = auth or get_auth()
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)

    # ===================================================================
    # 核心方法1：通过分享链接获取PPT文件信息
    # ===================================================================

    def get_ppt_info(self, feishu_url: str) -> PPTFileInfo:
        """
        通过飞书PPT分享链接，解析出文件token、获取文件基础信息。
        
        流程：
          1. 解析URL提取file_token
          2. 调用飞书API获取文件元数据
          3. 返回结构化的PPT文件信息
        
        Parameters
        ----------
        feishu_url : str
            飞书PPT分享链接，如：https://xxx.feishu.cn/slides/sldcnXXX
            
        Returns
        -------
        PPTFileInfo
            PPT文件的结构化信息
            
        Raises
        ------
        InvalidURLError
            URL格式无效，无法解析出file_token
        PermissionDeniedError
            应用无权限访问该文件（需申请 drive:drive:readonly 权限）
        FeishuFileNotFoundError
            文件不存在或已被删除
        FeishuAPIError
            飞书API返回其他业务错误
            
        Examples
        --------
        >>> info = fm.get_ppt_info("https://abc.feishu.cn/slides/sldcnABC123")
        >>> print(info.file_name)  # "产品汇报.pptx"
        >>> print(info.owner)      # "张三"
        """
        # Step 1: 解析URL获取file_token
        parsed = _parse_feishu_url(feishu_url)
        if not parsed:
            raise InvalidURLError(
                message=f"无法识别的飞书URL格式: {feishu_url}",
                solution="请确认URL为飞书PPT分享链接，格式如：https://xxx.feishu.cn/slides/xxx"
            )
        
        file_token, file_type = parsed
        logger.info(f"解析URL成功: token={file_token}, type={file_type}")
        
        # Step 2: 调用飞书API获取文件元数据
        # API文档: https://open.feishu.cn/document/server-docs/docs/drive-v1/file/get
        try:
            url = f"{FEISHU_API_BASE}/drive/v1/files/{file_token}"
            data = self.auth.get(url)
        except FeishuAPIError as e:
            # 根据错误码分类处理
            if e.code == 403:
                raise PermissionDeniedError(
                    message=f"无权限访问文件: {e.msg}",
                    solution="请确认：1) 应用已开通 'drive:drive:readonly' 权限；2) 文件已分享给应用所在企业"
                )
            elif e.code == 404 or e.code == 125404:
                raise FeishuFileNotFoundError(
                    message=f"文件不存在: {file_token}",
                    solution="请确认文件未被删除，且分享链接有效"
                )
            else:
                raise
        
        # Step 3: 解析响应数据
        # 飞书API返回的字段可能因文件类型而异，这里做防御性处理
        file_info = PPTFileInfo(
            file_token=file_token,
            file_name=data.get("name", f"unknown_{file_token}.pptx"),
            file_type=data.get("type", file_type),
            owner=data.get("owner", {}).get("name", ""),
            create_time=data.get("create_time", ""),
            version=str(data.get("revision", "")),
            size=data.get("size", 0),
            url=feishu_url,
        )
        
        logger.info(f"获取文件信息成功: {file_info.file_name}, 大小={_format_file_size(file_info.size)}")
        return file_info

    # ===================================================================
    # 核心方法2：下载飞书PPT文件到本地
    # ===================================================================

    def download_ppt(
        self, 
        source: str,
        local_name: str = "",
    ) -> DownloadResult:
        """
        下载飞书PPT文件到本地，供后续ppt_analyzer.py解析使用。
        
        支持多种来源：
          - 飞书PPT分享链接（自动解析token并下载）
          - file_token（直接传入token字符串）
          - 本地文件路径（直接返回，用于统一接口）
        
        Parameters
        ----------
        source : str
            下载来源：
            - 飞书URL: "https://xxx.feishu.cn/slides/xxx"
            - file_token: "sldcnABC123" 或 "filecnABC123"
            - 本地路径: "/path/to/local.pptx"
        local_name : str, optional
            指定保存到本地的文件名，不传则使用原始文件名
            
        Returns
        -------
        DownloadResult
            下载结果，包含本地路径、文件名、大小等信息
            
        Raises
        ------
        InvalidURLError
            无法识别的来源格式
        PermissionDeniedError
            应用无权限下载该文件（需申请 drive:file:download 权限）
        FileNotFoundError
            文件不存在
        FileManagerError
            文件过大（超过50MB限制）或其他下载错误
            
        Examples
        --------
        >>> # 从URL下载
        >>> result = fm.download_ppt("https://xxx.feishu.cn/slides/xxx")
        >>> 
        >>> # 从token下载
        >>> result = fm.download_ppt("sldcnABC123")
        >>> 
        >>> # 指定保存名称
        >>> result = fm.download_ppt("https://xxx.feishu.cn/slides/xxx", "template.pptx")
        """
        # 判断来源类型
        file_token = ""
        
        # 1. 本地文件路径 - 直接返回
        path = Path(source)
        if path.exists():
            if path.suffix.lower() not in [".pptx", ".ppt"]:
                raise FileManagerError(
                    message=f"不支持的文件格式: {path.suffix}",
                    solution="请提供 .pptx 或 .ppt 格式的PPT文件"
                )
            logger.info(f"本地文件已存在，直接使用: {path}")
            return DownloadResult(
                file_path=str(path),
                file_name=path.name,
                file_size=path.stat().st_size,
                mime_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        
        # 2. 飞书URL - 解析token
        parsed = _parse_feishu_url(source)
        if parsed:
            file_token, _ = parsed
            logger.info(f"从URL解析到file_token: {file_token}")
        # 3. 纯token字符串 - 直接识别
        elif re.fullmatch(r"[a-zA-Z0-9_-]+", source):
            file_token = source
            logger.info(f"识别为file_token: {file_token}")
        else:
            raise InvalidURLError(
                message=f"无法识别的下载来源: {source}",
                solution="请提供飞书PPT分享链接、file_token或本地文件路径"
            )
        
        # 调用飞书API下载文件
        # API文档: https://open.feishu.cn/document/server-docs/docs/drive-v1/file/download
        try:
            url = f"{FEISHU_API_BASE}/drive/v1/files/{file_token}/download"
            headers = self.auth.get_headers(content_type="")
            headers["Accept"] = "application/octet-stream"
            
            logger.info(f"开始下载文件: {file_token}")
            resp = requests.get(url, headers=headers, timeout=120, stream=True)
            
            # 处理HTTP错误
            if resp.status_code == 403:
                raise PermissionDeniedError(
                    message="无权限下载文件",
                    solution="请确认应用已开通 'drive:file:download' 权限，且文件已分享给应用"
                )
            elif resp.status_code == 404:
                raise FeishuFileNotFoundError(
                    message=f"文件不存在: {file_token}",
                    solution="请确认文件未被删除"
                )
            elif resp.status_code != 200:
                raise FeishuAPIError(
                    code=resp.status_code,
                    msg=f"下载失败: HTTP {resp.status_code}",
                )
            
            # 读取响应内容
            content = resp.content
            file_size = len(content)
            
            # 检查文件大小
            if file_size > MAX_DOWNLOAD_SIZE:
                raise FileManagerError(
                    message=f"文件过大 ({_format_file_size(file_size)})，超过 {_format_file_size(MAX_DOWNLOAD_SIZE)} 限制",
                    solution="请使用更小的PPT文件，或联系管理员调整大小限制"
                )
            
            # 提取文件名
            if not local_name:
                # 尝试从Content-Disposition获取
                cd = resp.headers.get("Content-Disposition", "")
                m = re.search(r'filename\*?=(?:UTF-8\'\')?["\']?([^"\';]+)["\']?', cd, re.IGNORECASE)
                if m:
                    local_name = m.group(1)
                else:
                    local_name = f"{file_token}.pptx"
            
            # 确保扩展名正确
            if not local_name.lower().endswith('.pptx'):
                local_name += '.pptx'
            
            # 保存到本地
            save_path = self.download_dir / local_name
            save_path.write_bytes(content)
            
            logger.info(f"文件下载完成: {save_path} ({_format_file_size(file_size)})")
            
            return DownloadResult(
                file_path=str(save_path),
                file_name=local_name,
                file_size=file_size,
                mime_type=resp.headers.get("Content-Type", "application/octet-stream"),
            )
            
        except (PermissionDeniedError, FileNotFoundError, FileManagerError):
            raise
        except Exception as e:
            raise FileManagerError(
                message=f"下载文件时发生错误: {str(e)}",
                solution="请检查网络连接，确认飞书API服务正常"
            )

    # ===================================================================
    # 核心方法3：上传本地PPT到飞书云文档
    # ===================================================================

    def upload_ppt(
        self,
        local_path: str | Path,
        folder_token: str = "",
        file_name: str = "",
        create_share: bool = True,
    ) -> UploadResult:
        """
        把本地生成好的复刻PPT文件，上传到飞书云文档，生成可编辑的飞书PPT链接。
        
        流程：
          1. 读取本地PPT文件
          2. 上传到飞书云盘（小文件单次上传，大文件分片上传）
          3. 可选：创建分享链接
          4. 返回包含飞书URL的上传结果
        
        Parameters
        ----------
        local_path : str | Path
            本地PPT文件路径
        folder_token : str, optional
            目标文件夹token，为空则上传到根目录
        file_name : str, optional
            指定上传后的文件名，不传则使用原始文件名
        create_share : bool, default True
            是否自动创建分享链接
            
        Returns
        -------
        UploadResult
            上传结果，包含file_token和飞书文档URL
            
        Raises
        ------
        FeishuFileNotFoundError
            本地文件不存在
        FileManagerError
            文件过大（超过200MB限制）或格式不支持
        PermissionDeniedError
            应用无权限上传文件（需申请 drive:file:upload 权限）
            
        Examples
        --------
        >>> # 基础上传
        >>> result = fm.upload_ppt("/path/to/output.pptx")
        >>> print(result.feishu_url)  # https://xxx.feishu.cn/slides/xxx
        >>> 
        >>> # 上传到指定文件夹
        >>> result = fm.upload_ppt("output.pptx", folder_token="fldcnXXX")
        >>> 
        >>> # 不创建分享链接
        >>> result = fm.upload_ppt("output.pptx", create_share=False)
        """
        path = Path(local_path)
        
        # 前置检查
        if not path.exists():
            raise FeishuFileNotFoundError(
                message=f"本地文件不存在: {path}",
                solution="请确认文件路径正确，且文件已成功生成"
            )
        
        if path.suffix.lower() != ".pptx":
            raise FileManagerError(
                message=f"不支持的文件格式: {path.suffix}",
                solution="请上传 .pptx 格式的文件"
            )
        
        file_size = path.stat().st_size
        if file_size > MAX_UPLOAD_SIZE:
            raise FileManagerError(
                message=f"文件过大 ({_format_file_size(file_size)})，超过 {_format_file_size(MAX_UPLOAD_SIZE)} 限制",
                solution="请优化PPT文件大小，或联系管理员调整限制"
            )
        
        upload_name = file_name or path.name
        logger.info(f"开始上传文件: {upload_name} ({_format_file_size(file_size)})")
        
        try:
            # 根据文件大小选择上传方式
            if file_size <= 20 * 1024 * 1024:
                result = self._upload_small_file(path, upload_name, folder_token)
            else:
                result = self._upload_large_file(path, upload_name, folder_token)
            
            # 创建分享链接
            if create_share and result.file_token:
                try:
                    share = self._create_share_link(result.file_token)
                    result.feishu_url = share.url
                    logger.info(f"分享链接已创建: {share.url}")
                except Exception as e:
                    logger.warning(f"创建分享链接失败: {e}")
                    # 构造基础URL（可能无法直接访问）
                    result.feishu_url = f"https://open.feishu.cn/open-apis/drive/v1/files/{result.file_token}"
            
            return result
            
        except FeishuAPIError as e:
            if e.code == 403:
                raise PermissionDeniedError(
                    message=f"无权限上传文件: {e.msg}",
                    solution="请确认应用已开通 'drive:file:upload' 权限"
                )
            raise
        except (PermissionDeniedError, FileManagerError):
            raise
        except Exception as e:
            raise FileManagerError(
                message=f"上传文件时发生错误: {str(e)}",
                solution="请检查网络连接，确认飞书API服务正常"
            )

    def _upload_small_file(
        self,
        path: Path,
        file_name: str,
        folder_token: str,
    ) -> UploadResult:
        """
        小文件单次上传（≤ 20MB）。
        
        API文档: https://open.feishu.cn/document/server-docs/docs/drive-v1/file/upload_all
        """
        url = f"{FEISHU_API_BASE}/drive/v1/files/upload_all"
        headers = self.auth.get_headers(content_type="")
        
        # 构造multipart表单
        form_data = {
            "file_size": str(path.stat().st_size),
            "file_name": file_name,
            "file_type": "pptx",
        }
        if folder_token:
            form_data["folder_token"] = folder_token
        
        with open(path, "rb") as f:
            files = {
                "file": (file_name, f, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            }
            resp = requests.post(
                url,
                headers=headers,
                data=form_data,
                files=files,
                timeout=120,
            )
        
        body = resp.json()
        if body.get("code") != 0:
            raise FeishuAPIError(
                code=body.get("code", -1),
                msg=f"上传失败: {body.get('msg', '未知错误')}",
            )
        
        file_token = body["data"]["file_token"]
        logger.info(f"小文件上传成功: token={file_token}")
        
        return UploadResult(
            file_token=file_token,
            file_name=file_name,
            file_type="pptx",
        )

    def _upload_large_file(
        self,
        path: Path,
        file_name: str,
        folder_token: str,
    ) -> UploadResult:
        """
        大文件分片上传（> 20MB）。
        
        流程：upload_prepare → upload_part（多次）→ upload_finish
        
        API文档: 
          - https://open.feishu.cn/document/server-docs/docs/drive-v1/file/upload_prepare
          - https://open.feishu.cn/document/server-docs/docs/drive-v1/file/upload_part
          - https://open.feishu.cn/document/server-docs/docs/drive-v1/file/upload_finish
        """
        file_size = path.stat().st_size
        CHUNK_SIZE = 4 * 1024 * 1024  # 4MB分片
        
        # Step 1: 准备上传
        prepare_url = f"{FEISHU_API_BASE}/drive/v1/files/upload_prepare"
        prepare_data = self.auth.post(prepare_url, json={
            "file_name": file_name,
            "file_type": "pptx",
            "size": file_size,
            **({"folder_token": folder_token} if folder_token else {}),
        })
        
        upload_id = prepare_data.get("upload_id", "")
        block_size = prepare_data.get("block_size", CHUNK_SIZE)
        
        logger.info(f"大文件上传准备完成: upload_id={upload_id}, 分片大小={_format_file_size(block_size)}")
        
        # Step 2: 分片上传
        part_url = f"{FEISHU_API_BASE}/drive/v1/files/upload_part"
        headers_base = self.auth.get_headers(content_type="")
        
        with open(path, "rb") as f:
            seq = 0
            while True:
                chunk = f.read(block_size)
                if not chunk:
                    break
                
                resp = requests.post(
                    part_url,
                    headers=headers_base,
                    data={
                        "upload_id": upload_id,
                        "seq": seq,
                        "size": len(chunk),
                    },
                    files={"file": (f"chunk_{seq}", chunk, "application/octet-stream")},
                    timeout=120,
                )
                
                body = resp.json()
                if body.get("code") != 0:
                    raise FeishuAPIError(
                        code=body.get("code", -1),
                        msg=f"分片上传失败 (seq={seq}): {body.get('msg')}",
                    )
                
                seq += 1
                logger.debug(f"分片 {seq} 上传完成")
        
        # Step 3: 完成上传
        finish_url = f"{FEISHU_API_BASE}/drive/v1/files/upload_finish"
        finish_data = self.auth.post(finish_url, json={
            "upload_id": upload_id,
        })
        
        file_token = finish_data.get("file_token", "")
        logger.info(f"大文件上传成功: token={file_token}, 共{seq}个分片")
        
        return UploadResult(
            file_token=file_token,
            file_name=file_name,
            file_type="pptx",
        )

    def _create_share_link(self, file_token: str) -> ShareResult:
        """
        为文件创建公开分享链接。
        
        API文档: https://open.feishu.cn/document/server-docs/docs/drive-v1/permission/public-create
        """
        url = f"{FEISHU_API_BASE}/drive/v1/permissions/{file_token}/public"
        
        # 先尝试创建分享链接
        try:
            # 默认允许任何人访问，可通过环境变量配置为仅企业内访问
            share_scope = os.getenv("FEISHU_SHARE_SCOPE", "anyone")
            data = self.auth.post(url, json={
                "external_access": True,    # 允许外部访问
                "security_entity": "none",  # 无密码
                "link_share_entity": share_scope,  # anyone 或 tenant
            })
        except FeishuAPIError as e:
            # 如果已存在分享链接，尝试获取现有链接
            if e.code == 400 and "already exists" in e.msg.lower():
                data = self.auth.get(url)
            else:
                raise
        
        share_url = data.get("link", "")
        if not share_url:
            # 尝试从其他字段获取
            share_url = data.get("url", "")
        
        if not share_url:
            raise FeishuAPIError(
                code=-1,
                msg="创建分享链接失败: 未返回URL",
            )
        
        return ShareResult(url=share_url, file_token=file_token)


# ---------------------------------------------------------------------------
# 测试入口
# ---------------------------------------------------------------------------

# ==================== 配置区域：请修改以下配置 ====================

# 测试用的飞书PPT分享链接（请替换为你自己的链接）
TEST_FEISHU_PPT_URL = "https://xxx.feishu.cn/slides/sldcnXXXXXXXXXXXX"

# 测试用的本地PPT文件路径（用于测试上传功能）
TEST_LOCAL_PPT_PATH = "./test_output.pptx"

# 飞书应用凭证（优先从环境变量读取，也可在此硬编码用于测试）
TEST_APP_ID = os.getenv("FEISHU_APP_ID", "")
TEST_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")

# =================================================================


def _print_separator(title: str = ""):
    """打印分隔线。"""
    print("=" * 60)
    if title:
        print(f"  {title}")
        print("=" * 60)


def _print_result(success: bool, message: str, solution: str = ""):
    """打印测试结果。"""
    status = "✅ 成功" if success else "❌ 失败"
    print(f"  [{status}] {message}")
    if solution and not success:
        print(f"      💡 解决方案: {solution}")


def test_get_ppt_info(fm: FeishuFileManager) -> bool:
    """测试：获取PPT文件信息。"""
    print("\n【测试1】获取PPT文件信息")
    print(f"  测试链接: {TEST_FEISHU_PPT_URL}")
    
    # 检查是否已配置测试链接
    if "xxx" in TEST_FEISHU_PPT_URL or not TEST_FEISHU_PPT_URL:
        _print_result(False, "未配置测试链接", 
                     "请修改代码中的 TEST_FEISHU_PPT_URL 为你的飞书PPT分享链接")
        return False
    
    try:
        info = fm.get_ppt_info(TEST_FEISHU_PPT_URL)
        print(f"  📄 文件名: {info.file_name}")
        print(f"  👤 创建人: {info.owner}")
        print(f"  📅 创建时间: {info.create_time}")
        print(f"  🏷️  版本号: {info.version}")
        print(f"  📦 文件大小: {_format_file_size(info.size)}")
        print(f"  🔑 file_token: {info.file_token}")
        _print_result(True, "成功获取文件信息")
        return True
        
    except InvalidURLError as e:
        _print_result(False, e.message, e.solution)
    except PermissionDeniedError as e:
        _print_result(False, e.message, e.solution)
    except FeishuFileNotFoundError as e:
        _print_result(False, e.message, e.solution)
    except FeishuAPIError as e:
        _print_result(False, f"飞书API错误: {e.msg}", 
                     "请检查应用权限配置和网络连接")
    except Exception as e:
        _print_result(False, f"未知错误: {str(e)}", 
                     "请查看详细错误日志")
    
    return False


def test_download_ppt(fm: FeishuFileManager) -> tuple[bool, str]:
    """测试：下载PPT文件。"""
    print("\n【测试2】下载PPT文件到本地")
    
    # 检查是否已配置测试链接
    if "xxx" in TEST_FEISHU_PPT_URL or not TEST_FEISHU_PPT_URL:
        _print_result(False, "跳过测试（未配置测试链接）")
        return False, ""
    
    try:
        result = fm.download_ppt(TEST_FEISHU_PPT_URL, local_name="test_download.pptx")
        print(f"  💾 保存路径: {result.file_path}")
        print(f"  📄 文件名: {result.file_name}")
        print(f"  📦 文件大小: {_format_file_size(result.file_size)}")
        _print_result(True, f"成功下载到: {result.file_path}")
        return True, result.file_path
        
    except PermissionDeniedError as e:
        _print_result(False, e.message, e.solution)
    except FeishuFileNotFoundError as e:
        _print_result(False, e.message, e.solution)
    except FileManagerError as e:
        _print_result(False, e.message, e.solution)
    except Exception as e:
        _print_result(False, f"下载失败: {str(e)}")
    
    return False, ""


def test_upload_ppt(fm: FeishuFileManager, local_path: str = "") -> bool:
    """测试：上传PPT文件到飞书。"""
    print("\n【测试3】上传本地PPT到飞书")
    
    # 确定测试文件路径
    test_path = local_path or TEST_LOCAL_PPT_PATH
    
    # 如果没有现成的测试文件，尝试使用刚下载的文件
    if not Path(test_path).exists() and local_path:
        test_path = local_path
    
    if not Path(test_path).exists():
        _print_result(False, f"测试文件不存在: {test_path}", 
                     "请先运行下载测试，或修改 TEST_LOCAL_PPT_PATH 为有效的本地PPT路径")
        return False
    
    try:
        result = fm.upload_ppt(test_path, file_name="test_upload.pptx")
        print(f"  🔑 file_token: {result.file_token}")
        print(f"  📄 文件名: {result.file_name}")
        print(f"  🔗 飞书链接: {result.feishu_url}")
        _print_result(True, f"成功上传，飞书链接: {result.feishu_url}")
        return True
        
    except PermissionDeniedError as e:
        _print_result(False, e.message, e.solution)
    except FileManagerError as e:
        _print_result(False, e.message, e.solution)
    except Exception as e:
        _print_result(False, f"上传失败: {str(e)}")
    
    return False


def main():
    """
    独立测试入口。
    
    运行方式:
      1. 设置环境变量:
         set FEISHU_APP_ID=cli_xxx
         set FEISHU_APP_SECRET=xxx
      
      2. 修改本文件中的 TEST_FEISHU_PPT_URL 为你的飞书PPT链接
      
      3. 运行测试:
         python -m feishu.file_manager
    """
    _print_separator("FeishuFileManager 功能测试")
    
    # 检查凭证配置
    print("\n【环境检查】")
    app_id = TEST_APP_ID or os.getenv("FEISHU_APP_ID", "")
    app_secret = TEST_APP_SECRET or os.getenv("FEISHU_APP_SECRET", "")
    
    if not app_id or not app_secret:
        _print_result(False, "缺少飞书应用凭证")
        print("\n  💡 请通过以下方式之一配置凭证:")
        print("     方式1 - 环境变量:")
        print("       set FEISHU_APP_ID=cli_xxxxxxxxxxxxxxxx")
        print("       set FEISHU_APP_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        print("     方式2 - 修改代码中的 TEST_APP_ID 和 TEST_APP_SECRET")
        print("\n  📖 获取凭证:")
        print("     1. 访问 https://open.feishu.cn/app")
        print("     2. 创建或选择你的自建应用")
        print("     3. 在「凭证信息」中获取 App ID 和 App Secret")
        return
    
    _print_result(True, f"凭证已配置 (App ID: {app_id[:10]}...)")
    
    # 初始化鉴权和文件管理器
    print("\n【初始化】")
    try:
        from .auth import init_auth
        init_auth(app_id=app_id, app_secret=app_secret)
        fm = FeishuFileManager()
        _print_result(True, "FeishuFileManager 初始化成功")
    except FeishuAuthError as e:
        _print_result(False, f"初始化失败: {e}")
        return
    except Exception as e:
        _print_result(False, f"初始化异常: {str(e)}")
        return
    
    # 运行测试
    results = []
    
    # 测试1: 获取PPT信息
    results.append(("获取PPT信息", test_get_ppt_info(fm)))
    
    # 测试2: 下载PPT
    success, downloaded_path = test_download_ppt(fm)
    results.append(("下载PPT", success))
    
    # 测试3: 上传PPT（使用刚下载的文件）
    if downloaded_path:
        results.append(("上传PPT", test_upload_ppt(fm, downloaded_path)))
    else:
        results.append(("上传PPT", test_upload_ppt(fm)))
    
    # 汇总结果
    _print_separator("测试结果汇总")
    print()
    for name, passed in results:
        status = "✅ 通过" if passed else "❌ 失败"
        print(f"  {status} - {name}")
    
    total = len(results)
    passed = sum(1 for _, p in results if p)
    print(f"\n  总计: {passed}/{total} 项测试通过")
    
    if passed == total:
        print("\n  🎉 所有测试通过！FeishuFileManager 工作正常。")
    else:
        print("\n  ⚠️  部分测试未通过，请根据上方提示排查问题。")
        print("\n  常见问题:")
        print("    1. 权限不足 → 检查飞书应用是否开通了所需权限")
        print("    2. 文件不存在 → 确认分享链接有效且文件未被删除")
        print("    3. 网络错误 → 检查网络连接和飞书API服务状态")


if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%H:%M:%S"
    )
    
    # 运行测试
    main()
