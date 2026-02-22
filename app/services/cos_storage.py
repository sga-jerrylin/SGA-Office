"""
COS 云对象存储服务封装。
所有文件的上传/下载都通过此模块统一处理，彻底剥离对宿主机本地文件系统的依赖。
"""

import os
import uuid
import tempfile
from io import BytesIO
from datetime import datetime
from urllib.parse import quote

import requests
from qcloud_cos import CosConfig, CosS3Client

from app.core.config import get_settings


class CosStorageService:
    """腾讯云 COS 对象存储操作封装"""

    def __init__(self):
        settings = get_settings()
        self._config = CosConfig(
            Region=settings.cos_region,
            SecretId=settings.cos_secret_id,
            SecretKey=settings.cos_secret_key,
        )
        self._client = CosS3Client(self._config)
        self._bucket = settings.cos_bucket_name
        self._region = settings.cos_region

    @property
    def base_url(self) -> str:
        return f"https://{self._bucket}.cos.{self._region}.myqcloud.com"

    def upload_bytes(self, data: bytes, cos_key: str) -> str:
        """
        上传字节流到 COS。
        Args:
            data: 文件字节内容
            cos_key: COS 路径键名 (如 'documents/报告_20250221.docx')
        Returns:
            文件的公网访问 URL
        """
        stream = BytesIO(data)
        self._client.put_object(
            Bucket=self._bucket,
            Body=stream,
            Key=cos_key,
        )
        return f"{self.base_url}/{quote(cos_key)}"

    def upload_file(self, local_path: str, cos_key: str) -> str:
        """
        上传本地文件到 COS。
        Args:
            local_path: 本地文件绝对路径
            cos_key: COS 路径键名
        Returns:
            文件的公网访问 URL
        """
        self._client.upload_file(
            Bucket=self._bucket,
            LocalFilePath=local_path,
            Key=cos_key,
        )
        return f"{self.base_url}/{quote(cos_key)}"

    def download_to_bytes(self, url: str) -> bytes:
        """
        从 URL 下载文件到内存。支持 COS 内部链接和任意外部 URL。
        Args:
            url: 文件的可下载链接
        Returns:
            文件字节内容
        """
        response = requests.get(
            url,
            timeout=60,
            headers={
                "User-Agent": "SGA-Office/1.0 (Agent-First File Processor)"
            },
        )
        response.raise_for_status()
        return response.content

    def download_to_tempfile(self, url: str, suffix: str = "") -> str:
        """
        从 URL 下载文件到临时文件。
        Args:
            url: 文件的可下载链接
            suffix: 临时文件后缀 (如 '.docx')
        Returns:
            临时文件的本地路径 (调用方需负责清理)
        """
        data = self.download_to_bytes(url)
        fd, temp_path = tempfile.mkstemp(suffix=suffix)
        try:
            os.write(fd, data)
        finally:
            os.close(fd)
        return temp_path

    @staticmethod
    def generate_cos_key(prefix: str, filename: str, ext: str) -> str:
        """
        生成标准化的 COS 存储路径。
        格式: {prefix}/{清理后文件名}_{日期}_{短UUID}.{ext}
        """
        import re
        clean_name = re.sub(r'[\\/:*?"<>|\s]', '', filename)[:30]
        if not clean_name:
            clean_name = "unnamed"
        date_str = datetime.now().strftime("%Y%m%d")
        short_id = uuid.uuid4().hex[:8]
        return f"{prefix}/{clean_name}_{date_str}_{short_id}.{ext}"


# 模块级单例 (惰性初始化)
_cos_service: CosStorageService | None = None


def get_cos_service() -> CosStorageService:
    """获取 COS 服务单例"""
    global _cos_service
    if _cos_service is None:
        _cos_service = CosStorageService()
    return _cos_service
