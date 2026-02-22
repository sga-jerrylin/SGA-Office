"""
腾讯云 COS 对象存储客户端 (单例)
负责所有与云端存储的 I/O 操作，隔离底层 SDK 细节。
"""

from io import BytesIO
from urllib.parse import quote
from qcloud_cos import CosConfig, CosS3Client
from app.core.config import get_settings


def _create_cos_client() -> CosS3Client:
    """构建 COS 客户端实例"""
    settings = get_settings()
    config = CosConfig(
        Region=settings.cos_region,
        SecretId=settings.cos_secret_id,
        SecretKey=settings.cos_secret_key,
    )
    return CosS3Client(config)


# ---------- 模块级单例 ----------
cos_client: CosS3Client = _create_cos_client()


def upload_bytes_to_cos(
    data: bytes | BytesIO,
    cos_key: str,
) -> str:
    """
    将二进制数据上传至 COS，返回公网可达下载 URL。

    Args:
        data:    bytes 或 BytesIO 对象
        cos_key: COS 中的完整路径，如 "excel_documents/report_20250221.xlsx"

    Returns:
        公网下载 URL，如 https://bucket.cos.region.myqcloud.com/excel_documents/report.xlsx
    """
    settings = get_settings()
    body = data if isinstance(data, bytes) else data.getvalue()

    cos_client.put_object(
        Bucket=settings.cos_bucket_name,
        Body=body,
        Key=cos_key,
        ContentType=_guess_content_type(cos_key),
    )

    return f"{settings.cos_base_url}/{quote(cos_key, safe='/')}"


def upload_file_to_cos(local_path: str, cos_key: str) -> str:
    """
    将本地文件上传至 COS，返回公网可达下载 URL。

    Args:
        local_path: 本地文件的绝对路径
        cos_key:    COS 中的完整路径

    Returns:
        公网下载 URL
    """
    settings = get_settings()
    cos_client.upload_file(
        Bucket=settings.cos_bucket_name,
        LocalFilePath=local_path,
        Key=cos_key,
    )
    return f"{settings.cos_base_url}/{quote(cos_key, safe='/')}"


def download_bytes_from_cos(cos_key: str) -> bytes:
    """
    从 COS 下载文件到内存，返回二进制数据。
    """
    settings = get_settings()
    response = cos_client.get_object(
        Bucket=settings.cos_bucket_name,
        Key=cos_key,
    )
    return response["Body"].get_raw_stream().read()


def _guess_content_type(key: str) -> str:
    """根据文件后缀推断 Content-Type"""
    ext = key.rsplit(".", 1)[-1].lower() if "." in key else ""
    mapping = {
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pdf": "application/pdf",
        "png": "image/png",
        "svg": "image/svg+xml",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
    }
    return mapping.get(ext, "application/octet-stream")
