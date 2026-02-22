"""
SGA-Office 全局配置管理
使用 Pydantic BaseSettings 从环境变量统一加载所有敏感凭证与运行参数。
"""

from pydantic_settings import BaseSettings
from pydantic import Field
from functools import lru_cache


class Settings(BaseSettings):
    """
    系统全局配置。所有字段均从环境变量或 .env 文件中读取。
    变量名映射关系：字段名自动对应大写的环境变量名。
    """

    # ========== 腾讯云 COS 对象存储 ==========
    cos_secret_id: str = Field(default="", description="腾讯云 COS SecretId")
    cos_secret_key: str = Field(default="", description="腾讯云 COS SecretKey")
    cos_region: str = Field(default="ap-guangzhou", description="COS 区域")
    cos_bucket_name: str = Field(default="", description="COS 存储桶名称")

    # ========== 服务运行参数 ==========
    api_host: str = Field(default="0.0.0.0", description="API 监听地址")
    api_port: int = Field(default=8000, description="API 监听端口")
    debug: bool = Field(default=False, description="是否开启调试模式")
    app_name: str = Field(default="SGA-Office", description="应用名称")
    api_version: str = Field(default="v1", description="API 版本号")

    # ========== 文件处理参数 ==========
    max_upload_size_mb: int = Field(default=50, description="最大上传文件体积(MB)")
    temp_dir: str = Field(default="/tmp/sga-office", description="临时文件目录")

    @property
    def cos_base_url(self) -> str:
        """COS 文件的基础访问 URL"""
        return f"https://{self.cos_bucket_name}.cos.{self.cos_region}.myqcloud.com"

    model_config = {
        "env_file": ".env",
        "env_file_encoding": "utf-8",
        "case_sensitive": False,  # 环境变量不区分大小写匹配
    }


@lru_cache()
def get_settings() -> Settings:
    """
    使用 lru_cache 确保全局只加载一次配置实例（单例模式）。
    在 FastAPI 的 Depends 依赖注入中使用。
    """
    return Settings()
