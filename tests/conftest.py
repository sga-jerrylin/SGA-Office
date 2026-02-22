"""
共享 fixtures: TestClient、Mock COS 等。
在所有测试文件中自动可用。
"""

import os
import sys
from unittest.mock import MagicMock, patch

import pytest
from fastapi.testclient import TestClient

# ---- 在 import app 之前 mock COS，避免读取真实环境变量 ----

# 设置伪环境变量，防止 Settings 初始化报错
os.environ.setdefault("COS_SECRET_ID", "fake_id_for_test")
os.environ.setdefault("COS_SECRET_KEY", "fake_key_for_test")
os.environ.setdefault("COS_REGION", "ap-test")
os.environ.setdefault("COS_BUCKET_NAME", "test-bucket-123")


def _make_mock_cos():
    """创建一个完整的 COS mock 对象"""
    mock = MagicMock()
    mock.base_url = "https://test-bucket-123.cos.ap-test.myqcloud.com"
    mock.upload_bytes.return_value = "https://test-bucket-123.cos.ap-test.myqcloud.com/test/file.docx"
    mock.upload_file.return_value = "https://test-bucket-123.cos.ap-test.myqcloud.com/test/file.docx"
    mock.download_to_bytes.return_value = b"fake file bytes"
    mock.generate_cos_key.return_value = "documents/test_20250222_abc12345.docx"
    return mock


@pytest.fixture(autouse=True)
def mock_cos_service():
    """全局自动 mock COS 服务，所有测试都不会真正调用云存储"""
    mock = _make_mock_cos()
    with patch("app.services.cos_storage.get_cos_service", return_value=mock) as _:
        # 同时 patch 各个 route 模块中直接 import 的 get_cos_service
        with patch("app.api.endpoints.doc_routes.get_cos_service", return_value=mock):
            with patch("app.api.endpoints.excel_routes.get_cos_service", return_value=mock):
                yield mock


@pytest.fixture
def client(mock_cos_service):
    """FastAPI TestClient"""
    from app.main import app
    return TestClient(app)


@pytest.fixture
def cos_mock(mock_cos_service):
    """直接获取 mock COS 对象（别名）"""
    return mock_cos_service

