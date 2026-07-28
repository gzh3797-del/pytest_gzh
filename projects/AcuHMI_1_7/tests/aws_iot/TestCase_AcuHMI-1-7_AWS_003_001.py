"""
FTS编号: FTS_WEB2_AWS_003_001
用例标题: 合法证书与密钥连接 AWS IoT Core
用例级别: LV0

预置条件:
  - 网关已登录
  - tests/protocols/aws_iot/certs/ 目录下存放合法的证书和密钥文件

测试步骤:
  1. Enable AWS IoT，填写合法 URL、Interval
  2. 上传合法 cert 文件和 key 文件
  3. 选择设备，点击 Save
  4. 点击 Test Connection，等待结果

预期结果:
  - 上传证书文件成功（无错误提示）
  - Test Connection 返回连接成功标识
"""

from pathlib import Path

import pytest

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent  # AcuHMI-1-7/


class TestCase_AcuHMI_1_7_AWS_003_001:

    @pytest.mark.aws_iot
    def test_valid_cert_connection(self, aws_page, aws_cfg):
        """上传合法证书 → Save → Test Connection 应成功"""
        aws = aws_cfg["aws_iot"]
        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.set_interval(aws.get("interval", "30 seconds"))
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.select_only_device()  # 随机选一台物理设备
        aws_page.save()

        result = aws_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"Test Connection 应返回成功，实际：{result!r}"
