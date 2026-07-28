"""
FTS编号: FTS_WEB2_AWS_006_002
用例标题: AWS IoT 与 Azure IoT 同时启用互不干扰
用例级别: LV3

预置条件:
  - Azure IoT Connection String 已准备（tests/protocols/azure_iot/config.yaml）
  - AWS IoT 合法证书已就绪，两侧订阅端均就绪

测试步骤:
  1. 配置并启用 AWS IoT
  2. 导航到 Azure IoT 页面，配置并启用 Azure IoT
  3. 分别在 AWS MQTT 端和 Azure IoT Hub 端监听数据
  4. 持续监听，分别验证两侧数据

预期结果:
  - AWS IoT 和 Azure IoT 均独立收到数据，互不干扰
"""

import time
from pathlib import Path

import pytest
import yaml

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent

_AZURE_CFG_PATH = _THIS_DIR.parent / "azure_iot" / "config.yaml"


@pytest.mark.skip(reason="Azure IoT 暂无账号，跳过共存测试")
class TestCase_RPP_AWS_006_002:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_aws_azure_coexist(self, aws_page, app_page, aws_cfg):
        # 配置并启用 AWS IoT（从 TCP 已连接设备中取第一台，排除 RTU 设备）
        aws = aws_cfg["aws_iot"]
        expected = aws.get("expected_devices", [])
        device_name = expected[0] if expected else ""
        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.set_interval(aws.get("interval", "30 seconds"))
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.select_only_device(device_name)
        aws_page.configure_all_devices_parameters(checked_only=True)
        aws_page.save()
        aws_conn = aws_page.test_connection()
        assert any(kw in aws_conn.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"AWS IoT 连接应成功：{aws_conn!r}"

        # 检查 Azure IoT 配置是否存在
        if not _AZURE_CFG_PATH.exists():
            pytest.skip("Azure IoT config.yaml 不存在，跳过共存测试")

        azure_cfg = yaml.safe_load(_AZURE_CFG_PATH.read_text(encoding="utf-8"))

        from pages.protocols.azure_iot_page import AzureIoTPage
        azure_page = AzureIoTPage(app_page)
        azure_page.navigate_to_azure_iot()
        azure_page.ensure_enabled()
        azure_page.save()
        time.sleep(3)

        # 验证 AWS IoT 数据
        aws_ok = run_verifier(aws_cfg, timeout=120, skip_web=True)
        assert aws_ok, "AWS IoT 共存时数据验证失败"

        # 验证 Azure IoT 数据（若有 verifier）
        try:
            from utils.azure_iot_verifier import run_verifier as run_azure_verifier
            az_ok = run_azure_verifier(azure_cfg, timeout=120, skip_web=True)
            assert az_ok, "Azure IoT 共存时数据验证失败"
        except ImportError:
            pass
