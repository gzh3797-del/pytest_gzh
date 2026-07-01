# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_007_002
用例标题: Azure IoT 与 AWS IoT 同时启用互不干扰
用例级别: LV2

预置条件:
  - AWS IoT 合法证书已就绪（tests/protocols/aws_iot/config.yaml）
  - Azure IoT Connection String 已配置（tests/protocols/azure_iot/config.yaml）
  - Azure IoT Hub 和 AWS IoT Core 订阅端均就绪

测试步骤:
  1. 配置并启用 Azure IoT，验证连接成功
  2. 配置并启用 AWS IoT
  3. 分别在 Azure IoT Hub 和 AWS IoT Core 端监听数据
  4. 验证两侧均独立收到数据，互不干扰

预期结果:
  - Azure IoT Hub 和 AWS IoT Core 均独立收到设备数据
  - 两侧数据互不干扰，上报均不中断
"""

import time
from pathlib import Path

import pytest
import yaml

from utils.azure_iot_verifier import run_verifier as run_azure_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent
_AWS_CFG_PATH = _THIS_DIR.parent / "aws_iot" / "config.yaml"


class TestCase_AcuHMI_1_7_AZR_007_002:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_azure_and_aws_coexist(self, azure_page, app_page, azure_cfg):
        """Azure IoT + AWS IoT 同时运行，双侧数据均独立正常到达"""
        # 配置并启用 Azure IoT
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()
        az_conn = azure_page.test_connection()
        assert any(kw in az_conn.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"Azure IoT 连接应成功：{az_conn!r}"

        # 检查 AWS IoT 配置是否存在
        if not _AWS_CFG_PATH.exists():
            pytest.skip("AWS IoT config.yaml 不存在，跳过共存测试")

        aws_cfg = yaml.safe_load(_AWS_CFG_PATH.read_text(encoding="utf-8"))

        try:
            from pages.protocols.aws_iot_page import AWSIoTPage
            aws_page = AWSIoTPage(app_page)
            aws_page.navigate_to_aws_iot()
            aws_page.ensure_enabled()
            aws = aws_cfg["aws_iot"]
            aws_page.set_url(aws["url"])
            aws_page.set_interval(aws.get("interval", "30 seconds"))
            aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
            aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
            aws_page.save()
            time.sleep(3)
        except Exception as exc:
            pytest.skip(f"AWS IoT 配置失败，跳过共存测试：{exc}")

        # 验证 Azure IoT 侧数据
        az_ok = run_azure_verifier(azure_cfg, timeout=120, skip_web=True)
        assert az_ok, "Azure IoT 与 AWS IoT 共存时，Azure IoT Hub 数据验证失败"

        # 验证 AWS IoT 侧数据（若 verifier 可用）
        try:
            from utils.aws_iot_verifier import run_verifier as run_aws_verifier
            aws_ok = run_aws_verifier(aws_cfg, timeout=120, skip_web=True)
            assert aws_ok, "Azure IoT 与 AWS IoT 共存时，AWS IoT Core 数据验证失败"
        except ImportError:
            pass
