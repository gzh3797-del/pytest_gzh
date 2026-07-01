# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_002_003
用例标题: Secondary Connection String 配置及 Primary 失效切换
用例级别: LV1

预置条件:
  1. 设备已正常上电，浏览器访问网关 Web UI 并以 admin 登录
  2. config.yaml 中已配置合法的 primary_conn_str 和 secondary_conn_str

测试步骤:
  1. 配置合法 Primary + Secondary Connection String，保存，验证连接成功
  2. 将 Primary 改为无效值，保存，等待 Primary 超时失效
  3. 验证系统自动切换到 Secondary，Azure IoT Hub 持续收到设备数据

预期结果:
  1. Primary + Secondary 均合法时连接成功
  2. Primary 失效后系统自动切换到 Secondary Connection String
  3. Azure IoT Hub 在切换后仍能持续收到设备数据，上报不中断
"""

import time
from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR = Path(__file__).resolve().parent

_INVALID_PRIMARY = (
    "HostName=test.azure-devices.net;"
    "DeviceId=nonexistent-device-failover-test;"
    "SharedAccessKey=AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="
)


class TestCase_AcuHMI_1_7_AZR_002_003:

    @pytest.mark.azure_iot
    def test_primary_and_secondary_both_valid(self, azure_page, azure_cfg):
        """Primary + Secondary 均配置时 Test Connection 应成功"""
        azure = azure_cfg["azure_iot"]
        secondary = azure.get("secondary_conn_str", "")
        if not secondary:
            pytest.skip("config.yaml 未配置 secondary_conn_str，跳过本用例")

        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_secondary_conn_str(secondary)
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()

        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"Primary + Secondary 均合法时连接应成功，实际：{result!r}"

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_failover_to_secondary_when_primary_invalid(self, azure_page, azure_cfg):
        """Primary 失效（无效凭据），系统应自动切换 Secondary 并继续上报"""
        azure = azure_cfg["azure_iot"]
        secondary = azure.get("secondary_conn_str", "")
        if not secondary:
            pytest.skip("config.yaml 未配置 secondary_conn_str，跳过 failover 测试")

        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(_INVALID_PRIMARY)
        azure_page.set_secondary_conn_str(secondary)
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()

        # 等待 Primary 超时后系统切换到 Secondary（留 60s 重连预算）
        time.sleep(60)

        ok = run_verifier(azure_cfg, timeout=120, skip_web=True)
        assert ok, (
            "Primary 失效后系统应自动切换到 Secondary Connection String，"
            "Azure IoT Hub 仍能收到设备数据"
        )
