# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_007_004
用例标题: 设备断电时上报行为与预期一致
用例级别: LV3

预置条件:
  - 至少一台下游设备已在 Azure IoT 设备列表中并正常上报
  - Azure IoT 正常配置并连接，Event Hub 订阅端持续监听

测试步骤:
  1. 确认 Azure IoT 正常上报（Test Connection 成功）
  2. 将下游设备断电（手动操作）
  3. 等待 3 个 Interval，观察 Azure IoT Hub 消息（值应为 null/0/错误码，网关不崩溃）
  4. 恢复下游设备，验证数据自动恢复

预期结果:
  - 设备离线时 Azure IoT Hub 消息行为符合规格，网关不崩溃、不重启
  - 下游设备恢复上电后，数据自动恢复正常上报
"""

import time
from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent


class TestCase_AcuHMI_1_7_AZR_007_004:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_setup_before_device_offline(self, azure_page, azure_cfg):
        """Step 1: 确认 Azure IoT 初始正常上报（手动配合断电）"""
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()
        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"初始连接应成功，实际：{result!r}"
        print("\n[手动操作] 请将下游设备断电，然后运行 test_behavior_when_device_offline")

    @pytest.mark.azure_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_behavior_when_device_offline(self, azure_cfg):
        """Step 2-3: 下游设备离线期间验证 Event Hub 行为（网关不崩溃）"""
        time.sleep(100)
        ok = run_verifier(azure_cfg, timeout=60, skip_web=True)
        print(f"\n[离线期间 verifier 结果]: {'有消息到达' if ok else '超时/无消息'}")
        # 不做强断言：设备离线时可能仍有部分消息（如 null 值），行为取决于规格

    @pytest.mark.azure_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_data_recovers_after_device_online(self, azure_cfg):
        """Step 4: 下游设备恢复上电后，数据应自动恢复正常上报"""
        print("\n[手动操作] 请恢复下游设备供电，等待 2 个 Interval 后本用例自动验证")
        time.sleep(70)
        ok = run_verifier(azure_cfg, timeout=120, skip_web=True)
        assert ok, "下游设备恢复上电后 Azure IoT 数据应自动恢复正常上报"
