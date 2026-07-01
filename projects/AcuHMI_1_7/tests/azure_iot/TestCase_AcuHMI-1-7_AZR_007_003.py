# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_007_003
用例标题: 多设备同时连接时系统资源正常，上报不中断
用例级别: LV2

预置条件:
  - 所有下游设备均已接入并在线
  - 合法 Connection String 已配置，Interval = 30 seconds

测试步骤:
  1. 勾选全部设备，为每台设备配置 ALL 参数，Interval = 30 seconds，保存连接
  2. 持续监听 10 分钟，验证所有设备按 Interval 正常上报
  3. 验证 Web UI 可正常访问（设备未重启）

预期结果:
  - 所有设备数据按 Interval 正常上报
  - CPU/内存等系统资源正常，无超时、丢包、重启
"""

from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

_MONITOR_DURATION = 600


class TestCase_AcuHMI_1_7_AZR_007_003:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_multi_device_all_params_upload(self, azure_page, azure_cfg):
        """全部设备 + ALL 参数同时上报，持续 10 分钟不中断"""
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_unchecked_devices()
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()

        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"多设备连接应成功，实际：{result!r}"

        ok = run_verifier(azure_cfg, timeout=_MONITOR_DURATION, skip_web=True)
        assert ok, "多设备同时上报 10 分钟内 Azure IoT 验证失败"

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_web_ui_accessible_during_multi_device_publish(self, azure_page):
        """多设备发布期间 Web UI 可正常访问（设备未重启）"""
        azure_page.navigate_to_azure_iot()
        assert azure_page.is_enabled(), "多设备发布期间 Azure IoT 应保持 Enable 状态"
        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"长时间运行后 Test Connection 仍应成功：{result!r}"
