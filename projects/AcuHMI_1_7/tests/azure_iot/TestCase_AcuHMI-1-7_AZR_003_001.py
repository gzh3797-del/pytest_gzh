"""
FTS编号: FTS_AcuHMI_AZR_003_001
用例标题: 合法 Connection String 连接 Azure IoT Hub
用例级别: LV0

预置条件:
  - 网关已登录，Event Hub 订阅端就绪
  - tests/protocols/azure_iot/config.yaml 中已配置合法的 primary_conn_str

测试步骤:
  1. Enable Azure IoT，填写合法 Primary Connection String、Interval
  2. 全选设备并配置参数，点击 Save
  3. 点击 Test Connection，等待结果
  4. 调用 azure_iot_verifier 验证 Azure IoT Hub 收到数据

预期结果:
  - Test Connection 返回连接成功标识
  - Azure IoT Hub Event Hub 端收到设备数据
"""

import pytest

from utils.azure_iot_verifier import run_verifier


class TestCase_AcuHMI_1_7_AZR_003_001:

    @pytest.mark.azure_iot
    def test_valid_conn_str_connection(self, azure_page, azure_cfg):
        """填写合法 Connection String → Save → Test Connection 应成功"""
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.save()

        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"Test Connection 应返回成功，实际：{result!r}"

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_event_hub_data_received(self, azure_cfg):
        """验证 Azure IoT Hub Event Hub 实际收到 MQTT 消息（四阶段验证）"""
        ok = run_verifier(azure_cfg, timeout=120, skip_web=True)
        assert ok, "Azure IoT 数据验证失败，请查看 reports/ 下的 HTML 报告"
