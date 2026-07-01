"""
FTS编号: FTS_AcuHMI_AZR_006_001
用例标题: 禁用 Azure IoT 后停止发布，重新启用后恢复
用例级别: LV3

预置条件: Azure IoT 已正常配置并连接，Event Hub 订阅端就绪

测试步骤:
  1. 确认 Azure IoT 正在发布数据
  2. 切换到 Disable 并保存，等待约 2 个 Interval，验证无新消息到达
  3. 切换回 Enable 并保存，验证消息恢复到达

预期结果:
  - Disable 后 Azure IoT Hub 不再收到新消息
  - Enable 后 Azure IoT Hub 恢复收到数据
"""

import time
from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent


class TestCase_AcuHMI_1_7_AZR_006_001:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_disable_and_reenable(self, azure_page, azure_cfg):
        """配置 Azure IoT → Test Connection → Disable验证无数据 → Re-enable验证数据恢复"""
        azure = azure_cfg["azure_iot"]

        # Step 1: 配置并确认初始连接
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()
        conn = azure_page.test_connection()
        assert any(kw in conn.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"初始连接应成功，实际：{conn!r}"

        # Step 2: Disable → 等待约 2 个 Interval → 验证无新消息
        azure_page.disable()
        assert not azure_page.is_enabled()

        time.sleep(70)
        ok_disabled = run_verifier(azure_cfg, timeout=30, skip_web=True)
        assert not ok_disabled, "Disable 后 Azure IoT Hub 不应有新消息到达"

        # Step 3: Re-enable → 保存 → 等待数据恢复
        azure_page.ensure_enabled()
        azure_page.save()
        conn2 = azure_page.test_connection()
        assert any(kw in conn2.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"重新 Enable 后连接应成功，实际：{conn2!r}"

        ok_enabled = run_verifier(azure_cfg, timeout=120, skip_web=True)
        assert ok_enabled, "重新 Enable 后 Azure IoT 数据应恢复"
