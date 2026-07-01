"""
FTS编号: FTS_AcuHMI_AZR_004_004
用例标题: AcuHMI-1-7选择AcuvimIIR设备连接 Azure IoT Hub 数据上报（三段式比较）
用例级别: LV1

预置条件:
  - AcuvimIIR 设备已接入网关并在线，合法 Connection String 已配置

测试步骤:
  1. 仅勾选 AcuvimIIR 设备，Parameter Type=AI，Parameter=ALL
  2. 保存并 Test Connection
  3. 三段式验证：参数单位 vs 模板，参数值 vs Modbus 实时读数

预期结果:
  - 连接成功，上报参数单位与模板一致，参数值与 Modbus 读数误差在容差内
"""

import copy
from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

DEVICE_NAME      = "AcuvimIIR"
EXPECTED_DEVICES = ["AcuvimIIR"]


class TestCase_AcuHMI_1_7_AZR_004_004:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_acuvim_iir_connect_and_verify(self, azure_page, azure_cfg):
        """配置 AcuvimIIR → Test Connection → 三段式验证"""
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.select_only_device(DEVICE_NAME)
        azure_page.configure_all_devices_parameters(checked_only=True)
        azure_page.save()

        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"{DEVICE_NAME} 连接应成功，实际：{result!r}"

        cfg = copy.deepcopy(azure_cfg)
        cfg["azure_iot"]["expected_devices"] = EXPECTED_DEVICES
        ok = run_verifier(cfg, timeout=120, skip_web=True)
        assert ok, "三段式验证失败，请查看 reports/ 下的 HTML 报告"
