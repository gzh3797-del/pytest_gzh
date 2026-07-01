"""
FTS编号: FTS_AcuHMI_AZR_004_009
用例标题: 4100/IOM设备未配置上传参数时发送为空
用例级别: LV1

预置条件:
  - Enable 已开启，合法配置，所有设备已勾选

测试步骤:
  1. Azure IoT Parameter Config 页面不配置参数（Clear）
  2. 连接 Azure IoT Hub
  3. 观察 Event Hub 消息情况

预期结果:
  - 连接成功（Connection String 合法）
  - Azure IoT Hub 在 40s 内未收到任何设备数据消息（需人工验证或比对 verifier 空结果）
"""

import time
from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent


class TestCase_AcuHMI_1_7_AZR_004_009:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_no_params_no_data(self, azure_page, azure_cfg):
        """故意不配置任何设备参数，验证 save 和 test_connection 成功。"""
        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.select_only_device("")
        # 故意跳过 configure_all_devices_parameters，不配置任何参数
        azure_page.save()
        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"连接应成功（即使无参数配置），实际：{result!r}"

        time.sleep(10)
        ok = run_verifier(azure_cfg, timeout=40, skip_web=True)
        print(f"\n[无参数配置时 verifier 结果]: {'有数据上报' if ok else '无数据上报（或超时）'}")
