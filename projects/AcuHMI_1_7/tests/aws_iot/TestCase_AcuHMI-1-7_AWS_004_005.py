"""
FTS编号: FTS_AcuHMI_AWS_004_005
用例标题: AcuHMI-1-7选择AcuvimIIW设备连接 AWS IoT Core 数据上报（三段式比较）
用例级别: LV1

预置条件:
  - AcuvimIIW 设备已接入网关并在线，合法证书已配置，MQTT 订阅端就绪

测试步骤:
  1. 仅勾选 AcuvimIIW 设备，Parameter Type=AI，Parameter=ALL
  2. 保存并 Test Connection
  3. 三段式验证：参数单位 vs MQTT 模板，参数值 vs Modbus 实时读数，设备信息/时间戳

预期结果:
  - 连接成功，上报参数单位与模板一致，参数值与 Modbus 读数误差在容差内
"""

import copy
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent

DEVICE_NAME      = "AcuvimIIW"
EXPECTED_DEVICES = ["AcuvimIIW"]


class TestCase_AcuHMI_1_7_AWS_004_005:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_acuvim_iiw_connect_and_verify(self, aws_page, aws_cfg):
        """配置 AcuvimIIW → Test Connection → 三段式验证，一次浏览器完成"""
        aws = aws_cfg["aws_iot"]
        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.set_interval(aws.get("interval", "30 seconds"))
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.select_only_device(DEVICE_NAME, exact=True)
        aws_page.configure_all_devices_parameters(checked_only=True)
        aws_page.save()

        result = aws_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"{DEVICE_NAME} 连接应成功，实际：{result!r}"

        cfg = copy.deepcopy(aws_cfg)
        cfg["aws_iot"]["expected_devices"] = EXPECTED_DEVICES
        ok = run_verifier(cfg, timeout=120, skip_web=True)
        assert ok, "三段式验证失败，请查看 reports/ 下的 HTML 报告"
