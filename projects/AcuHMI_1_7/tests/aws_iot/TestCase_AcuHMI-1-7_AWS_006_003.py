"""
FTS编号: FTS_WEB2_AWS_006_003
用例标题: 多设备多参数同时发布系统性能正常
用例级别: LV3

预置条件: 所有下游设备均已接入并在线，合法证书已配置

测试步骤:
  1. 勾选全部设备，为每台设备配置 ALL 参数，设置 Interval=30 seconds
  2. 保存并连接
  3. 持续监听 10 分钟，监控 MQTT 消息按时到达及设备无重启

预期结果:
  - 所有设备数据按 Interval 正常上报，CPU/内存正常，无超时、丢包、重启
"""

from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

MONITOR_DURATION = 600


class TestCase_AcuHMI_1_7_AWS_006_003:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_multi_device_performance(self, aws_session_page, aws_cfg):
        aws = aws_cfg["aws_iot"]
        expected = aws.get("expected_devices", [])
        aws_session_page.ensure_enabled()
        aws_session_page.set_url(aws["url"])
        aws_session_page.set_interval(aws.get("interval", "30 seconds"))
        aws_session_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_session_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_session_page.select_devices(expected)
        aws_session_page.configure_all_devices_parameters(checked_only=True)
        aws_session_page.save()
        result = aws_session_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"多设备连接应成功，实际：{result!r}"

        ok = run_verifier(aws_cfg, timeout=MONITOR_DURATION, skip_web=True)
        assert ok, "多设备性能测试期间 MQTT 验证失败"

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_no_device_restart_during_publish(self, aws_session_page):
        """验证多设备发布期间 Web UI 可访问（设备未重启）"""
        aws_session_page.navigate_to_aws_iot()
        assert aws_session_page.is_enabled()
        result = aws_session_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"长时间运行后 Test Connection 仍应成功：{result!r}"
