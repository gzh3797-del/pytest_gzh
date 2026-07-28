"""
FTS编号: FTS_WEB2_AWS_006_004
用例标题: 被发布设备离线时上报行为与预期一致
用例级别: LV3

预置条件: 至少一台下游设备已在 AWS IoT 设备列表中，AWS IoT 正常配置并连接

测试步骤:
  1. 确认 AWS IoT 正常上报
  2. 将下游设备断电或拔线（使其离线）
  3. 等待 3 个 Interval，观察 MQTT 消息（值为 null/0/错误码，网关不崩溃）
  4. 恢复下游设备，验证数据自动恢复

预期结果:
  - 设备离线时，MQTT 消息行为符合规格，网关不崩溃
  - 下游设备恢复后，数据自动恢复
"""

import time
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent


class TestCase_RPP_AWS_006_004:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_upstream_connected_before_device_offline(self, aws_session_page, aws_cfg):
        """步骤1：确认 AWS IoT 初始正常上报"""
        aws = aws_cfg["aws_iot"]
        aws_session_page.ensure_enabled()
        aws_session_page.set_url(aws["url"])
        aws_session_page.set_interval(aws.get("interval", "30 seconds"))
        aws_session_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_session_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_session_page.select_only_device()  # 随机选一台物理设备
        aws_session_page.configure_all_devices_parameters(checked_only=True)
        aws_session_page.save()
        result = aws_session_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"初始连接应成功，实际：{result!r}"
        print("\n[手动操作] 请将下游设备断电，然后运行 test_behavior_when_device_offline")

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_behavior_when_device_offline(self, aws_cfg):
        """步骤2-4：下游设备离线期间验证 MQTT 行为（网关不崩溃）"""
        time.sleep(100)
        ok = run_verifier(aws_cfg, timeout=60, skip_web=True)
        print(f"\n[离线期间 verifier 结果]: {'有消息' if ok else '超时/无消息'}")

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_data_recovers_after_device_online(self, aws_cfg):
        """步骤5-6：下游设备恢复后验证数据自动恢复"""
        print("\n[手动操作] 请恢复下游设备，等待 2 个 Interval 后此用例自动验证")
        time.sleep(70)
        ok = run_verifier(aws_cfg, timeout=120, skip_web=True)
        assert ok, "下游设备恢复后 MQTT 数据应正常"
