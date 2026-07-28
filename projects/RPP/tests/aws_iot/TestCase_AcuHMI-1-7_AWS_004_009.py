"""
FTS编号: FTS_AcuHMI_AWS_004_009
用例标题: 未配置参数时设备不上报数据

预置条件:
  - Enable 已开启，合法配置，所有设备已勾选

测试步骤:
  1. Enable AWS IoT，勾选所有设备
  2. 为每台设备打开参数配置弹窗，点击 Clear → Confirm（清除所有已选参数）
  3. 配置连接参数（URL / Interval / 证书）并 Save
  4. Test Connection（失败时等 2s 重试，最多 3 次）
  5. 等待 2 个 Interval 周期，订阅 MQTT 计算收到的消息数
  6. 恢复：重新为所有设备配置全量参数并 Save，还原测试前状态

预期结果:
  - 未配置任何参数的设备不向 AWS IoT 上报数据（消息数 = 0）
  - 测试完成后参数配置自动恢复
"""

import time
from pathlib import Path

import pytest

from utils.aws_iot_verifier import count_mqtt_messages

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent


class TestCase_RPP_AWS_004_009:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_no_params_no_data(self, aws_page, aws_cfg):
        """勾选所有设备、清空参数配置，验证 AWS IoT 无数据上报。"""
        aws = aws_cfg["aws_iot"]
        interval_str = aws.get("interval", "30 seconds")
        interval_secs = int(str(interval_str).split()[0])

        # ── 1. Enable + 勾选所有设备 ───────────────────────────────────────────
        aws_page.ensure_enabled()
        aws_page.select_unchecked_devices()

        # ── 2. 清除所有设备的参数配置 ──────────────────────────────────────────
        aws_page.clear_all_devices_parameters(checked_only=True)

        # ── 3. 配置连接参数并保存 ──────────────────────────────────────────────
        aws_page.set_url(aws["url"])
        aws_page.set_interval(interval_str)
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.save()

        # ── 4. Test Connection（失败时等 2s 重试，最多 3 次）────────────────────
        _RETRY_MAX = 3
        result = ""
        for _attempt in range(1, _RETRY_MAX + 1):
            result = aws_page.test_connection()
            _success = any(kw in result.lower() for kw in ("success", "connected", "ok", "成功"))
            print(f"\n[无参数配置 Test Connection 第{_attempt}次]: {result!r}")
            if _success:
                break
            if _attempt < _RETRY_MAX:
                time.sleep(2)

        # ── 5. 订阅 MQTT，监测 2 个 Interval 内是否有数据上报 ──────────────────
        monitor_secs = interval_secs * 2 + 10
        print(f"\n[监测] 订阅 {monitor_secs}s，期望收到 0 条消息…")
        msg_count = count_mqtt_messages(aws_cfg, timeout=monitor_secs)
        print(f"[监测结果] 共收到 {msg_count} 条消息")

        # ── 6. 恢复：重新为所有设备配置全量参数并保存 ─────────────────────────
        print("\n[恢复] 重新配置所有设备全量参数…")
        aws_page.navigate_to_aws_iot()
        aws_page.ensure_enabled()
        aws_page.configure_all_devices_parameters(checked_only=True)
        aws_page.save()
        print("[恢复] 参数配置已还原")

        # ── 断言 ───────────────────────────────────────────────────────────────
        assert msg_count == 0, (
            f"未配置参数的设备不应上报数据，"
            f"实际在 {monitor_secs}s 内收到 {msg_count} 条消息"
        )
