"""
FTS编号: FTS_WEB2_AWS_006_001
用例标题: 禁用 AWS IoT 后停止发布，重新启用后恢复
用例级别: LV3

预置条件: AWS IoT 已正常配置并连接，MQTT 订阅端就绪

测试步骤:
  1. 确认 AWS IoT 正在发布数据（Test Connection 成功）
  2. 切换到 Disable 并保存，等待约 2 个 Interval，验证无新消息到达
  3. 重新完整配置 + Enable + 保存，验证 MQTT 消息恢复到达

预期结果:
  - Disable 后 MQTT 不再收到新消息
  - Enable 后 MQTT 恢复收到数据
"""

import copy
import time
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent


class TestCase_AcuHMI_1_7_AWS_006_001:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_disable_and_reenable(self, aws_page, aws_cfg):
        """配置 AWS IoT → Disable 验证无数据 → 重新配置 Enable 验证数据恢复"""
        aws = aws_cfg["aws_iot"]

        interval_s = int(str(aws.get("interval", "30 seconds")).split()[0])

        def _configure():
            """完整配置一次 AWS IoT，随机选一台物理设备，返回实际选中的设备名列表。"""
            aws_page.ensure_enabled()
            aws_page.set_url(aws["url"])
            aws_page.set_interval(aws.get("interval", "30 seconds"))
            aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
            aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
            aws_page.select_only_device()  # 随机选一台物理设备
            aws_page.configure_all_devices_parameters(checked_only=True)
            aws_page.save()
            return aws_page.get_checked_device_names()

        # ── Step 1: 初始配置，确认连接成功 ──────────────────────────────────
        selected = _configure()
        single_cfg = copy.deepcopy(aws_cfg)
        single_cfg["aws_iot"]["expected_devices"] = selected
        conn = aws_page.test_connection()
        assert any(kw in conn.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"初始连接应成功，实际：{conn!r}"

        # ── Step 2: Disable → 等待 2 个 Interval → 验证无新消息 ─────────────
        aws_page.disable()
        assert not aws_page.is_enabled(), "Disable 操作后 is_enabled 应为 False"

        wait_secs = interval_s * 2 + 10
        time.sleep(wait_secs)

        ok_disabled = run_verifier(single_cfg, timeout=interval_s + 10, skip_web=True)
        assert not ok_disabled, "Disable 后 AWS IoT 不应有新消息到达"

        # ── Step 3: 重新完整配置 Enable → 验证数据恢复 ──────────────────────
        selected = _configure()
        single_cfg["aws_iot"]["expected_devices"] = selected
        conn2 = aws_page.test_connection()
        assert any(kw in conn2.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"重新 Enable 后连接应成功，实际：{conn2!r}"

        ok_enabled = run_verifier(single_cfg, timeout=120, skip_web=True)
        assert ok_enabled, "重新 Enable 后 MQTT 数据应恢复"
