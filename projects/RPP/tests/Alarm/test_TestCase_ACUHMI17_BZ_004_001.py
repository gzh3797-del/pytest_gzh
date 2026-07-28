from playwright.sync_api import Page

from projects.RPP.tests.Alarm import config_alarm as cfg
from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_004_001
# 用例标题：Unacknowledged Alarms 页面 8 列完整展示
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 触发至少一条告警（不确认）
#   2. 进入 Unacknowledged Alarms 页面
#   3. 查看页面列表列头和各列内容
# 预期结果：
#   3. 页面展示 8 列：Timestamp / Device Name / Serial Number / Monitor Label /
#      Parameter / Status / Reason / Action，各列内容与告警信息一致

_LABEL = "at_bz004001"
_EXPECTED_COLUMNS = ["Timestamp", "Device Name", "Serial Number",
                     "Monitor Label", "Parameter", "Status", "Reason", "Action"]


def test_TestCase_ACUHMI17_BZ_004_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 触发一条告警（不确认）──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        # ── Step 2-3: 列头 8 列完整、顺序一致 ──
        headers = ha.table_headers(page)
        assert headers == _EXPECTED_COLUMNS, (
            f"Unacknowledged Alarms 列头应为 {_EXPECTED_COLUMNS}，实际: {headers}"
        )

        # 各列内容与告警信息一致
        row = ha.data_rows(page, _LABEL).first
        assert ha.cell_text(page, row, "Device Name") == cfg.TRIGGER_DEVICE, \
            "Device Name 列应为触发告警的设备名"
        assert ha.cell_text(page, row, "Monitor Label") == _LABEL, \
            "Monitor Label 列应为告警规则 Label"
        assert ha.cell_text(page, row, "Timestamp"), "Timestamp 列不应为空"
        reason = ha.cell_text(page, row, "Reason").upper()
        assert reason in ("UNDERFLOW", "OVERFLOW"), \
            f"Reason 列应为 UNDERFLOW/OVERFLOW，实际: {reason!r}"
        ack_btn = row.get_by_role("button", name="Acknowledge")
        assert ack_btn.count() > 0, "Action 列应提供 Acknowledge 操作"
    finally:
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
