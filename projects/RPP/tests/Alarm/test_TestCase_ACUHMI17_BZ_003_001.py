from playwright.sync_api import Page, expect

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_003_001
# 用例标题：Enable 状态下 Alarm Logs 显示 Ack Status 列
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 将 Alarm Acknowledgement Enable 设为开启
#   2. 触发一条告警，进入 Alarm Logs 页面
#   3. 查看列表列头和 Ack Status 列内容
# 预期结果：
#   3. Alarm Logs 显示 Ack Status 列，未确认告警显示 Unacknowledge，
#      已确认显示 Acknowledged

_LABEL = "at_bz003001"


def test_TestCase_ACUHMI17_BZ_003_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 确认开关置为 Enable ──
        ha.set_ack_enable(page, True)

        # ── Step 2: 触发一条告警 ──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        # ── Step 3: Alarm Logs 显示 Ack Status 列，未确认显示 Unacknowledge ──
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=False), \
            f"全局 Alarm Logs 中应存在 {_LABEL!r} 的 Unacknowledge 记录"
        expect(page.locator(".el-table__header-wrapper th").filter(
            has_text="Ack Status")).to_be_visible()

        # 确认后同一告警显示 Acknowledged
        ha.ack_alarm(page, _LABEL)
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True), \
            f"确认后全局 Alarm Logs 中 {_LABEL!r} 应出现 Acknowledged 记录"
    finally:
        ha.cleanup_test_rules(page)
