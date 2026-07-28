from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_005_001
# 用例标题：Enable 状态下 Ack Status 列正确展示，Disable 时隐藏
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. Alarm Acknowledgement Enable 为开启，触发并确认一条告警，
#      进入 Alarm Logs 查看 Ack Status 列
#   2. 将开关切换为关闭，重新查看 Alarm Logs
# 预期结果：
#   1. Ack Status 列可见，已确认告警显示 Acknowledged，未确认显示 Unacknowledge
#   2. Ack Status 列隐藏，不显示

_LABEL = "at_bz005001"


def test_TestCase_ACUHMI17_BZ_005_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1: Enable 下触发一条告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        # 未确认时 Ack Status 列可见且显示 Unacknowledge
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=False), \
            f"Alarm Logs 中应存在 {_LABEL!r} 的 Unacknowledge 记录"
        assert "Ack Status" in ha.table_headers(page), \
            "Enable 状态下 Alarm Logs 应显示 Ack Status 列"

        # 确认后显示 Acknowledged
        ha.ack_alarm(page, _LABEL)
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True), \
            f"确认后 Alarm Logs 中 {_LABEL!r} 应出现 Acknowledged 记录"

        # ── Step 2: 切换 Disable 后 Ack Status 列隐藏 ──
        ha.set_ack_enable(page, False)
        ha.goto_global_alarm(page, "Alarm Logs")
        headers = ha.table_headers(page)
        assert "Ack Status" not in headers, (
            f"Disable 状态下 Alarm Logs 不应显示 Ack Status 列，实际列: {headers}"
        )
    finally:
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
