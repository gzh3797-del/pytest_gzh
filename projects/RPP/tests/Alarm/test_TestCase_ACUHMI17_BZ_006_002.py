from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_006_002
# 用例标题：Enable/Disable 切换时历史告警 Ack Status 展示行为一致
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. Enable 状态下触发并确认一条告警，Alarm Log 记录 Acknowledged
#   2. 切换为 Disable，再切回 Enable
#   3. 查看之前已确认告警的 Alarm Log 记录
# 预期结果：
#   3. 切换前已记录的 Ack Status（Acknowledged）在切回 Enable 后仍正确展示，
#      历史日志不丢失、不错乱

_LABEL = "at_bz006002"


def test_TestCase_ACUHMI17_BZ_006_002(app_page: Page):
    page = app_page
    try:
        # ── Step 1: Enable 下触发并确认一条告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"
        ha.ack_alarm(page, _LABEL)
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True), \
            f"确认后 Alarm Logs 中 {_LABEL!r} 应出现 Acknowledged 记录"

        ha.goto_global_alarm(page, "Alarm Logs")
        total_before = ha.data_rows(page, _LABEL).count()
        acked_before = ha.count_rows_by_ack(page, _LABEL, acknowledged=True)

        # ── Step 2: Disable → Enable 来回切换 ──
        ha.set_ack_enable(page, False)
        ha.set_ack_enable(page, True)

        # ── Step 3: 历史 Acknowledged 记录不丢失、不错乱 ──
        ha.goto_global_alarm(page, "Alarm Logs")
        assert "Ack Status" in ha.table_headers(page), \
            "切回 Enable 后 Alarm Logs 应重新显示 Ack Status 列"
        total_after = ha.data_rows(page, _LABEL).count()
        acked_after = ha.count_rows_by_ack(page, _LABEL, acknowledged=True)
        assert total_after >= total_before, (
            f"切换开关后 {_LABEL!r} 历史日志条数不应减少"
            f"（切换前 {total_before}，切换后 {total_after}）"
        )
        assert acked_after >= acked_before, (
            f"切换开关后 {_LABEL!r} 的 Acknowledged 记录不应丢失"
            f"（切换前 {acked_before}，切换后 {acked_after}）"
        )
    finally:
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
