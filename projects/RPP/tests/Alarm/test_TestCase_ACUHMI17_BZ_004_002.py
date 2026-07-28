from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_004_002
# 用例标题：Acknowledge 后条目从 Unacknowledged Alarms 消失
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 未确认告警存在于 Unacknowledged Alarms 列表中
#   2. 点击该告警 Action 列的 Acknowledge 按钮
#   3. 查看 Unacknowledged Alarms 页面
# 预期结果：
#   3. 已确认的告警从 Unacknowledged Alarms 列表中消失，列表其余条目不受影响

_LABEL_TARGET = "at_bz004002a"
_LABEL_OTHER = "at_bz004002b"


def test_TestCase_ACUHMI17_BZ_004_002(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 制造两条未确认告警（一条被确认、一条对照）──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL_TARGET, param_idx=0)
        ha.trigger_alarm(page, _LABEL_OTHER, param_idx=1)
        assert ha.wait_for_unack(page, _LABEL_TARGET), \
            f"规则 {_LABEL_TARGET!r} 未在轮询周期内触发告警"
        assert ha.wait_for_unack(page, _LABEL_OTHER), \
            f"规则 {_LABEL_OTHER!r} 未在轮询周期内触发告警"

        # ── Step 2: 确认目标告警 ──
        ha.ack_alarm(page, _LABEL_TARGET)

        # ── Step 3: 目标条目消失，其余条目不受影响 ──
        ha.goto_global_alarm(page, "Unacknowledged Alarms")
        assert ha.data_rows(page, _LABEL_TARGET).count() == 0, \
            f"已确认告警 {_LABEL_TARGET!r} 应从 Unacknowledged Alarms 中消失"
        assert ha.data_rows(page, _LABEL_OTHER).count() > 0, \
            f"未确认告警 {_LABEL_OTHER!r} 应仍保留在列表中，不受确认操作影响"
    finally:
        try:
            ha.ack_alarm(page, _LABEL_OTHER)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
