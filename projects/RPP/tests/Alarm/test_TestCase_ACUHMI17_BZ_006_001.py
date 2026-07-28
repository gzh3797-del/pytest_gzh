from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_006_001
# 用例标题：多条未确认告警确认其中一条不影响其他蜂鸣
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 同时触发 3 条告警，均处于未确认蜂鸣状态
#   2. 在 Unacknowledged Alarms 页面确认其中第 1 条
#   3. 观察蜂鸣器状态及剩余未确认告警
# 预期结果：
#   2. 第 1 条告警从 Unacknowledged Alarms 消失，Alarm Log 记录确认
#   3. 蜂鸣器持续发声（因仍有 2 条未确认告警，听觉验证需人工确认）；
#      其余告警状态不变

_LABELS = ("at_bz006001a", "at_bz006001b", "at_bz006001c")


def test_TestCase_ACUHMI17_BZ_006_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 同时触发 3 条告警（三个不同参数的必越限规则）──
        ha.set_ack_enable(page, True)
        for idx, label in enumerate(_LABELS):
            ha.trigger_alarm(page, label, param_idx=idx)
        for label in _LABELS:
            assert ha.wait_for_unack(page, label), \
                f"规则 {label!r} 未在轮询周期内触发告警"

        # ── Step 2: 确认第 1 条 ──
        ha.ack_alarm(page, _LABELS[0])

        ha.goto_global_alarm(page, "Unacknowledged Alarms")
        assert ha.data_rows(page, _LABELS[0]).count() == 0, \
            f"已确认告警 {_LABELS[0]!r} 应从 Unacknowledged Alarms 消失"
        assert ha.wait_for_log_status(page, _LABELS[0], acknowledged=True), \
            f"Alarm Logs 中应记录 {_LABELS[0]!r} 的确认日志"

        # ── Step 3: 其余 2 条未确认告警状态不变 ──
        ha.goto_global_alarm(page, "Unacknowledged Alarms")
        for label in _LABELS[1:]:
            assert ha.data_rows(page, label).count() > 0, \
                f"未确认告警 {label!r} 应仍保留在 Unacknowledged Alarms 中"
        # 蜂鸣器持续发声——物理蜂鸣需人工确认，脚本不做断言
    finally:
        for label in _LABELS[1:]:
            try:
                ha.ack_alarm(page, label)
            except Exception:
                pass
        ha.cleanup_test_rules(page)
