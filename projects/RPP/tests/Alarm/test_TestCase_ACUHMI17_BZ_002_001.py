from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_002_001
# 用例标题：报警开启且未确认时触发蜂鸣
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 确保 Alarm Acknowledgement Enable 为开启状态
#   2. 触发一条告警（使监控参数超过告警阈值）
#   3. 不进行任何确认操作，观察蜂鸣器状态
#   4. 查看 Alarm Log 中该告警的记录
# 预期结果：
#   3. 蜂鸣器持续发声（听觉验证，需人工现场确认，脚本不断言）
#   4. Alarm Log 中记录该告警已被触发，Ack Status 为 Unacknowledge

_LABEL = "at_bz002001"


def test_TestCase_ACUHMI17_BZ_002_001(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 确认开关置为 Enable ──
        ha.set_ack_enable(page, True)

        # ── Step 2: 触发一条告警（必越限规则，等一个轮询周期）──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        # ── Step 3: 蜂鸣器持续发声——物理蜂鸣需人工确认，脚本不做断言 ──

        # ── Step 4: Alarm Log 中该告警 Ack Status 为 Unacknowledge ──
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=False), (
            f"全局 Alarm Logs 中应存在 {_LABEL!r} 的 Unacknowledge 记录"
        )
    finally:
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
