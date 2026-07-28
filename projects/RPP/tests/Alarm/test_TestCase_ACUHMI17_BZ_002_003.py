from playwright.sync_api import Page

from projects.RPP.tests.Alarm import config_alarm as cfg
from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_002_003
# 用例标题：报警关闭时蜂鸣器不发声
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 触发告警并确认
#   2. 使告警条件消除（参数恢复正常），告警变为关闭状态
#   3. 观察蜂鸣器状态
# 预期结果：
#   3. 告警关闭后蜂鸣器不发声，无论告警是否曾被确认
#      （蜂鸣为听觉验证需人工确认；脚本断言告警状态确已转为 OFF）

_LABEL = "at_bz002003"
# 恢复正常的阈值区间：覆盖任何真实读数，告警条件必然消除
_NORMAL_MIN = "-2000000000"
_NORMAL_MAX = "2000000000"


def test_TestCase_ACUHMI17_BZ_002_003(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 触发告警并确认 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"
        ha.ack_alarm(page, _LABEL)

        # ── Step 2: 阈值改为全量程区间，使告警条件消除 ──
        ha.goto_device_alarm(page, "Alarm Config")
        ha.edit_alarm_rule_range(page, _LABEL, _NORMAL_MIN, _NORMAL_MAX)

        # 等待下一轮询周期告警转为 OFF（Status 列图标 warning→success）
        alarm_off = False
        for _ in range(cfg.POLL_ROUNDS):
            if not ha.rule_alarm_on(page, _LABEL):
                alarm_off = True
                break
            page.wait_for_timeout(cfg.POLL_STEP_MS)
            page.reload()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(800)
        assert alarm_off, (
            f"阈值恢复正常后，{_LABEL!r} 的告警状态应在轮询周期内转为 OFF"
        )

        # ── Step 3: 蜂鸣器不发声——物理蜂鸣需人工确认，脚本不做断言 ──
    finally:
        ha.cleanup_test_rules(page)
