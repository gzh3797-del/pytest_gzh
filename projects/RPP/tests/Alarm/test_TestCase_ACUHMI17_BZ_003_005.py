from playwright.sync_api import Page

from projects.RPP.tests.Alarm import config_alarm as cfg
from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_003_005
# 用例标题：Enable → Disable 切换后告警激活持续蜂鸣
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. Alarm Acknowledgement Enable 为开启，存在已确认告警（蜂鸣关闭）
#   2. 再触发一条新告警（未确认）
#   3. 将开关切换为关闭
#   4. 观察蜂鸣器状态；使告警条件消除，再次观察
# 预期结果：
#   3. 切换后告警处于 ON 状态，蜂鸣器持续发声（蜂鸣需人工确认）
#   4. 告警条件消除（告警 OFF）后蜂鸣器停止（蜂鸣需人工确认；
#      脚本断言告警 ON→OFF 的状态迁移）

_LABEL_ACKED = "at_bz003005a"
_LABEL_NEW = "at_bz003005b"
_NORMAL_MIN = "-2000000000"
_NORMAL_MAX = "2000000000"


def test_TestCase_ACUHMI17_BZ_003_005(app_page: Page):
    page = app_page
    try:
        # ── Step 1: Enable 状态下制造一条已确认告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL_ACKED, param_idx=0)
        assert ha.wait_for_unack(page, _LABEL_ACKED), \
            f"规则 {_LABEL_ACKED!r} 未在轮询周期内触发告警"
        ha.ack_alarm(page, _LABEL_ACKED)

        # ── Step 2: 再触发一条新告警（未确认）──
        ha.trigger_alarm(page, _LABEL_NEW, param_idx=1)
        assert ha.wait_for_unack(page, _LABEL_NEW), \
            f"规则 {_LABEL_NEW!r} 未在轮询周期内触发告警"

        # ── Step 3: 切换为 Disable，新告警仍处于 ON 状态 ──
        ha.set_ack_enable(page, False)
        ha.goto_device_alarm(page, "Alarm Config")
        assert ha.rule_alarm_on(page, _LABEL_NEW), (
            f"切换 Disable 后告警 {_LABEL_NEW!r} 应保持 ON 状态"
        )
        # 蜂鸣器持续发声——物理蜂鸣需人工确认，脚本不做断言

        # ── Step 4: 消除告警条件，告警应转为 OFF ──
        ha.edit_alarm_rule_range(page, _LABEL_NEW, _NORMAL_MIN, _NORMAL_MAX)
        alarm_off = False
        for _ in range(cfg.POLL_ROUNDS):
            if not ha.rule_alarm_on(page, _LABEL_NEW):
                alarm_off = True
                break
            page.wait_for_timeout(cfg.POLL_STEP_MS)
            page.reload()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(800)
        assert alarm_off, (
            f"告警条件消除后，{_LABEL_NEW!r} 应在轮询周期内转为 OFF"
        )
        # 蜂鸣器停止——物理蜂鸣需人工确认，脚本不做断言
    finally:
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
