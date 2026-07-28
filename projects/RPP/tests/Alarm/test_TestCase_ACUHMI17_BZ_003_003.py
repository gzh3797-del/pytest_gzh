from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_003_003
# 用例标题：Disable 状态下告警激活持续蜂鸣无确认功能
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 将 Alarm Acknowledgement Enable 设为关闭
#   2. 触发一条告警
#   3. 观察蜂鸣器状态及 Unacknowledged Alarms 页面
# 预期结果：
#   3. 告警激活后蜂鸣器持续发声（听觉验证，需人工现场确认，脚本不断言）；
#      Unacknowledged Alarms 页面无 Acknowledge 操作入口或入口不可用

_LABEL = "at_bz003003"


def test_TestCase_ACUHMI17_BZ_003_003(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 确认开关置为 Disable ──
        ha.set_ack_enable(page, False)

        # ── Step 2: 触发一条告警 ──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_log_rows(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，Alarm Logs 未在轮询周期内出现记录"

        # ── Step 3: 无确认入口（实测 Disable 时 Unacknowledged Alarms
        #    二级 tab 整体隐藏，即入口不可达；若 tab 仍在则按钮必须不可用）──
        if ha.unack_tab_present(page):
            ha.goto_global_alarm(page, "Unacknowledged Alarms")
            row = ha.data_rows(page, _LABEL)
            if row.count() > 0:
                ack_btn = row.first.get_by_role("button", name="Acknowledge")
                assert ack_btn.count() == 0 or not ack_btn.first.is_enabled(), (
                    "Disable 状态下告警行不应提供可用的 Acknowledge 按钮"
                )
            ack_all = page.get_by_role("button", name="Ack All Alarms")
            assert ack_all.count() == 0 or not ack_all.first.is_enabled(), (
                "Disable 状态下不应提供可用的 Ack All Alarms 按钮"
            )
        # 蜂鸣器持续发声——物理蜂鸣需人工确认，脚本不做断言
    finally:
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
