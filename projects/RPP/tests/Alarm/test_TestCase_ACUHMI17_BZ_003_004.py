from playwright.sync_api import Page, expect

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_003_004
# 用例标题：Disable → Enable 切换后已有未确认告警可被确认
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. Alarm Acknowledgement Enable 为关闭，触发一条告警
#   2. 将开关切换为开启
#   3. 进入 Unacknowledged Alarms 页面，点击该告警的 Acknowledge
#   4. 查看蜂鸣器状态和 Alarm Log
# 预期结果：
#   3. 切换后该告警可被确认，Acknowledge 按钮可用
#   4. 蜂鸣器停止（听觉验证，需人工现场确认，脚本不断言）；
#      Alarm Log 新增确认日志，Ack Status 为 Acknowledged

_LABEL = "at_bz003004"


def test_TestCase_ACUHMI17_BZ_003_004(app_page: Page):
    page = app_page
    try:
        # ── Step 1: Disable 状态下触发一条告警 ──
        ha.set_ack_enable(page, False)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_log_rows(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，Alarm Logs 未在轮询周期内出现记录"

        # ── Step 2: 切换为 Enable ──
        ha.set_ack_enable(page, True)

        # ── Step 3: 该告警可被确认，Acknowledge 按钮可用 ──
        ha.goto_global_alarm(page, "Unacknowledged Alarms")
        row = ha.data_rows(page, _LABEL).first
        row.wait_for(timeout=10_000)
        ack_btn = row.get_by_role("button", name="Acknowledge")
        expect(ack_btn.first).to_be_enabled()
        ha.ack_alarm(page, _LABEL)

        # ── Step 4: Alarm Log 新增确认日志 ──
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True), \
            f"确认后全局 Alarm Logs 中 {_LABEL!r} 应出现 Acknowledged 记录"
    finally:
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
