from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_002_002
# 用例标题：告警开启后确认操作使蜂鸣停止
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 已有蜂鸣器发声的未确认告警
#   2. 在 Unacknowledged Alarms 页面点击该告警的 Acknowledge
#   3. 观察蜂鸣器状态
#   4. 查看 Alarm Log 中该告警的记录
# 预期结果：
#   3. 蜂鸣器停止发声（听觉验证，需人工现场确认，脚本不断言）
#   4. Alarm Log 新增一条确认日志，Ack Status 变为 Acknowledged

_LABEL = "at_bz002002"


def test_TestCase_ACUHMI17_BZ_002_002(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 制造一条未确认告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        # 记录确认前 Alarm Logs 中 Acknowledged 记录数（日志跨轮次累积，需差值比对）
        ha.goto_global_alarm(page, "Alarm Logs")
        acked_before = ha.count_rows_by_ack(page, _LABEL, acknowledged=True)

        # ── Step 2: Unacknowledged Alarms 页确认该告警 ──
        ha.ack_alarm(page, _LABEL)

        # ── Step 3: 蜂鸣器停止发声——物理蜂鸣需人工确认，脚本不做断言 ──

        # 确认后该条目应从 Unacknowledged Alarms 消失
        ha.goto_global_alarm(page, "Unacknowledged Alarms")
        assert ha.data_rows(page, _LABEL).count() == 0, \
            f"确认后 {_LABEL!r} 不应再出现在 Unacknowledged Alarms 列表中"

        # ── Step 4: Alarm Log 新增确认日志，Ack Status 为 Acknowledged ──
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True,
                                      min_count=acked_before + 1), (
            f"确认后全局 Alarm Logs 中 {_LABEL!r} 的 Acknowledged 记录数应增加"
            f"（确认前 {acked_before} 条）"
        )
    finally:
        ha.cleanup_test_rules(page)
