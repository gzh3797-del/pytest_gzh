from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_005_003
# 用例标题：物理设备 Alarm Log 含 Ack Status 信息
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. Alarm Acknowledgement Enable 为开启
#   2. 进入 Devices 页面，打开某物理设备的 Alarm Log
#   3. 触发并确认一条告警，查看该物理设备 Alarm Log
# 预期结果：
#   3. 物理设备 Alarm Log 中包含 Ack Status 列，已确认告警显示 Acknowledged

_LABEL = "at_bz005003"


def test_TestCase_ACUHMI17_BZ_005_003(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 确认开关置为 Enable ──
        ha.set_ack_enable(page, True)

        # ── Step 3: 触发并确认一条告警 ──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"
        ha.ack_alarm(page, _LABEL)

        # ── Step 2+3: 物理设备详情页 Alarm Logs 含 Ack Status 列，
        #             已确认告警显示 Acknowledged ──
        ha.goto_device_alarm(page, "Alarm Logs")
        headers = ha.table_headers(page)
        assert "Ack Status" in headers, \
            f"物理设备 Alarm Logs 应包含 Ack Status 列，实际列: {headers}"

        rows = ha.data_rows(page, _LABEL)
        assert rows.count() > 0, f"物理设备 Alarm Logs 中应有 {_LABEL!r} 的记录"
        acked = any(
            ha.ack_status_of_row(rows.nth(i)).lower().startswith("acknowledged")
            for i in range(rows.count())
        )
        assert acked, (
            f"物理设备 Alarm Logs 中 {_LABEL!r} 应至少有一条 Acknowledged 记录"
        )
    finally:
        ha.cleanup_test_rules(page)
