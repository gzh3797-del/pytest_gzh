from datetime import datetime

from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_AcuHMI_002_02_case06
# 用例标题：通过告警间隔 Interval+Monitor ID，检索对应告警正常
# 预置条件：
#   1. AcuHMI 上电
#   2. 已接入 1 个设备并在线
#   3. Alarms 栏有至少 2 条告警显示
# 测试步骤：
#   1. 通过 Interval+Monitor ID 组合检索告警
#   2. 检索是否告警成功，显示目标告警准确
# 预期结果：
#   2. 检索告警成功，显示目标告警准确


def test_TestCase_AcuHMI_002_02_case06(app_page: Page):
    page = app_page
    data = ha.ensure_alarm_log_data(page)
    try:
        # ── Step 1: Interval + Monitor ID 组合检索 ──
        ha.set_interval_filter(page, "today")
        ha.fill_monitor_id_filter(page, data["monitor_id"])
        ha.click_search(page)

        # ── Step 2: 检索成功且所有结果同时满足两个条件 ──
        assert ha.data_rows(page, data["label"]).count() > 0, \
            f"组合检索应包含目标告警 {data['label']!r}"
        today_str = datetime.now().strftime("%Y-%m-%d")
        for ts in ha.column_values(page, "Timestamp"):
            assert ts.startswith(today_str), \
                f"结果中存在区间外记录，Timestamp={ts!r}"
        for mid in ha.column_values(page, "Monitor ID"):
            assert mid == data["monitor_id"], \
                f"结果混入其他 Monitor ID 记录：{mid!r}"
    finally:
        ha.click_reset(page)
