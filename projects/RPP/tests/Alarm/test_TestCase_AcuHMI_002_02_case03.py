from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_AcuHMI_002_02_case03
# 用例标题：通过告警间隔 Monitor ID，检索对应告警正常
# 预置条件：
#   1. AcuHMI 上电
#   2. 已接入 1 个设备并在线
#   3. Alarms 栏有至少 2 条告警显示
# 测试步骤：
#   1. 通过 Monitor ID 检索告警
#   2. 检索是否告警成功，显示目标告警准确
# 预期结果：
#   2. 检索告警成功，显示目标告警准确


def test_TestCase_AcuHMI_002_02_case03(app_page: Page):
    page = app_page
    data = ha.ensure_alarm_log_data(page)
    try:
        # ── Step 1: 按目标记录的 Monitor ID 检索 ──
        ha.fill_monitor_id_filter(page, data["monitor_id"])
        ha.click_search(page)

        # ── Step 2: 检索成功、结果只含该 Monitor ID 且含目标告警 ──
        assert ha.data_rows(page, data["label"]).count() > 0, \
            f"按 Monitor ID 检索应包含目标告警 {data['label']!r}"
        for mid in ha.column_values(page, "Monitor ID"):
            assert mid == data["monitor_id"], (
                f"检索结果混入其他 Monitor ID 的记录：{mid!r}，"
                f"期望均为 {data['monitor_id']!r}"
            )
    finally:
        ha.click_reset(page)
