from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_AcuHMI_002_02_case02
# 用例标题：通过告警间隔 Serial Number，检索对应告警正常
# 预置条件：
#   1. AcuHMI 上电
#   2. 已接入 1 个设备并在线
#   3. Alarms 栏有至少 2 条告警显示
# 测试步骤：
#   1. 通过 Serial Number 检索告警
#   2. 检索是否告警成功，显示目标告警准确
# 预期结果：
#   2. 检索告警成功，显示目标告警准确


def test_TestCase_AcuHMI_002_02_case02(app_page: Page):
    page = app_page
    data = ha.ensure_alarm_log_data(page)
    try:
        # ── Step 1: 按目标设备 Serial Number 检索 ──
        ha.select_serial_filter(page, data["serial"])
        ha.click_search(page)

        # ── Step 2: 检索成功、结果只含该 Serial Number 且含目标告警 ──
        assert ha.data_rows(page, data["label"]).count() > 0, \
            f"按 Serial Number 检索应包含目标告警 {data['label']!r}"
        for sn in ha.column_values(page, "Serial Number"):
            assert sn == data["serial"], (
                f"检索结果混入其他设备的记录，Serial Number={sn!r}，"
                f"期望均为 {data['serial']!r}"
            )
    finally:
        ha.click_reset(page)
