from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_AcuHMI_002_02_case08
# 用例标题：Reset 搜索条件，搜索条件被清除
# 预置条件：
#   1. AcuHMI 上电
#   2. 已接入 1 个设备并在线
#   3. Alarms 栏有至少 2 条告警显示
# 测试步骤：
#   1. 通过 Serial Number 检索告警
#   2. 检索是否告警成功，显示目标告警准确
#   3. 点击 "Reset" 重置搜索条件
#   4. 检查搜索框中搜索条件是否被清除
# 预期结果：
#   2. 检索告警成功，显示目标告警准确
#   4. 搜索框中搜索条件被清除


def test_TestCase_AcuHMI_002_02_case08(app_page: Page):
    page = app_page
    data = ha.ensure_alarm_log_data(page)

    # ── Step 1-2: 按 Serial Number 检索成功 ──
    ha.select_serial_filter(page, data["serial"])
    ha.click_search(page)
    assert ha.data_rows(page, data["label"]).count() > 0, \
        f"按 Serial Number 检索应包含目标告警 {data['label']!r}"
    for sn in ha.column_values(page, "Serial Number"):
        assert sn == data["serial"], \
            f"检索结果混入其他设备记录，Serial Number={sn!r}"

    # ── Step 3: 点击 Reset ──
    ha.click_reset(page)

    # ── Step 4: 各搜索框条件被清除 ──
    select_text = page.locator(".el-select:visible").first.inner_text().strip()
    assert data["serial"] not in select_text, (
        f"Reset 后 Serial Number 下拉应恢复占位提示，实际仍显示: {select_text!r}"
    )
    assert page.locator("input.el-range-input").nth(0).input_value() == "", \
        "Reset 后 Start Date 应被清空"
    assert page.locator("input.el-range-input").nth(1).input_value() == "", \
        "Reset 后 End Date 应被清空"
    assert page.locator("[placeholder='Enter Monitor ID']").input_value() == "", \
        "Reset 后 Monitor ID 应被清空"
