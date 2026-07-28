from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_AcuHMI_002_02_case07
# 用例标题：清除告警功能正常，再次通过搜索条件仍可以搜到对应的告警
# 预置条件：
#   1. AcuHMI 上电
#   2. 已接入 1 个设备并在线
#   3. Alarms 栏有至少 2 条告警显示
# 测试步骤：
#   1. Alarm log 栏，点击 "Clear Logs"
#   2. Alarm log 栏中显示告警信息是否立即消失
# 预期结果：
#   2. Alarm log 栏中显示告警信息立即消失
# 说明：Clear Logs 为破坏性操作（清空全部历史告警日志），本用例放在
#      002_02 组最后执行；清空后其他检索用例的 ensure 前置会自动重建数据。


def test_TestCase_AcuHMI_002_02_case07(app_page: Page):
    page = app_page
    # 前置：Alarm Logs 中有告警记录（不足则触发补齐）
    ha.ensure_alarm_log_data(page)
    assert ha.data_rows(page).count() > 0

    # ── Step 1: 点击 Clear Logs（有二次确认弹窗）──
    page.get_by_role("button", name="Clear Logs").click()
    page.wait_for_timeout(400)
    ha._confirm_dialog(page)

    # ── Step 2: 告警信息立即消失 ──
    page.wait_for_timeout(1000)
    assert ha.data_rows(page).count() == 0, \
        "点击 Clear Logs 后 Alarm Logs 列表应立即清空"
    assert page.locator(".el-message--error").count() == 0, \
        "清除日志后页面不应出现错误提示"
