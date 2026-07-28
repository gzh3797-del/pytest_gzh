from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_004_004
# 用例标题：无未确认告警时 Unacknowledged Alarms 页面显示空列表
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 确保当前无未确认告警（全部已确认或无告警）
#   2. 进入 Unacknowledged Alarms 页面
# 预期结果：
#   2. 页面展示空列表，无异常报错，不显示已确认告警


def test_TestCase_ACUHMI17_BZ_004_004(app_page: Page):
    page = app_page

    # ── Step 1: 清空未确认告警（一键确认全部，列表已空则不操作）──
    ha.set_ack_enable(page, True)
    ha.ack_all_alarms(page)

    # ── Step 2: 页面为空列表、无报错、不显示已确认告警 ──
    ha.goto_global_alarm(page, "Unacknowledged Alarms")
    assert ha.data_rows(page).count() == 0, \
        "无未确认告警时 Unacknowledged Alarms 应显示空列表（不显示已确认告警）"
    assert page.locator(".el-message--error").count() == 0, \
        "空列表页面不应出现错误提示"
    assert ha.unack_total(page) == 0, "未确认告警总数应为 0"
