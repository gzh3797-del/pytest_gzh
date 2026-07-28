from playwright.sync_api import Page, expect

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_001_002
# 用例标题：Alarm 页面包含 Active Alarms 和 Alarm Logs 子页面
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 点击 Alarm 导航入口进入 Alarm 页面
#   2. 查看页面内子页面结构
# 预期结果：
#   2. 页面内正确展示 Active Alarms（实装名 Unacknowledged Alarms）和
#      Alarm Logs 两个子页面，均可正常跳转访问


def test_TestCase_ACUHMI17_BZ_001_002(app_page: Page):
    page = app_page

    # ── Step 1: 点击 Alarm 导航入口 ──
    ha.goto_global_alarm(page)

    # ── Step 2: 两个子页面入口均展示且可跳转 ──
    active_tab = page.get_by_role("menuitem", name="Unacknowledged Alarms")
    logs_tab = page.get_by_role("menuitem", name="Alarm Logs")
    expect(active_tab.first).to_be_visible()
    expect(logs_tab.first).to_be_visible()

    # 跳转 Unacknowledged Alarms（Active Alarms）
    active_tab.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    assert "activeAlarm" in page.url, \
        f"点击 Unacknowledged Alarms 后应进入 activeAlarm 路由，实际: {page.url}"
    assert page.locator(".el-message--error").count() == 0, \
        "Unacknowledged Alarms 页面不应出现错误提示"
    expect(page.locator(".el-table").first).to_be_visible()

    # 跳转 Alarm Logs
    logs_tab.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    assert "alarmLog" in page.url, \
        f"点击 Alarm Logs 后应进入 alarmLogs 路由，实际: {page.url}"
    assert page.locator(".el-message--error").count() == 0, \
        "Alarm Logs 页面不应出现错误提示"
    expect(page.locator(".el-table").first).to_be_visible()
