import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_06_case01
# 用例标题：Datalog Parameter Config界面布局和标题验证
# 预置条件：
#   1. AcuHMI上电启动正常
#   2. 下挂PXB PXE1,PXE2,PXM350设备
# 测试步骤：
#   1. 点击Data Log->Data Loggers下的Data Log Parameter Config页面
#   2. 查看标题，下拉框是否正常，单词拼写等是否正确
# 预期结果：
#   2. 页面标签布局正常，显示正常，单词拼写正常


def _nav_to_data_log_param_config(page):
    """Navigate to Data Log > Data Loggers > Data Log Parameter Config."""
    if "/#/dataLog" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Data Log").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    # 点击顶部 "Data Loggers" Tab（el-sub-menu__title）展开下拉
    data_loggers_tab = page.locator("div.el-sub-menu__title").filter(has_text="Data Loggers")
    if data_loggers_tab.count() > 0 and data_loggers_tab.first.is_visible():
        data_loggers_tab.first.click()
        page.wait_for_timeout(500)

    # 点击下拉中的 "Data Log Parameter Config"
    cfg_item = page.locator(".el-menu-item").filter(has_text="Data Log Parameter Config")
    if cfg_item.count() > 0 and cfg_item.first.is_visible():
        cfg_item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_AcuHMI_003_06_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 导航到 Data Log Parameter Config 页面
    _nav_to_data_log_param_config(page)
    page.wait_for_timeout(800)

    # Step 2: 验证页面标题、标签布局和拼写

    # 面包屑或页面标题包含 "Data Log Parameter Config"
    expect(page.get_by_text("Data Log Parameter Config", exact=False).first).to_be_visible()

    # Device 下拉框存在
    expect(
        page.get_by_text("Device", exact=False).first
    ).to_be_visible()

    # Parameter Type 标签存在
    expect(
        page.get_by_text("Parameter Type", exact=False).first
    ).to_be_visible()

    # Parameters 区域存在
    expect(
        page.get_by_text("Parameters", exact=False).first
    ).to_be_visible()

    # Not Selected / Selected 双列存在（拼写验证）
    expect(page.get_by_text("Not Selected", exact=False).first).to_be_visible()
    expect(page.get_by_text("Selected", exact=False).first).to_be_visible()

    # All / Clear 按钮存在（拼写验证）
    expect(page.get_by_text("All", exact=False).first).to_be_visible()
    expect(page.get_by_text("Clear", exact=False).first).to_be_visible()

    # 页面无错误提示
    assert page.locator(".el-message--error").count() == 0, \
        "页面初始加载不应出现错误提示"
