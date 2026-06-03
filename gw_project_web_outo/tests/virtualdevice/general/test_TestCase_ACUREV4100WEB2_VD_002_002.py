# 用例编号: ACUREV4100WEB2_VD_002_002
# 用例标题: Virtual Device Name 非法输入阻止保存
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 分别输入：空值、超过40字符、含空格的名称、含特殊符号（@）的名称
#   2. 各次点击保存
# 预期结果: 每次均阻止保存，出现错误提示，不生成新设备

import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_virtual_devices(page):
    # Must be on the list page exactly (not on addVirtualMeter sub-path)
    on_list = "/#/virtualMeter" in page.url and "addVirtualMeter" not in page.url
    if not on_list:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Virtual Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_ACUREV4100WEB2_VD_002_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    invalid_names = [
        ("", "空值名称"),
        ("A" * 41, "超过40字符的名称"),
        ("Name With Space", "含空格的名称"),
        ("Name@Invalid", "含特殊符号@的名称"),
    ]

    for invalid_name, description in invalid_names:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        name_input = page.get_by_label("Virtual Device Name", exact=True)
        name_input.fill(invalid_name)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)

        assert page.locator(".el-form-item__error").count() > 0, \
            f"{description} 应显示错误提示但未出现"

        row = page.locator("tbody").get_by_role("row").filter(has_text=invalid_name)
        assert row.count() == 0, f"{description} 不应创建新设备，但列表中出现了对应条目"
