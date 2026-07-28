# 用例编号: AcuHMI_VD_001_004
# 用例标题: Calculated Interval 超出范围或非整数阻止保存
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 分别输入 4（低于下限）、16（高于上限）、5.5（非整数）、空值
#   2. 各次点击保存
# 预期结果: 每次均阻止保存并提示错误

import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


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


def test_TestCase_AcuHMI_VD_001_004(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    invalid_intervals = [
        ("4", "低于下限值 4"),
        ("16", "高于上限值 16"),
        ("5.5", "非整数 5.5"),
        ("", "空值"),
    ]

    for interval_value, description in invalid_intervals:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill("VD_IntervalTest")
        interval_input = page.get_by_placeholder("---Enter Calculated Interval---")
        interval_input.fill(interval_value)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)

        # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
        page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
        assert page.locator(".el-form-item__error").count() > 0, \
            f"{description} 应显示错误提示但未出现"
