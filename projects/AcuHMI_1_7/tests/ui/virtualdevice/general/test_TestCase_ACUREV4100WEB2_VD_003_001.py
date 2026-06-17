# 用例编号: ACUREV4100WEB2_VD_003_001
# 用例标题: Parameter Name 与 Post Label 合法配置保存展示正确
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 添加 Virtual Device，新增 Parameter
#   2. 输入合法 Parameter Name（≤40字符，字母/数字/下划线）
#   3. 输入合法 Post Label（≤40字符，字母/数字/下划线/空格）
#   4. 保存
# 预期结果: Parameter Name 与 Post Label 均保存成功，页面正确展示

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


def _delete_virtual_device(page, name: str):
    _nav_to_virtual_devices(page)
    row = page.locator("tbody").get_by_role("row").filter(has_text=name)
    if row.count() == 0:
        return
    row.locator(".el-button").last.click(force=True)
    page.wait_for_timeout(500)
    try:
        page.get_by_role("button", name="Yes").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.get_by_role("button", name="Confirm").click(timeout=2000)
            page.wait_for_timeout(500)
        except Exception:
            pass


def test_TestCase_ACUREV4100WEB2_VD_003_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_name = "VD_Param001"
    param_name = "param_name_1"
    post_label = "PostLabel1"
    unit = "kW"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)

        # Form already has one parameter row — fill it directly (no Add Parameter click)
        page.get_by_placeholder("---Enter Parameter Name---").first.fill(param_name)
        page.get_by_placeholder("---Enter Post Label---").first.fill(post_label)
        page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").first.fill(unit)

        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        _nav_to_virtual_devices(page)
        row = page.locator("tbody").get_by_role("row").filter(has_text=vd_name)
        assert row.count() > 0, f"Virtual Device '{vd_name}' 保存后未在列表中显示"

        row.locator("td").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.locator(".el-tabs__item").filter(has_text="Configuration").click()
        page.wait_for_timeout(500)

        expect(page.get_by_placeholder("---Enter Parameter Name---").first).to_have_value(param_name)
    finally:
        _delete_virtual_device(page, vd_name)
