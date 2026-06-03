# 用例编号: ACUREV4100WEB2_VD_003_004
# 用例标题: Unit 字段最多40字符、无字符限制，保存展示正确
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 在 Unit 字段输入含特殊字符的字符串（如 kWh/m²-[]{}），长度不超过40
#   2. 保存后查看
# 预期结果: Unit 保存成功，无字符类型限制，展示与输入一致

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


def test_TestCase_ACUREV4100WEB2_VD_003_004(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_name = "VD_Unit001"
    special_unit = "kWh/m²-[]{}"

    assert len(special_unit) <= 40, "测试用 Unit 字符串超过40字符限制"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)

        # Form already has one parameter row — fill it directly (no Add Parameter click)
        page.get_by_placeholder("---Enter Parameter Name---").first.fill("param_unit_test")
        page.get_by_placeholder("---Enter Post Label---").first.fill("param_unit_test")
        page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").first.fill(special_unit)

        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        _nav_to_virtual_devices(page)
        row = page.locator("tbody").get_by_role("row").filter(has_text=vd_name)
        assert row.count() > 0, f"含特殊字符 Unit 的 Virtual Device '{vd_name}' 保存后未在列表中显示"
    finally:
        _delete_virtual_device(page, vd_name)
