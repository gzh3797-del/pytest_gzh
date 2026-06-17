# 用例编号: ACUREV4100WEB2_VD_004_001
# 用例标题: 删除 Virtual Device 后从列表移除，不影响其他设备
# 预置条件: 已登录系统，系统中已存在多个 Virtual Device
# 测试步骤:
#   1. 创建两个 Virtual Device：VD_Keep001 和 VD_ToDelete001
#   2. 点击 VD_ToDelete001 行的 Delete 按钮并确认删除
#   3. 查看列表
# 预期结果: VD_ToDelete001 从列表消失；VD_Keep001 仍然存在于列表中

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


def _add_virtual_device(page, name: str):
    _nav_to_virtual_devices(page)
    page.get_by_role("button", name="Add Virtual Device").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_label("Virtual Device Name", exact=True).fill(name)
    # Form already has one parameter row — fill it directly (no Add Parameter click)
    page.get_by_placeholder("---Enter Parameter Name---").first.fill("param_1")
    page.get_by_placeholder("---Enter Post Label---").first.fill("param_1")
    page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
    page.get_by_placeholder("---Enter Unit---").first.fill("kW")
    page.get_by_role("button", name="Save").click()
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


def test_TestCase_ACUREV4100WEB2_VD_004_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_keep = "VD_Keep001"
    vd_to_delete = "VD_ToDelete001"

    try:
        _add_virtual_device(page, vd_keep)
        _add_virtual_device(page, vd_to_delete)

        _delete_virtual_device(page, vd_to_delete)

        _nav_to_virtual_devices(page)
        deleted_row = page.locator("tbody").get_by_role("row").filter(has_text=vd_to_delete)
        assert deleted_row.count() == 0, \
            f"已删除的 Virtual Device '{vd_to_delete}' 仍显示在列表中"

        keep_row = page.locator("tbody").get_by_role("row").filter(has_text=vd_keep)
        assert keep_row.count() > 0, \
            f"应保留的 Virtual Device '{vd_keep}' 在删除其他设备后从列表消失"
    finally:
        _delete_virtual_device(page, vd_keep)
        _delete_virtual_device(page, vd_to_delete)
