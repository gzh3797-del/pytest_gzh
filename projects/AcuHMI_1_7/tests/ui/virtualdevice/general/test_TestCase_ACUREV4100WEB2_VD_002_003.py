# 用例编号: ACUREV4100WEB2_VD_002_003
# 用例标题: Calculated Interval 默认值与边界值配置正确
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 新建 Virtual Device，查看 Calculated Interval 默认值
#   2. 输入边界值 5，保存后查看展示值
#   3. 再次新建，输入边界值 15，保存后查看展示值
# 预期结果: 默认值为 5；5 和 15 均保存成功，展示正确

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


def _add_vd_with_param(page, name: str, interval: str = None):
    _nav_to_virtual_devices(page)
    page.get_by_role("button", name="Add Virtual Device").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_label("Virtual Device Name", exact=True).fill(name)
    if interval is not None:
        page.get_by_placeholder("---Enter Calculated Interval---").fill(interval)
    # Form already has one parameter row — fill it directly (no Add Parameter click)
    page.get_by_placeholder("---Enter Parameter Name---").first.fill("param_1")
    page.get_by_placeholder("---Enter Post Label---").first.fill("param_1")
    page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
    page.get_by_placeholder("---Enter Unit---").first.fill("kW")
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def test_TestCase_ACUREV4100WEB2_VD_002_003(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_interval5 = "VD_Interval5"
    vd_interval15 = "VD_Interval15"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        interval_input = page.get_by_placeholder("---Enter Calculated Interval---")
        default_value = interval_input.input_value()
        assert default_value == "5", f"Calculated Interval 默认值应为 5，实际为 {default_value}"

        # Navigate back and create VD with interval=5 (default)
        _add_vd_with_param(page, vd_interval5)

        _nav_to_virtual_devices(page)
        row5 = page.locator("tbody").get_by_role("row").filter(has_text=vd_interval5)
        assert row5.count() > 0, "边界值 5 的 Virtual Device 保存后未在列表中显示"

        _add_vd_with_param(page, vd_interval15, interval="15")

        _nav_to_virtual_devices(page)
        row15 = page.locator("tbody").get_by_role("row").filter(has_text=vd_interval15)
        assert row15.count() > 0, "边界值 15 的 Virtual Device 保存后未在列表中显示"
    finally:
        _delete_virtual_device(page, vd_interval5)
        _delete_virtual_device(page, vd_interval15)
