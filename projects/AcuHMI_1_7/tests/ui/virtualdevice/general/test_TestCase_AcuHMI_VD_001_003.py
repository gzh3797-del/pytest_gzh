# 用例编号: AcuHMI_VD_001_003
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


def _goto_first_page(page):
    """若当前不在第 1 页，回到第 1 页（列表默认每页 10 行，翻页后需归位再查找）。"""
    first_li = page.locator(".el-pagination .el-pager li").first
    if first_li.count() > 0 and "is-active" not in (first_li.get_attribute("class") or ""):
        first_li.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(400)


def _confirm_delete(page):
    for btn in ("Yes", "Confirm"):
        try:
            page.get_by_role("button", name=btn).click(timeout=2000)
            page.wait_for_timeout(500)
            return
        except Exception:
            pass


def _delete_virtual_device(page, name: str):
    """删除列表中所有同名虚拟设备（跨分页翻页查找），删不到即返回。"""
    for _ in range(40):
        _nav_to_virtual_devices(page)
        _goto_first_page(page)
        deleted = False
        while True:
            row = page.locator("tbody").get_by_role("row").filter(has_text=name)
            if row.count() > 0:
                row.first.locator(".el-button").last.click(force=True)
                page.wait_for_timeout(500)
                _confirm_delete(page)
                deleted = True
                break
            next_btn = page.locator(".el-pagination .btn-next")
            if next_btn.count() == 0 or next_btn.get_attribute("aria-disabled") == "true":
                break
            next_btn.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        if not deleted:
            return


def _vd_in_list(page, name: str) -> bool:
    """跨分页查找虚拟设备是否存在。"""
    _nav_to_virtual_devices(page)
    _goto_first_page(page)
    while True:
        if page.locator("tbody").get_by_role("row").filter(has_text=name).count() > 0:
            return True
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0 or next_btn.get_attribute("aria-disabled") == "true":
            return False
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


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


def test_TestCase_AcuHMI_VD_001_003(login_page: LoginPage):
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
        assert _vd_in_list(page, vd_interval5), "边界值 5 的 Virtual Device 保存后未在列表中显示"

        _add_vd_with_param(page, vd_interval15, interval="15")
        assert _vd_in_list(page, vd_interval15), "边界值 15 的 Virtual Device 保存后未在列表中显示"
    finally:
        _delete_virtual_device(page, vd_interval5)
        _delete_virtual_device(page, vd_interval15)
