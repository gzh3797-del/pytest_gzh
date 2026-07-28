# 用例编号: AcuHMI_VD_002_001
# 用例标题: Serial Number 自动生成格式正确且不重复
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 添加 3 个 Virtual Device，记录各自 Serial Number
#   2. 对比各 Serial Number 格式与唯一性
# 预期结果: 每个 Serial Number 均为 AEVM+5 位数字格式；所有 Serial Number 互不重复

import re
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


def _get_serial_number(page, name: str) -> str:
    """跨分页翻页查找指定虚拟设备行，返回其 AEVM+5位 Serial Number。"""
    _nav_to_virtual_devices(page)
    _goto_first_page(page)
    while True:
        row = page.locator("tbody").get_by_role("row").filter(has_text=name)
        if row.count() > 0:
            cells = row.first.locator("td")
            for i in range(cells.count()):
                cell_text = cells.nth(i).inner_text().strip()
                if re.match(r"^AEVM\d{5}$", cell_text):
                    return cell_text
            return ""
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0 or next_btn.get_attribute("aria-disabled") == "true":
            return ""
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_AcuHMI_VD_002_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_names = ["VD_SN001", "VD_SN002", "VD_SN003"]

    try:
        for name in vd_names:
            _add_virtual_device(page, name)

        serial_numbers = []
        for name in vd_names:
            sn = _get_serial_number(page, name)
            assert re.match(r"^AEVM\d{5}$", sn), \
                f"{name} 的 Serial Number '{sn}' 格式不符合 AEVM+5位数字"
            serial_numbers.append(sn)

        assert len(set(serial_numbers)) == len(serial_numbers), \
            f"Serial Number 存在重复: {serial_numbers}"
    finally:
        for name in vd_names:
            _delete_virtual_device(page, name)
