# 用例编号: ACUREV4100WEB2_VD_002_001
# 用例标题: Virtual Device Name 合法字符与长度范围验证
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 点击 Add Virtual Device
#   2. 输入仅含数字/字母/下划线、长度为1~40字符的名称 "VD_Test001"
#   3. 点击 Save
#   4. 返回列表查看
# 预期结果: 名称保存成功，列表正确显示 "VD_Test001"

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


def test_TestCase_ACUREV4100WEB2_VD_002_001(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_name = "VD_Test001"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)
        # Form already has one parameter row — fill it directly (no Add Parameter click)
        page.get_by_placeholder("---Enter Parameter Name---").first.fill("param_1")
        page.get_by_placeholder("---Enter Post Label---").first.fill("param_1")
        page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").first.fill("kW")
        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        _nav_to_virtual_devices(page)
        row = page.locator("tbody").get_by_role("row").filter(has_text=vd_name)
        assert row.count() > 0, "合法名称 VD_Test001 保存后未在列表中显示"
    finally:
        _delete_virtual_device(page, vd_name)
