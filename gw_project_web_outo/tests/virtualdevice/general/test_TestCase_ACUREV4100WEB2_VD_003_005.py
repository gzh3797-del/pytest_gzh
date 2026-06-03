# 用例编号: ACUREV4100WEB2_VD_003_005
# 用例标题: 单个 Virtual Device 支持最多 20 个 Parameter
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 新建一个 Virtual Device
#   2. 逐一添加 20 个 Parameter，每个均填写 Name
#   3. 保存并查看
#   4. 尝试添加第 21 个 Parameter，查看是否被阻止
# 预期结果: 所有 20 个 Parameter 均保存成功；第 21 个 Parameter 添加失败或按钮不可用

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


def test_TestCase_ACUREV4100WEB2_VD_003_005(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_name = "VD_MaxParam"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)

        # Form already has one row — fill it as param_1 (no click needed)
        page.get_by_placeholder("---Enter Parameter Name---").last.fill("param_1")
        page.get_by_placeholder("---Enter Post Label---").last.fill("param_1")
        page.get_by_placeholder("---Enter Calculated Meter Formula---").last.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").last.fill("kW")

        # Add and fill rows 2–20
        for i in range(2, 21):
            page.get_by_role("button", name="Add Parameter").click()
            page.wait_for_timeout(300)
            page.get_by_placeholder("---Enter Parameter Name---").last.fill(f"param_{i}")
            page.get_by_placeholder("---Enter Post Label---").last.fill(f"param_{i}")
            page.get_by_placeholder("---Enter Calculated Meter Formula---").last.fill("1+1")
            page.get_by_placeholder("---Enter Unit---").last.fill("kW")

        add_param_btn = page.get_by_role("button", name="Add Parameter")
        is_disabled = add_param_btn.is_disabled()
        if not is_disabled:
            add_param_btn.click()
            page.wait_for_timeout(500)
            error_visible = page.locator(".el-form-item__error").count() > 0 or \
                page.get_by_text("maximum", case_sensitive=False).count() > 0
            assert error_visible or page.get_by_placeholder("---Enter Parameter Name---").count() <= 20, \
                "第 21 个 Parameter 应被阻止添加，但未出现错误提示且输入行数量超过 20"

        page.get_by_role("button", name="Save").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        _nav_to_virtual_devices(page)
        row = page.locator("tbody").get_by_role("row").filter(has_text=vd_name)
        assert row.count() > 0, f"包含 20 个 Parameter 的 Virtual Device '{vd_name}' 保存后未在列表中显示"
    finally:
        _delete_virtual_device(page, vd_name)
