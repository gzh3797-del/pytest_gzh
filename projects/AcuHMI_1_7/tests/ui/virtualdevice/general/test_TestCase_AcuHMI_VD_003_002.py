# 用例编号: AcuHMI_VD_003_002
# 用例标题: Delete 删除单个 Parameter 项正常
# 预置条件: 已登录系统，进入 Virtual Devices 页面
# 测试步骤:
#   1. 配置 Virtual Device 含多个 Parameter
#   2. 点击某 Parameter 的 Delete 按钮
#   3. 保存后查看
# 预期结果: 被删除参数从列表移除，其余参数不受影响

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


def test_TestCase_AcuHMI_VD_003_002(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    vd_name = "VD_DelParam"
    param_keep = "param_to_keep"
    param_delete = "param_to_delete"

    try:
        _nav_to_virtual_devices(page)
        page.get_by_role("button", name="Add Virtual Device").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

        page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)

        # Form already has one parameter row — use it as param_keep
        page.get_by_placeholder("---Enter Parameter Name---").first.fill(param_keep)
        page.get_by_placeholder("---Enter Post Label---").first.fill(param_keep)
        page.get_by_placeholder("---Enter Calculated Meter Formula---").first.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").first.fill("kW")

        # Click Add Parameter once to add the row we will delete
        page.get_by_role("button", name="Add Parameter").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("---Enter Parameter Name---").last.fill(param_delete)
        page.get_by_placeholder("---Enter Post Label---").last.fill(param_delete)
        page.get_by_placeholder("---Enter Calculated Meter Formula---").last.fill("1+1")
        page.get_by_placeholder("---Enter Unit---").last.fill("kW")

        # Delete the second parameter row via the last danger button in the form
        page.locator(".el-button--danger").last.click()
        page.wait_for_timeout(500)

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

        param_inputs = page.get_by_placeholder("---Enter Parameter Name---")
        param_names_in_form = [param_inputs.nth(i).input_value() for i in range(param_inputs.count())]
        assert param_keep in param_names_in_form, \
            f"保留的参数 '{param_keep}' 在编辑页面未显示"
        assert param_delete not in param_names_in_form, \
            f"已删除的参数 '{param_delete}' 仍显示在编辑页面"
    finally:
        _delete_virtual_device(page, vd_name)
