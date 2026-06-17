# 用例编号: TestCase_AcuHMI_005_02_case02_02
# 用例标题: 以太网2 IP地址为非法值，DNS为非法值，保存配置失败，提示错误信息准确
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. IP=192.168.1.a, DNS=8.8.a.w → Save → 失败
#   2. IP=192.168.*.123, DNS=8.8.8.@ → Save → 失败
#   3. IP=192.168.1123, DNS=8.8.83 → Save → 失败
# 预期结果: 所有非法IP/DNS值保存失败，显示错误信息
# 注：以太网2仅有IP和DNS字段，无Mask/Gateway；DHCP Enable选项为Auto/Manual

import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def test_TestCase_AcuHMI_005_02_case02_02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Network")

    # Switch Ethernet 2 to Manual (static) mode so its IP field becomes editable
    page.locator(".el-form-item").filter(has_text="DHCP Enable").nth(1).locator(
        ".el-radio"
    ).filter(has_text="Manual").click()
    page.wait_for_timeout(500)

    # Ethernet 2 has IP and DNS fields only (no Mask/Gateway)
    # All saves with invalid values should fail with form validation errors
    invalid_combos = [
        ("192.168.1.a", "8.8.a.w"),
        ("192.168.*.123", "8.8.8.@"),
        ("192.168.1123", "8.8.83"),
    ]
    try:
        for ip, dns in invalid_combos:
            page.get_by_placeholder("Enter IP").nth(1).fill(ip)
            page.get_by_placeholder("Enter DNS 1").first.fill(dns)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)
            # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
            page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
            assert page.locator(".el-form-item__error").count() > 0, \
                f"以太网2非法IP ({ip!r}) 应保存失败并显示错误信息"
    finally:
        # Navigate away to discard all unsaved changes (no valid save was made)
        page.goto(BASE_URL + "/#/dashboard")
        page.wait_for_load_state("networkidle")
