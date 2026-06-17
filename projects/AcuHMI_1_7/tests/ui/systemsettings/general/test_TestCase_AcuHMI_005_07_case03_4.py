# 用例编号: TestCase_AcuHMI_005_07_case03_4
# 用例标题: 添加白名单，IP值为非法值，添加失败，系统提示错误信息准确
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. Add Allow List, IP Range=Yes, From=192.168.1 / From=192.168.1.q / From=192.168.1.@ → 添加失败
#   2. Add Allow List, IP Range=No, IP Address=192.168.1 / 192.168.1.q / 192.168.1.@ → 添加失败
# 预期结果: 所有非法IP值添加失败，显示错误信息

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


def test_TestCase_AcuHMI_005_07_case03_4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Access Control")

    # Enable IP Allow List first so the Add Allow List button appears
    page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    invalid_ips = ["192.168.1", "192.168.1.q", "192.168.1.@"]

    try:
        # 测试IP Range=No时的非法IP (single IP, placeholder="Enter IP Address")
        for ip in invalid_ips:
            page.get_by_role("button", name="Add Allow List").click()
            page.wait_for_timeout(500)
            # Switch to No (single IP mode)
            page.locator(".el-dialog").locator(".el-radio").filter(has_text="No").click()
            page.wait_for_timeout(300)
            page.get_by_placeholder("Enter IP Address").fill(ip)
            page.get_by_role("button", name="Confirm").click()
            # EP 表单校验错误（.el-form-item__error）在 Confirm 后异步渲染，需 auto-wait，
            # 否则即时 .count() 取到 0 而误判（见兄弟用例 case03_5）。
            expect(page.locator(".el-form-item__error").first).to_be_visible(timeout=5000)
            try:
                page.get_by_role("button", name="Cancel").click(timeout=1000)
            except Exception:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
            page.wait_for_timeout(300)

        # 测试IP Range=Yes时的非法IP (range mode, placeholder="Enter From Address")
        for ip in invalid_ips:
            page.get_by_role("button", name="Add Allow List").click()
            page.wait_for_timeout(500)
            # Default is Yes (IP range mode) — "Enter From Address" field
            page.get_by_placeholder("Enter From Address").fill(ip)
            page.get_by_role("button", name="Confirm").click()
            expect(page.locator(".el-form-item__error").first).to_be_visible(timeout=5000)
            try:
                page.get_by_role("button", name="Cancel").click(timeout=1000)
            except Exception:
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
            page.wait_for_timeout(300)
    finally:
        # Disable IP Allow List to restore state
        try:
            page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
                ".el-radio"
            ).filter(has_text="Disable").click()
            page.wait_for_timeout(300)
            page.get_by_role("button", name="Save").click()
            page.wait_for_timeout(500)
        except Exception:
            pass
