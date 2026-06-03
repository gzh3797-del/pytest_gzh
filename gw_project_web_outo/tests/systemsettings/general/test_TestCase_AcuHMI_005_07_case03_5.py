# 用例编号: TestCase_AcuHMI_005_07_case03_5
# 用例标题: 白名单描述验证，描述≤40字符保存成功
# 预置条件: 1、管理权限登录AcuHMI网页
# 测试步骤:
#   1. 添加白名单，描述为40字符混合字符串 → 保存成功
#   2. 添加白名单，描述为41字符混合字符串 → 验证结果
# 预期结果: 40字符描述的白名单配置保存成功

import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def test_TestCase_AcuHMI_005_07_case03_5(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Access Control")

    # Enable IP Allow List first so the Add Allow List button appears
    page.locator(".el-form-item").filter(has_text="IP Allow List Enable").locator(
        ".el-radio"
    ).filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    # 构造40字符的混合描述（原字符串为39字符，末尾补一个字符）
    desc_40 = "qweRTYUIOP0123456789_ 23466789!@#RTYUIOP"
    assert len(desc_40) == 40, f"desc_40长度应为40，实际为{len(desc_40)}"

    try:
        # 添加白名单，描述为40字符，保存应成功
        page.get_by_role("button", name="Add Allow List").click()
        page.wait_for_timeout(500)
        # Switch to No (single IP mode) — placeholder becomes "Enter IP Address"
        page.locator(".el-dialog").locator(".el-radio").filter(has_text="No").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("Enter IP Address").fill("192.168.1.200")
        page.get_by_placeholder("Enter Description").fill(desc_40)
        page.get_by_role("button", name="Confirm").click()
        page.wait_for_timeout(500)
        assert page.locator(".el-form-item__error").count() == 0, \
            "40字符描述的白名单应配置保存成功"

        # 清理：删除测试创建的白名单条目（el-popconfirm: Yes button）
        try:
            row = page.locator("tbody").get_by_role("row").filter(has_text="192.168.1.200")
            if row.count() > 0:
                row.locator(".el-button").last.click(force=True)
                page.wait_for_timeout(500)
                page.locator(".el-popconfirm__action .el-button--primary").click(timeout=2000)
                page.wait_for_timeout(500)
        except Exception:
            pass
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
