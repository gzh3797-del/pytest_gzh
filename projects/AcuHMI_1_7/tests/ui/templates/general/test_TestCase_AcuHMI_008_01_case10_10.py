import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case10_10
# 用例标题：切换不同接线方式查看模板需要显示的参数是否一致
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 New Typical Energy Meter Template，填写 Template Name、Version、Typical Model
#   2. 依次切换每种 Wiring Configuration
#   3. 每次切换后验证 Block 区的参数字段（Function/Address Format/Start/Count/Save Block）始终存在
# 预期结果：
#   切换不同接线方式后，Block 区必要参数字段均一致显示，表单无异常

WIRING_OPTIONS = [
    "3 Element 4 Wire Y",
    "1 Element 2 Wire",
    "2 Element 3 Wire 1 Phase",
    "2 Element 3 Wire Network",
    "2 Element 3 Wire Delta",
    "3 Element 3 Wire Delta",
    "3 Element 4 Wire Delta",
    "2 1/2 Element 4 Wire Y",
]


def _nav_to_templates(page):
    if "/#/templates" not in page.url:
        if not any(s in page.url for s in [
            "/#/systemSettings", "/#/templates", "/#/protocols",
            "/#/maintenance", "/#/diagnostics", "/#/userManagement", "/#/firmwareUpdate",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Templates").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_visible_option(page, option_text: str = ""):
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    for item in all_items:
        try:
            if item.is_visible():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


def _select_wiring(page, wiring_name: str):
    """Switch Wiring Configuration to the given option."""
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, wiring_name)


def _block_fields_present(page) -> list:
    """Return list of missing required Block fields."""
    missing = []
    # Function
    if page.locator(".el-form-item").filter(has_text="Function").count() == 0:
        missing.append("Function")
    # Address Format
    if page.locator(".el-form-item").filter(has_text="Address Format").count() == 0:
        missing.append("Address Format")
    # Start
    if page.locator(".el-form-item").filter(has_text="Start").count() == 0:
        missing.append("Start")
    # Count
    if page.locator(".el-form-item").filter(has_text="Count").count() == 0:
        missing.append("Count")
    # Save Block button
    if page.get_by_role("button", name="Save Block").count() == 0:
        missing.append("Save Block button")
    return missing


def test_TestCase_AcuHMI_008_01_case10_10(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入创建页，填写 Device 区
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(f"WirTest_{ts}")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")
    page.wait_for_timeout(200)

    # Typical Model
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    # Step 2: 依次切换每种 Wiring Configuration，验证 Block 区字段一致
    for wiring in WIRING_OPTIONS:
        _select_wiring(page, wiring)

        # Step 3: 验证 Block 区必要字段均存在
        missing = _block_fields_present(page)
        assert not missing, \
            f"Wiring Configuration='{wiring}' 时，Block 区以下字段缺失：{missing}"
