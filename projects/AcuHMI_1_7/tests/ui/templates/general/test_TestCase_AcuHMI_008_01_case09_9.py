import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case09_9
# 用例标题：Block输入框测试：READ_COILS+Decimal Start=20保存成功，Hex Start=abcdef或特殊字符保存失败
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 New Typical Energy Meter Template，填写 Device 区
#   2. Block 选 READ_COILS + Decimal + Start=20 → Save Block → 应成功
#   3. Block 选 READ_COILS + Hex + Start=abcdef → Save Block → 应报错
#   4. Block 选 READ_COILS + Hex + Start=@#$% → Save Block → 应报错
# 预期结果：
#   步骤2保存成功；步骤3、4出现验证错误提示


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


def _click_visible_option(page, option_text: str):
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


def _select_function(page, func_name: str):
    """Select Function from dropdown by name."""
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, func_name)


def _select_address_format(page, fmt: str):
    """Select Address Format (Hex / Decimal) from dropdown."""
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    addr_fi = page.locator(".el-form-item").filter(has_text="Address Format").first
    addr_fi.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, fmt)


def _has_error(page) -> bool:
    if page.locator(".el-message--error").count() > 0:
        return True
    for el in page.locator(".el-form-item__error").all():
        try:
            if el.is_visible() and el.inner_text().strip():
                return True
        except Exception:
            pass
    return False


def test_TestCase_AcuHMI_008_01_case09_9(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 进入创建页，填写 Device 区（必填字段）
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(f"BlkTest_{ts}")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill("v1.00")
    page.wait_for_timeout(200)

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)

    # ── Test 1: READ_COILS + Decimal + Start=20 → Save Block 应成功 ────────
    _select_function(page, "READ_COILS")
    _select_address_format(page, "Decimal")

    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("20")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    assert not _has_error(page), \
        "READ_COILS + Decimal + Start=20 应保存成功，但出现了错误提示"

    # ── Test 2: READ_COILS + Hex + Start=abcdef → Save Block 应报错 ────────
    _select_function(page, "READ_COILS")
    _select_address_format(page, "Hex")

    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("abcdef")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    assert _has_error(page), \
        "READ_COILS + Hex + Start=abcdef 超出有效范围，应保存失败，但未检测到错误提示"

    # ── Test 3: READ_COILS + Hex + Start=@#$% → Save Block 应报错 ─────────
    _select_function(page, "READ_COILS")
    _select_address_format(page, "Hex")

    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("@#$%")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    assert _has_error(page), \
        "READ_COILS + Hex + Start=@#$% 含非法字符，应保存失败，但未检测到错误提示"
