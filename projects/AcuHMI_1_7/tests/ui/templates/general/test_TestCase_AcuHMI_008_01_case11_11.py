import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case11_11
# 用例标题：Function Code下拉选择不同的code，Start Address输入不同的值，保存成功
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 New Typical Energy Meter Template，填写 Device 区
#   2. Block 选 READ_COILS + Decimal + Start=20 → Save Block → 应成功
#   3. Block 选 READ_DISCRETE_INPUTS + Decimal + Start=10 → Save Block → 应成功
#   4. Block 选 READ_HOLDING_REGISTERS + Decimal + Start=100 → Save Block → 应成功
#   5. Block 选 READ_INPUT_REGISTERS + Decimal + Start=200 → Save Block → 应成功
# 预期结果：
#   四种功能码配合不同 Start 地址均保存成功，无错误提示


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


def _select_function(page, func_name: str):
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, func_name)


def _select_address_format(page, fmt: str):
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    page.locator(".el-form-item").filter(has_text="Address Format").first.locator(".el-select").click()
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


def test_TestCase_AcuHMI_008_01_case11_11(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # 进入创建页，填写 Device 区
    _nav_to_templates(page)
    page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(f"FnTest_{ts}")
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

    # 四种 Function Code + Decimal + 不同 Start 值，逐一验证保存成功
    test_cases = [
        ("READ_COILS",              "Decimal", "20",  "1"),
        ("READ_DISCRETE_INPUTS",    "Decimal", "10",  "5"),
        ("READ_HOLDING_REGISTERS",  "Decimal", "100", "10"),
        ("READ_INPUT_REGISTERS",    "Decimal", "200", "2"),
    ]

    for func, fmt, start, count in test_cases:
        _select_function(page, func)
        _select_address_format(page, fmt)

        page.keyboard.press("Escape")
        page.wait_for_timeout(150)
        page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill(start)
        page.wait_for_timeout(200)
        page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill(count)
        page.wait_for_timeout(200)

        page.get_by_role("button", name="Save Block").click()
        page.wait_for_timeout(1500)

        assert not _has_error(page), \
            f"Function={func} + {fmt} + Start={start} + Count={count} 应保存成功，但出现了错误提示"
