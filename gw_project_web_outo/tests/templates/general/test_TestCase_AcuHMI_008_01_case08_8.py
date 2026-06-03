import time
import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case08_8
# 用例标题：Template Name和Version数字名称有效，Name超40字符Version只能字母/数字
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. Template Name 填纯数字，其余字段填有效值 → 创建应成功
#   2. Template Name 超过 40 字符，其余字段填不同有效值 → 应报验证错误
#   3. Version 填特殊字符，其余字段填不同有效值 → 应报验证错误
# 预期结果：
#   纯数字名称创建成功；超长名称和非法 Version 给出验证错误提示


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


def _click_nth_visible_option(page, nth: int = 0):
    """Click the nth visible dropdown item (0-based)."""
    all_items = page.locator(".el-select-dropdown__item").all()
    visible = []
    for item in all_items:
        try:
            if item.is_visible():
                visible.append(item)
        except Exception:
            pass
    if not visible:
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)
        return False
    target = visible[min(nth, len(visible) - 1)]
    target.click()
    page.wait_for_timeout(400)
    return True


def _open_create_page(page):
    _nav_to_templates(page)
    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def _fill_and_save_block(page, template_name: str, version: str,
                          func_nth: int, addr_fmt_nth: int,
                          start: str, count: str):
    """Fill full form with given values and click Save Block."""
    # Template Name
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(template_name)
    page.wait_for_timeout(200)

    # Version
    page.locator(".el-form-item").filter(has_text="Version").first.locator("input").fill(version)
    page.wait_for_timeout(200)

    # Typical Model (first option, same for all — not under test here)
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_nth_visible_option(page, 0)

    # Wiring Configuration (first option)
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_nth_visible_option(page, 0)

    # Function (by nth)
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_nth_visible_option(page, func_nth)

    # Address Format (by nth)
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    addr_fi = page.locator(".el-form-item").filter(has_text="Address Format").first
    if addr_fi.count() > 0:
        addr_fi.locator(".el-select").click()
        page.wait_for_timeout(400)
        _click_nth_visible_option(page, addr_fmt_nth)

    # Start
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill(start)
    page.wait_for_timeout(200)

    # Count
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill(count)
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)


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


def test_TestCase_AcuHMI_008_01_case08_8(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    ts = str(int(time.time()))[-6:]

    # ── Test 1: Template Name 纯数字 → 应合法，创建成功 ───────────────────
    # Version=v1.00, Function=第1个, Address Format=第1个, Start=0001, Count=10
    _open_create_page(page)
    _fill_and_save_block(
        page,
        template_name=f"{ts}001",      # 纯数字（含时间戳保证唯一）
        version="v1.00",
        func_nth=0,
        addr_fmt_nth=0,
        start="0001",
        count="10",
    )
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)
    assert not _has_error(page), \
        f"Template Name 为纯数字应合法，创建应成功，但出现了错误提示"

    # ── Test 2: Template Name 超过 40 字符 → 应报验证错误 ─────────────────
    # Version=v2.01, Function=第2个, Address Format=第1个, Start=00FF, Count=5
    _open_create_page(page)
    _fill_and_save_block(
        page,
        template_name="B" * 41,        # 41字符，超出40字符限制
        version="v2.01",
        func_nth=1,
        addr_fmt_nth=0,
        start="00FF",
        count="5",
    )
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)
    assert _has_error(page), \
        "Template Name 超过 40 字符时应显示验证错误，但未检测到任何错误提示"

    # ── Test 3: Version 含特殊字符 → 应报验证错误 ─────────────────────────
    # TemplateName=有效值, Version=@#$%!, Function=第1个, Address Format=第1个, Start=0100, Count=2
    _open_create_page(page)
    _fill_and_save_block(
        page,
        template_name=f"VerTest{ts}",
        version="@#$%!",               # 含非法特殊字符
        func_nth=0,
        addr_fmt_nth=0,
        start="0100",
        count="2",
    )
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)
    assert _has_error(page), \
        "Version 含特殊字符 '@#$%!' 时应显示验证错误，但未检测到任何错误提示"
