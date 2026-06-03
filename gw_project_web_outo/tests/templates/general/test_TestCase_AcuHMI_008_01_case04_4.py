import time
import pytest
from playwright.sync_api import expect
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case04_4
# 用例标题：用户自定义创建模板成功，模板支持修改模板
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 New Typical Energy Meter Template，创建一个自定义模板
#   2. 进入 Template List，找到 Customized 区的新建模板行
#   3. 点击黄色（Edit）图标按钮进入编辑页面
#   4. 验证 Template Name 字段不可编辑
# 预期结果：
#   进入编辑页后，Template Name 字段处于不可编辑状态（禁用/只读）


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
    page.wait_for_timeout(300)
    return False


def _create_template(page) -> str:
    """Create a new custom template, return template name."""
    _nav_to_templates(page)
    page.wait_for_timeout(500)

    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    ts = str(int(time.time()))[-6:]
    template_name = f"TestTpl_{ts}"

    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(template_name)
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
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Start").first.locator("input").fill("0001")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Count").first.locator("input").fill("1")
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)

    assert page.locator(".el-message--error").count() == 0, "创建模板应成功"
    return template_name


def test_TestCase_AcuHMI_008_01_case04_4(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 创建自定义模板
    template_name = _create_template(page)

    # Step 2: 进入 Template List
    _nav_to_templates(page)
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    row = page.locator("tbody tr").filter(has_text=template_name)
    assert row.count() > 0, f"Template List 中未找到模板 '{template_name}'"

    # Step 3: 点击黄色（Edit）图标按钮
    row.first.locator(".el-button--warning").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 4: 验证 Template Name 不可编辑
    name_input = page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").first
    assert not name_input.is_editable(), \
        "Template Name 在编辑模式下应为不可编辑状态"
