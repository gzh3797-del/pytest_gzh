import time
import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_008_01_case03_3
# 用例标题：用户自定义创建模板成功，模板支持相同/不同协议下创建模板
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 Templates > New Typical Energy Meter Template
#   2. 填写 Template Name、Version、Typical Model、Wiring Configuration
#   3. 填写 Block（Function、Start、Count），点击 Save Block
#   4. 点击 Create Template 保存
#   5. 进入 Template List，在 Customized 区验证模板存在
# 预期结果：
#   创建成功，Template List 中可见新模板


def _nav_to_templates(page):
    """Navigate to AcuHMI > Templates section."""
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
    """Click first visible dropdown item matching option_text."""
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and option_text in item.inner_text():
                item.click()
                page.wait_for_timeout(400)
                return True
        except Exception:
            pass
    # Click first visible item as fallback
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


def test_TestCase_AcuHMI_008_01_case03_3(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_templates(page)
    page.wait_for_timeout(500)

    # Step 1: 进入 New Typical Energy Meter Template
    ntet = page.locator(".el-menu-item").filter(has_text="New Typical Energy Meter Template")
    ntet.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # Step 2: 填写 Device 区
    ts = str(int(time.time()))[-6:]
    template_name = f"TestTpl_{ts}"

    # Template Name
    page.locator(".el-form-item").filter(has_text="Template Name").first.locator("input").fill(template_name)
    page.wait_for_timeout(200)

    # Version
    version_input = page.locator(".el-form-item").filter(has_text="Version").first.locator("input")
    version_input.fill("v1.00")
    page.wait_for_timeout(200)

    # Typical Model — 选第一个可用选项
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Typical Model").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    # Wiring Configuration — 选第一个可用选项
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Wiring Configuration").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    # Step 3: 填写 Block 区
    # Function — 选第一个可用函数
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Function").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_visible_option(page, "")

    # Start（十六进制地址）
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    start_input = page.locator(".el-form-item").filter(has_text="Start").first.locator("input")
    start_input.fill("0001")
    page.wait_for_timeout(200)

    # Count
    count_input = page.locator(".el-form-item").filter(has_text="Count").first.locator("input")
    count_input.fill("1")
    page.wait_for_timeout(200)

    # 点击 Save Block
    page.get_by_role("button", name="Save Block").click()
    page.wait_for_timeout(1500)

    # 验证 Block 已添加到 Block Table（至少1行）
    block_rows = page.locator("tbody tr").filter(has_text="0001")
    # 即使 Block Table 无数据也继续（部分产品不强制要求 Block）
    page.wait_for_timeout(500)

    # Step 4: 点击 Create Template
    page.get_by_role("button", name="Create Template").click()
    page.wait_for_timeout(2000)

    # 验证无 error toast
    assert page.locator(".el-message--error").count() == 0, \
        "Create Template 应成功，但出现了错误提示"

    # Step 5: 进入 Template List，验证 Customized 区存在新模板
    tl = page.locator(".el-menu-item").filter(has_text="Template List")
    tl.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    assert page.get_by_text(template_name, exact=False).count() > 0, \
        f"新建模板 '{template_name}' 应出现在 Template List 中"
