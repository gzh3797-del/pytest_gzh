import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_03_case11
# 用例标题：Logger3 LogFileNamePrefix超过20字符保存失败
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. Logger3开关为enable，PostChannel选择Channel3，LogFileFormat选择Json，
#      LogFileLength选择10 minutes，LogFileNameFormat选择Time interval Format，
#      LogFileNamePrefix输入超过20字符的字符串（21字符），
#      Log Interval选择1 minute，保存配置
# 预期结果：
#   1. 保存配置失败，提示"Log File Name Prefix cannot exceed 20 characters"


def _nav_to_data_loggers3(page):
    """Navigate to Data Log > Data Loggers > Data Loggers 3."""
    if "/#/dataLog" not in page.url:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Data Log").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)

    tab = page.locator("div.el-sub-menu__title").filter(has_text="Data Loggers")
    if tab.count() > 0 and tab.first.is_visible():
        tab.first.click()
        page.wait_for_timeout(400)

    item = page.locator(".el-menu-item").filter(has_text="Data Loggers 3")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_option(page, option_text: str):
    """Click the visible dropdown item matching option_text, then ensure dropdown is closed."""
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and item.inner_text().strip() == option_text:
                item.click()
                page.wait_for_timeout(400)
                return
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)


def test_TestCase_AcuHMI_003_03_case11(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_loggers3(page)
    page.wait_for_timeout(800)

    # 开启 Logger 3
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(500)

    # Post Channel → Channel3
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Post Channel").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Channel3")

    # Log File Format → Json
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Format").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "Json")

    # Log File Length → 10 minutes
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Length").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "10 minutes")

    # Log File Name Format → Time interval Format（单选按钮）
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log File Name Format").first.locator(
        ".el-radio"
    ).filter(has_text="Time interval Format").click()
    page.wait_for_timeout(300)

    # Log Interval → 1 minute
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    page.locator(".el-form-item").filter(has_text="Log Interval").first.locator(".el-select").click()
    page.wait_for_timeout(400)
    _click_option(page, "1 minute")

    # Log File Name Prefix：输入超过20字符（21字符）
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    prefix_input = page.locator(".el-form-item").filter(
        has_text="Log File Name Prefix"
    ).first.locator("input")
    prefix_input.fill("meter3_logger1_123456")  # 21 chars — 超出20字符限制
    page.wait_for_timeout(300)
    page.locator("body").click()  # 触发 blur 激活实时校验
    page.wait_for_timeout(400)

    # 点击 Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # 断言：出现表单校验错误（实时校验或保存时校验均可）
    form_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert form_errors > 0 or msg_errors > 0, (
        "Logger3 LogFileNamePrefix超过20字符时保存应失败并显示错误提示，"
        f"但未检测到任何错误（form errors={form_errors}, message errors={msg_errors}）"
    )
