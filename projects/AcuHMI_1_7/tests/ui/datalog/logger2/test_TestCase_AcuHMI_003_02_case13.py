import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_02_case13
# 用例标题：验证logger2的log File length、log interval逻辑关系
# 预置条件：同上
# 测试步骤：
#   1.  log file length为1 minute时，选择log interval
#   2.  log file length为5 minutes时，选择log interval
#   3-11. 各file length对应的可选interval验证
# 预期结果：
#   各file length对应的interval可选范围正确（联动关系同Logger1）

_LOGGER_LINKAGE = {
    "1 minute":   ["1 minute"],
    "5 minutes":  ["1 minute", "5 minutes"],
    "10 minutes": ["1 minute", "5 minutes", "10 minutes"],
    "15 minutes": ["1 minute", "5 minutes", "10 minutes", "15 minutes"],
    "30 minutes": ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes"],
    "1 hour":     ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour"],
    "6 hours":    ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours"],
    "12 hours":   ["1 minute", "5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours"],
    "1 day":      ["5 minutes", "10 minutes", "15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day"],
    "7 days":     ["15 minutes", "30 minutes", "1 hour", "6 hours", "12 hours", "1 day", "7 days"],
    "1 month":    ["1 hour", "6 hours", "12 hours", "1 day", "7 days", "1 month"],
}


def _nav_to_data_loggers2(page):
    """Navigate to Data Log > Data Loggers > Data Loggers 2."""
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

    item = page.locator(".el-menu-item").filter(has_text="Data Loggers 2")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _get_visible_options(page) -> list:
    """Return texts of currently visible dropdown items."""
    all_items = page.locator(".el-select-dropdown__item").all()
    result = []
    for item in all_items:
        try:
            if item.is_visible():
                txt = item.inner_text().strip()
                if txt:
                    result.append(txt)
        except Exception:
            pass
    return result


def _select_file_length(page, length_text: str):
    """Open Log File Length dropdown and select the given option."""
    lfl_select = page.locator(".el-form-item").filter(
        has_text="Log File Length"
    ).first.locator(".el-select")
    lfl_select.click()
    page.wait_for_timeout(400)
    all_items = page.locator(".el-select-dropdown__item").all()
    for item in all_items:
        try:
            if item.is_visible() and item.inner_text().strip() == length_text:
                item.click()
                break
        except Exception:
            pass
    page.wait_for_timeout(400)


def _get_interval_options(page) -> list:
    """Open the Log Interval dropdown and return visible option texts."""
    li_select = page.locator(".el-form-item").filter(
        has_text="Log Interval"
    ).first.locator(".el-select")
    li_select.click()
    page.wait_for_timeout(400)
    options = _get_visible_options(page)
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return options


def test_TestCase_AcuHMI_003_02_case13(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_data_loggers2(page)
    page.wait_for_timeout(800)

    # 开启 Logger 2
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(500)

    failures = []

    for file_length, expected_intervals in _LOGGER_LINKAGE.items():
        _select_file_length(page, file_length)

        actual_options = _get_interval_options(page)

        for expected in expected_intervals:
            if expected not in actual_options:
                failures.append(
                    f"[Logger2] FileLength={file_length}: "
                    f"期望interval选项'{expected}'存在，实际选项={actual_options}"
                )

        for actual in actual_options:
            if actual not in expected_intervals:
                failures.append(
                    f"[Logger2] FileLength={file_length}: "
                    f"不期望出现interval选项'{actual}'，期望选项={expected_intervals}"
                )

    assert not failures, "\n".join(failures)
