import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


# 用例编号：TestCase_AcuHMI_003_05_case01
# 用例标题：Post Ch1设置为disable，Logger数据记录PostChannel选项无法选中Post Ch1
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 Post Channels > Post Channel 1，将开关置为 Disable 并保存
#   2. 进入 Data Loggers > Data Loggers 1，Enable Logger 1
#   3. 打开 Post Channel 下拉，查看 Channel1 是否不可选（disabled 或不存在）
# 预期结果：
#   1. 保存成功
#   2. Logger 1 Post Channel 下拉中 Channel1 选项不可选或不存在


def _nav_to_post_channel1(page):
    """Navigate to Data Log > Post Channels > Post Channel 1."""
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

    # 展开 Post Channels Tab
    tab = page.locator("div.el-sub-menu__title").filter(has_text="Post Channels")
    if tab.count() > 0 and tab.first.is_visible():
        tab.first.click()
        page.wait_for_timeout(400)

    # 点击 Post Channel 1
    item = page.locator(".el-menu-item").filter(has_text="Post Channel 1")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _nav_to_data_loggers1(page):
    """Navigate to Data Log > Data Loggers > Data Loggers 1."""
    tab = page.locator("div.el-sub-menu__title").filter(has_text="Data Loggers")
    if tab.count() > 0 and tab.first.is_visible():
        tab.first.click()
        page.wait_for_timeout(400)

    item = page.locator(".el-menu-item").filter(has_text="Data Loggers 1")
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def test_TestCase_AcuHMI_003_05_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 进入 Post Channel 1，设为 Disable 并保存
    _nav_to_post_channel1(page)
    page.wait_for_timeout(500)

    page.locator(".el-radio").filter(has_text="Disable").click()
    page.wait_for_timeout(300)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # 验证保存成功（无 error toast）
    assert page.locator(".el-message--error").count() == 0, \
        "Post Channel 1 Disable 保存应成功，但出现了错误提示"

    # Step 2: 进入 Data Loggers 1，Enable Logger 1
    _nav_to_data_loggers1(page)
    page.wait_for_timeout(800)

    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and enable_radio.first.is_visible():
        enable_radio.first.click()
        page.wait_for_timeout(500)

    # Step 3: 打开 Post Channel 下拉，验证 Channel1 不可选
    post_ch_select = page.locator(".el-form-item").filter(
        has_text="Post Channel"
    ).first.locator(".el-select")
    post_ch_select.click()
    page.wait_for_timeout(400)

    # 获取所有可见下拉选项
    all_items = page.locator(".el-select-dropdown__item").all()
    channel1_item = None
    for it in all_items:
        try:
            if it.is_visible() and "Channel1" in it.inner_text():
                channel1_item = it
                break
        except Exception:
            pass

    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    if channel1_item is None:
        # Channel1 不在选项列表中 — 符合预期（disabled 后被移除）
        pass
    else:
        cls = channel1_item.get_attribute("class") or ""
        assert "disabled" in cls, (
            "Post Channel 1 Disable后，Logger 1 的 Post Channel 下拉中 "
            f"Channel1 选项应不可选，但 class='{cls}'"
        )
