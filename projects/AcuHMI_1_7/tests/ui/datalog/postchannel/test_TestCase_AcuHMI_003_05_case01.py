from playwright.sync_api import Page
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 下拉里 Post Channel 1 选项的真实文案（带空格）
_POST_CHANNEL_1 = "Post Channel 1"


def _enter_data_log(page: Page):
    """进入 Data Log 模块（若当前不在 dataLog 页）。"""
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


def _nav_submenu(page: Page, tab_text: str, item_text: str):
    """在 Data Log 顶部菜单展开 tab_text 并点击其子项 item_text。"""
    _enter_data_log(page)
    tab = page.locator("div.el-sub-menu__title").filter(has_text=tab_text)
    if tab.count() > 0 and tab.first.is_visible():
        tab.first.click()
        page.wait_for_timeout(400)
    item = page.locator(".el-menu-item").filter(has_text=item_text)
    if item.count() > 0 and item.first.is_visible():
        item.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _click_radio(page: Page, label: str):
    """点击 Enable/Disable 单选。优先 locator 点击；被 popper 遮挡时降级坐标点击。"""
    radio = page.locator(".el-radio").filter(has_text=label).first
    radio.scroll_into_view_if_needed()
    try:
        radio.click(timeout=3000)
    except Exception:
        # el-menu popper 偶发遮挡事件链：移开鼠标后按坐标点击（conventions 允许的降级）
        page.mouse.move(400, 300)
        page.wait_for_timeout(200)
        box = radio.bounding_box()
        if box is not None:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
    page.wait_for_timeout(400)


def _find_post_channel1_class(page: Page):
    """读取当前展开的 Post Channel 下拉里 'Post Channel 1' 选项的 class；不存在返回 None。"""
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if item.is_visible() and item.inner_text().strip() == _POST_CHANNEL_1:
                return item.get_attribute("class") or ""
        except Exception:
            pass
    return None


# 用例编号：TestCase_AcuHMI_003_05_case01
# 用例标题：Post Ch1 设置为 disable，Logger 数据记录 Post Channel 选项无法选中 Post Ch1
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. 进入 Post Channels > Post Channel 1，将开关置为 Disable 并保存
#   2. 进入 Data Loggers > Data Loggers 1，Enable Logger 1
#   3. 打开 Post Channel 下拉，检查 Post Channel 1 是否不可选
# 预期结果：
#   1. 保存成功
#   2. Logger 1 的 Post Channel 下拉中 Post Channel 1 选项不可选（带 disabled 类）
def test_TestCase_AcuHMI_003_05_case01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1：进入 Post Channel 1，置为 Disable 并保存
    _nav_submenu(page, "Post Channels", "Post Channel 1")
    page.wait_for_timeout(500)

    _click_radio(page, "Disable")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)
    assert page.locator(".el-message--error").count() == 0, \
        "Post Channel 1 置为 Disable 保存应成功，但出现了错误提示"

    # Step 2：进入 Data Loggers 1 并 Enable，Post Channel 配置栏才会出现
    _nav_submenu(page, "Data Loggers", "Data Loggers 1")
    page.wait_for_timeout(800)
    _click_radio(page, "Enable")

    # Step 3：打开 Post Channel 下拉，校验 Post Channel 1 不可选
    pc_select = page.locator(".el-form-item").filter(
        has_text="Post Channel"
    ).first.locator(".el-select")
    assert pc_select.count() > 0, "Logger 1 Enable 后未出现 Post Channel 下拉，无法校验"
    pc_select.first.click()
    page.wait_for_timeout(400)

    cls = _find_post_channel1_class(page)
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    # 断言“必须找到”防止下拉未展开时的空转假通过；“带 disabled”捕获“仍可选”这一缺陷
    assert cls is not None, \
        "Post Channel 下拉中未找到 'Post Channel 1' 选项，无法确认其不可选"
    assert "disabled" in cls, \
        f"Post Channel 1 Disable 后，Logger 1 下拉中该项应不可选，但 class='{cls}'"
