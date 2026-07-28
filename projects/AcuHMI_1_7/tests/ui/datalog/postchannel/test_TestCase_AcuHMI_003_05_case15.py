from playwright.sync_api import Page
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 四个 Data Logger 的子菜单文案（Rapid Logger 无编号）
_LOGGER_ITEMS = ["Data Loggers 1", "Data Loggers 2", "Data Loggers 3", "Rapid Logger"]
# 需逐一确认不可选的 Post Channel（下拉选项真实文案，带空格）
_POST_CHANNEL_NAMES = ["Post Channel 1", "Post Channel 2", "Post Channel 3"]


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


def _set_post_channel_disable(page: Page, pc_name: str):
    """进入指定 Post Channel，置为 Disable 并保存，校验保存成功。"""
    _nav_submenu(page, "Post Channels", pc_name)
    page.wait_for_timeout(500)

    _click_radio(page, "Disable")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)
    assert page.locator(".el-message--error").count() == 0, \
        f"{pc_name} 置为 Disable 保存应成功，但出现了错误提示"


def _collect_pc_option_classes(page: Page) -> dict:
    """读取当前展开的 Post Channel 下拉里 PC1/2/3 选项的 class（不在列表则缺省不收录）。"""
    result: dict[str, str] = {}
    for item in page.locator(".el-select-dropdown__item").all():
        try:
            if not item.is_visible():
                continue
            txt = item.inner_text().strip()
            if txt in _POST_CHANNEL_NAMES:
                result[txt] = item.get_attribute("class") or ""
        except Exception:
            pass
    return result


def _assert_pc_not_selectable(page: Page, logger_item: str, failures: list):
    """进入指定 Logger，Enable 后打开 Post Channel 下拉，断言 PC1/2/3 均不可选。"""
    _nav_submenu(page, "Data Loggers", logger_item)
    page.wait_for_timeout(800)

    # Enable 该 Logger，Post Channel 配置栏才会出现
    _click_radio(page, "Enable")

    pc_select = page.locator(".el-form-item").filter(
        has_text="Post Channel"
    ).first.locator(".el-select")
    if pc_select.count() == 0:
        failures.append(f"[{logger_item}] Enable 后未出现 Post Channel 下拉，无法校验")
        return
    pc_select.first.click()
    page.wait_for_timeout(400)

    option_classes = _collect_pc_option_classes(page)
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    # 预期：PC 全 Disable 时，三项均出现在下拉且带 disabled 类（不可选）。
    # 断言“必须找到”可防止下拉未展开/选项未采到时的空转假通过；
    # 断言“带 disabled”则捕获“选项仍可选”这一真实缺陷。
    for name in _POST_CHANNEL_NAMES:
        cls = option_classes.get(name)
        if cls is None:
            failures.append(
                f"[{logger_item}] Post Channel 下拉中未找到 '{name}' 选项，无法确认其不可选"
            )
        elif "disabled" not in cls:
            failures.append(
                f"[{logger_item}] '{name}' 应不可选，但下拉中该项可选（class='{cls}'）"
            )


# 用例编号：TestCase_AcuHMI_003_05_case15
# 用例标题：Post Ch1~3 全部 Disable 后，各 Data Logger 均无法选择 Post Channel 1/2/3
# 预置条件：
#   1. 服务启动正常，账号登录成功
# 测试步骤：
#   1. Devices > Data Log > Post Channels > Post Channel 1/2/3，全部置为 Disable 并保存
#   2. Data Loggers > Data Loggers 1，Enable，Post Channel 栏检查 Post Channel 1/2/3
#   3. Data Loggers > Data Loggers 2，Enable，Post Channel 栏检查 Post Channel 1/2/3
#   4. Data Loggers > Data Loggers 3，Enable，Post Channel 栏检查 Post Channel 1/2/3
#   5. Data Loggers > Rapid Logger，Enable，Post Channel 栏检查 Post Channel 1/2/3
# 预期结果：
#   1. 保存配置成功
#   2~5. 四个 Logger 的 Post Channel 下拉均无法选择 Post Channel 1/2/3
def test_TestCase_AcuHMI_003_05_case15(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1：PC1/2/3 全部置为 Disable 并保存
    for pc_name in _POST_CHANNEL_NAMES:
        _set_post_channel_disable(page, pc_name)

    # Step 2~5：四个 Logger 逐一 Enable 并校验 Post Channel 下拉不可选
    failures: list[str] = []
    for logger_item in _LOGGER_ITEMS:
        _assert_pc_not_selectable(page, logger_item, failures)

    assert not failures, "\n".join(failures)
