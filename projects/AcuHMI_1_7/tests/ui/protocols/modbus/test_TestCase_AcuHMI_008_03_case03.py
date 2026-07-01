from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_pass_through(page):
    """Navigate to Protocols > Modbus > Pass Through, then click Enable to show content."""
    # 等待上一步可能残留的 toast 消失（El UI 默认 3s）
    page.wait_for_timeout(3000)

    if "/protocols/" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/maintenance",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Protocols").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Modbus").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    page.get_by_role("menuitem", name="Pass Through").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # 点击 Enable 使页面内容展开
    _ensure_enabled(page)


def _ensure_enabled(page):
    """收起导航遗留的 hover 菜单 popper，确保 Pass Through 已启用，并等待设备行出现。

    Pass Through 的设备行表格受 Enable 门控（v-if）：点 Enable 单选即响应式渲染，无需
    Save；行来自已配置的下挂 Modbus 设备。导航点完子菜单后 hover 菜单 popper 仍悬在表单
    上方会拦截 Enable 单选点击（原写法因此点不上 Enable，表格 0 行而失败），故先收起再点。
    """
    # 收起 hover 菜单 popper（否则拦截 Enable 单选点击）
    page.keyboard.press("Escape")
    page.mouse.move(640, 500)
    page.wait_for_timeout(300)

    # 未启用则点 Enable（响应式渲染表格，无需 Save）
    enabled = page.locator(".el-radio.is-checked").filter(has_text="Enable").count() > 0
    if not enabled:
        page.locator(".el-radio").filter(has_text="Enable").first.click()
        page.wait_for_load_state("networkidle")

    # 轮询等待设备行出现（行来自下挂设备，需等后端列表渲染）
    rows = page.locator("tbody tr")
    for _ in range(15):
        if rows.count() > 0:
            break
        page.wait_for_timeout(200)
    assert rows.count() > 0, \
        "启用 Pass Through 后设备行表格仍为空（0 行）——请确认设备已配置下挂 Modbus 设备"


def _set_first_row_slave_id(page, slave_id: int):
    """勾选第一行 Enable 复选框，并设置该行的 Slave ID。"""
    rows = page.locator("tbody tr")
    rows.first.wait_for(timeout=5000)
    first_row = rows.first

    # 勾选行首的 Enable 复选框（若未勾选）
    checkbox = first_row.locator(".el-checkbox__original, input[type='checkbox']").first
    if checkbox.count() > 0 and not checkbox.is_checked():
        first_row.locator(".el-checkbox__inner, .el-checkbox").first.click()
        page.wait_for_timeout(300)

    # 填写 Slave ID
    slave_id_input = first_row.locator("input:not(.el-checkbox__original):not([type='checkbox'])").first
    slave_id_input.fill(str(slave_id))
    page.wait_for_timeout(200)


def _click_save(page):
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)


def _get_field_error_text(page) -> str:
    """返回第一个 .el-form-item__error 的文本，若无则返回空字符串。"""
    tip = page.locator(".el-form-item__error").first
    if tip.count() > 0 and tip.is_visible():
        return tip.inner_text().strip()
    return ""


def _is_rejected(page) -> bool:
    """返回 True 表示保存被拒绝（表单错误或错误 toast）。"""
    if page.locator(".el-form-item__error").count() > 0:
        return True
    if page.locator(".el-message--error").count() > 0:
        return True
    if page.locator(".el-message-box").count() > 0:
        for btn in ["OK", "Cancel"]:
            try:
                page.get_by_role("button", name=btn).click(timeout=2000)
                break
            except Exception:
                pass
        return True
    return False


def _is_saved_successfully(page) -> bool:
    """返回 True 表示保存成功（无错误 toast 且无表单错误）。"""
    if page.locator(".el-message--error").count() > 0:
        return False
    try:
        expect(page.get_by_text("success", exact=False)).to_be_visible(timeout=3000)
        return True
    except Exception:
        return page.locator(".el-form-item__error").count() == 0


# 用例编号：TestCase_AcuHMI_008_03_case03
# 用例标题：Pass Through Slave ID边界验证：<101保存失败，>247保存失败，101/247保存成功
# 预置条件：
#   1. 管理员账号登录AcuHMI-1-7网页
# 测试步骤：
#   1. Protocols→Modbus→Pass Through，点击Enable展开内容
#   2. 将某设备Slave ID设置为100（低于最小值101），点击Save
#   3. Slave ID设置为248（高于最大值247），点击Save
#   4. Slave ID设置为边界最小值101，点击Save
#   5. Slave ID设置为边界最大值247，点击Save
# 预期结果：
#   2. 保存失败，输入框提示Slave ID超出有效范围（101–247）
#   3. 保存失败，输入框有提示信息
#   4. 保存成功
#   5. 保存成功
def test_TestCase_AcuHMI_008_03_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 2: Slave ID=100（低于最小值101），应保存失败且有提示信息
    _nav_pass_through(page)
    _set_first_row_slave_id(page, 100)
    _click_save(page)
    assert _is_rejected(page), \
        "Slave ID=100 低于最小值101，保存应失败，但系统接受了该值"
    tip = _get_field_error_text(page)
    assert len(tip) > 0, \
        f"Slave ID=100 保存失败后，输入框下方应有提示信息，但未检测到"
    print(f"\nSlave ID=100 提示信息：{tip}")

    # Step 3: Slave ID=248（高于最大值247），应保存失败且有提示信息
    _nav_pass_through(page)
    _set_first_row_slave_id(page, 248)
    _click_save(page)
    assert _is_rejected(page), \
        "Slave ID=248 高于最大值247，保存应失败，但系统接受了该值"
    tip = _get_field_error_text(page)
    assert len(tip) > 0, \
        f"Slave ID=248 保存失败后，输入框下方应有提示信息，但未检测到"
    print(f"Slave ID=248 提示信息：{tip}")

    # Step 4: Slave ID=101（边界最小值），应保存成功
    _nav_pass_through(page)
    _set_first_row_slave_id(page, 101)
    _click_save(page)
    assert _is_saved_successfully(page), \
        "Slave ID=101 为边界最小值，保存应成功，但出现错误"

    # Step 5: Slave ID=247（边界最大值），应保存成功
    _nav_pass_through(page)
    _set_first_row_slave_id(page, 247)
    _click_save(page)
    assert _is_saved_successfully(page), \
        "Slave ID=247 为边界最大值，保存应成功，但出现错误"
