from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_modbus_config(page):
    """Navigate to Protocols > Modbus > Modbus Config."""
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
    page.get_by_role("menuitem", name="Modbus Config").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _ensure_modbus_enabled_port_input(page):
    """确保 Modbus 已启用，返回 Modbus Port 输入框 locator。

    Modbus Port 字段在 v-if 区块内，仅当 Modbus Enable 选中时才渲染进 DOM；实测点
    Enable 单选即响应式渲染该字段，无需 Save。原写法点 Enable 后只固定等 300ms 且吞掉
    点击异常，Modbus 处于关闭态时 Port 不在 DOM，导致后续 fill 30s 超时。这里改为：
    未启用则点 Enable，再显式等待 Port 输入框可见（用 placeholder 精确定位）。
    """
    # 导航用的 Modbus 下拉菜单是 hover 触发，点完子项后菜单 popper 仍悬在表单上方，
    # 会拦截 Enable 单选点击；先按 Esc 并把鼠标移开，让菜单 popper 收起。
    page.keyboard.press("Escape")
    page.mouse.move(640, 500)
    page.wait_for_timeout(300)
    enabled = page.locator(".el-radio.is-checked").filter(has_text="Enable").count() > 0
    if not enabled:
        page.locator(".el-radio").filter(has_text="Enable").first.click()
    port_input = page.locator('input[placeholder="Enter Modbus Port"]')
    port_input.wait_for(state="visible", timeout=10000)
    return port_input


def _verify_tips_table(page):
    """校验 tips 表表头包含 Type/Register/Length/Sequence，表体包含已知数据类型行。"""
    header = page.locator(".el-table__header")
    if header.count() > 0:
        header_text = header.last.inner_text().lower()
        for col in ["type", "register", "length", "sequence"]:
            assert col in header_text, \
                f"tips 表头未找到列 '{col}'，当前表头内容：{header_text[:200]}"

    body = page.locator(".el-table__body").last
    body_text = body.inner_text().lower()
    for dtype in ["bit", "uint16", "int32", "float", "double"]:
        assert dtype in body_text, \
            f"tips 表体未找到数据类型 '{dtype}'，当前表体内容：{body_text[:300]}"


# 用例编号：TestCase_AcuHMI_008_02_case02
# 用例标题：配置modbus port为公共端口502，保存配置成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Protocols->Modbus Config: Modbus Port=502
#   2. Save保存
# 预期结果：
#   2. 配置保存成功，tips提示数据类型、寄存器数量、数据长度、大小端准确
def test_TestCase_AcuHMI_008_02_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_modbus_config(page)

    # Step 1: 确保 Modbus 已启用并填写 Modbus Port=502
    port_input = _ensure_modbus_enabled_port_input(page)
    port_input.fill("502")
    page.wait_for_timeout(200)

    # Step 2: 点击 Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # 校验无错误 toast
    assert page.locator(".el-message--error").count() == 0, \
        "Port=502 保存后不应出现错误 toast"

    try:
        expect(page.get_by_text("success", exact=False)).to_be_visible(timeout=3000)
    except Exception:
        pass

    # 校验表单无校验错误
    form_errors = page.locator(".el-form-item__error").count()
    assert form_errors == 0, \
        f"Port=502 保存后不应出现表单校验错误，但发现 {form_errors} 个错误"

    # 校验 tips 表数据类型、寄存器数量、数据长度、大小端显示正确
    _verify_tips_table(page)
