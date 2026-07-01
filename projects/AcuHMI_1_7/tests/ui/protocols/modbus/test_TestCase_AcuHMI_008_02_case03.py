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


def _check_invalid_port(page, port: int):
    """
    填入越界端口，校验：
    1. 输入框下方出现提示信息（非空）
    2. 提示信息拼写合理（含数字或 port/range 等关键词）
    3. 点击 Save 后无法保存成功（存在表单错误或错误 toast）
    """
    _nav_modbus_config(page)

    # 确保 Modbus 已启用并取 Modbus Port 输入框
    port_input = _ensure_modbus_enabled_port_input(page)
    port_input.fill(str(port))
    page.wait_for_timeout(200)

    # 点击 Save 触发表单校验（El UI 通常在 submit 时才渲染 .el-form-item__error）
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    # 处理可能弹出的警告弹窗
    if page.locator(".el-message-box").count() > 0:
        for btn in ["OK", "Cancel"]:
            try:
                page.get_by_role("button", name=btn).click(timeout=2000)
                break
            except Exception:
                pass

    # 1. 校验输入框下方出现提示信息
    error_tip = page.locator(".el-form-item__error").first
    tip_visible = error_tip.count() > 0 and error_tip.is_visible()
    assert tip_visible, \
        f"Port={port} 超出范围，点击 Save 后输入框下方应出现提示信息，但未检测到 .el-form-item__error"

    # 2. 读取提示文本，校验非空且拼写合理
    tip_text = error_tip.inner_text().strip()
    assert len(tip_text) > 0, \
        f"Port={port} 提示信息为空字符串，应有具体说明"

    tip_lower = tip_text.lower()
    has_range_hint = any(k in tip_lower for k in ["2000", "5999", "port", "range", "between", "valid", "1999", "6000"])
    assert has_range_hint, \
        f"Port={port} 提示信息内容可能有误，未包含范围说明关键词。\n实际提示：'{tip_text}'"

    # 3. 确认无成功 toast（保存被阻止）
    assert page.locator(".el-message--success").count() == 0, \
        f"Port={port} 超出范围，不应出现成功 toast，但系统显示了保存成功"

    return tip_text


# 用例编号：TestCase_AcuHMI_008_02_case03
# 用例标题：配置modbus port为1999、6000（越界），保存配置失败
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Protocols->Modbus Config: Modbus Port=1999
#   2. 查看输入框提示信息，点击Save
#   3. 端口设置为6000，重复步骤1-2
# 预期结果：
#   输入框超过限制会有提示信息，且点击Save无法保存
def test_TestCase_AcuHMI_008_02_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1-2: Port=1999（低于下界2000）
    tip_1999 = _check_invalid_port(page, port=1999)
    print(f"\nPort=1999 提示信息：{tip_1999}")

    # Step 3: Port=6000（高于上界5999）
    tip_6000 = _check_invalid_port(page, port=6000)
    print(f"Port=6000 提示信息：{tip_6000}")
