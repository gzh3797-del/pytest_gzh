import pytest
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_045
# 用例标题: COV Increment 非法值被拦截
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 进入 BACnet/IP 页面，启用 BACnet，打开第一个设备的 Parameter Config
#   2. 在参数表格中启用 COV Enable（el-switch），使 COV Increment 可编辑
#   3. 输入 COV Increment=-0.001 并点击 Save
# 预期结果: 小于 0 的非法值被拦截，保存失败并出现错误提示


def _nav_protocol(page, protocol: str, sub: str = None):
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
    page.get_by_role("menuitem", name=protocol).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    if sub:
        page.get_by_role("menuitem", name=sub).click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _open_param_config_dialog(page):
    """Enable BACnet and open Parameter Config dialog for the first device."""
    # Enable BACnet
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    if "is-checked" not in (bacnet_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Find device table and click first row's Parameter Config button
    device_table_fi = None
    for fi in page.locator(".el-form-item").all():
        if fi.locator(".el-switch, .el-checkbox").count() > 3:
            device_table_fi = fi
            break
    assert device_table_fi is not None, "未找到设备选择表格"

    first_row = device_table_fi.locator("tbody tr").first
    first_row.locator(".el-button--primary").first.click()
    page.wait_for_timeout(1000)

    dialog = page.locator(".el-dialog").filter(has_text="Parameter Config").first
    assert dialog.count() > 0, "Parameter Config 对话框未打开"
    return dialog


def test_TestCase_WEB2_033_001_045(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "BACnet/IP")
    dialog = _open_param_config_dialog(page)

    # Step 1: 验证 COV Enable / COV Increment 列存在
    headers = [h.inner_text().strip() for h in dialog.locator("thead th").all()]
    assert "COV Enable" in headers, f"表头应含 COV Enable，实际：{headers}"
    assert "COV Increment" in headers, f"表头应含 COV Increment，实际：{headers}"

    # Step 2: 对第一行启用 COV Enable（el-switch，点击 .el-switch__core）
    first_row = dialog.locator("tbody tr").first
    cov_enable_cell = first_row.locator("td").nth(2)   # COV Enable 列
    cov_enable_switch_input = cov_enable_cell.locator("input[type='checkbox']").first

    if not cov_enable_switch_input.is_checked():
        cov_enable_cell.locator(".el-switch__core").first.click()
        page.wait_for_timeout(400)

    assert cov_enable_switch_input.is_checked(), "COV Enable 开关应已开启"

    # COV Increment 应变为可编辑
    cov_inc_cell = first_row.locator("td").nth(3)      # COV Increment 列
    cov_inc_input = cov_inc_cell.locator("input[type='text']").first
    assert not cov_inc_input.get_attribute("disabled"), \
        "COV Enable 开启后 COV Increment 应变为可编辑"

    # Step 3: 输入非法值 -0.001 并保存
    cov_inc_input.fill("-0.001")
    page.wait_for_timeout(200)

    dialog.locator("button").filter(has_text="Save").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    # 验证保存被阻止
    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    dialog_errors = dialog.locator(".el-form-item__error, .error, [class*='error']").count()

    assert field_errors > 0 or msg_errors > 0 or dialog_errors > 0, \
        f"COV Increment=-0.001 应被拦截并出现错误提示（field_errors={field_errors}, msg_errors={msg_errors}, dialog_errors={dialog_errors}）"
