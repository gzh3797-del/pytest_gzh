import pytest
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_033
# 用例标题: COV Batch Update 配置 AcuRev-4100 参数保存正常
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 进入 BACnet/IP，启用 BACnet，点击 AcuRev4100 的 Parameter Config
#   2. 在参数表中对第一行开启 Polling Enable + COV Enable（el-switch）
#   3. 点击 COV Batch Update 按钮，打开 Batch Update 子对话框
#   4. 在 Parameters 下拉选择参数，输入 COV Increment=2.000，点击 Confirm
#   5. 在 Parameter Config 对话框中点击 Save，验证保存成功
# 预期结果:
#   4. Batch Update Confirm 无错误
#   5. Parameter Config Save 无错误


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


def _enable_switch_in_cell(row, cell_index: int, page):
    """Enable el-switch in the given cell of a table row if not already on."""
    cell = row.locator("td").nth(cell_index)
    switch_input = cell.locator("input[type='checkbox']").first
    if not switch_input.is_checked():
        cell.locator(".el-switch__core").first.click()
        page.wait_for_timeout(300)


def test_TestCase_WEB2_033_001_033(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "BACnet/IP")

    # Step 1: 启用 BACnet
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    if "is-checked" not in (bacnet_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # 找到设备表格
    device_table_fi = None
    for fi in page.locator(".el-form-item").all():
        if fi.locator(".el-switch, .el-checkbox").count() > 3:
            device_table_fi = fi
            break
    assert device_table_fi is not None, "未找到设备选择表格"

    # 找 AcuRev4100 行并打开 Parameter Config
    acurev_row = device_table_fi.locator("tbody tr").filter(has_text="AcuRev4100")
    assert acurev_row.count() > 0, "设备列表中未找到 AcuRev4100"
    acurev_row.first.locator(".el-button--primary").first.click()
    page.wait_for_timeout(1000)

    param_dialog = page.locator(".el-dialog").filter(has_text="Parameter Config").first
    assert param_dialog.is_visible(), "Parameter Config 对话框未打开"

    # Step 2: 对参数表格中所有行开启 Polling Enable（cell[1]）和 COV Enable（cell[2]）
    rows = param_dialog.locator("tbody tr").all()
    assert len(rows) > 0, "Parameter Config 对话框中参数表格无数据行"
    for row in rows:
        try:
            _enable_switch_in_cell(row, 1, page)  # Polling Enable
            _enable_switch_in_cell(row, 2, page)  # COV Enable
        except Exception:
            pass

    # Step 3: 点击 COV Batch Update 按钮，打开 Batch Update 子对话框
    param_dialog.get_by_role("button").filter(has_text="COV Batch Update").first.click()
    page.wait_for_timeout(1000)

    batch_dialog = page.locator(".el-dialog").filter(has_text="Batch Update").last
    assert batch_dialog.is_visible(), "Batch Update 子对话框未打开"

    # Step 4: 在 Parameters 下拉选择所有参数
    params_select = batch_dialog.locator(".el-form-item").filter(has_text="Parameters").first.locator(".el-select").first
    params_select.click()
    page.wait_for_timeout(600)

    # 选择所有可见选项（多选下拉）
    selected = 0
    opts = page.locator(".el-select-dropdown__item").all()
    for opt in opts:
        try:
            if opt.is_visible():
                opt.click()
                page.wait_for_timeout(150)
                selected += 1
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(300)

    assert selected > 0, \
        "Parameters 下拉中未找到可选参数（请确认已开启 Polling Enable 和 COV Enable）"

    # 填写 COV Increment = 2.000
    cov_inc_input = batch_dialog.locator(".el-form-item").filter(has_text="COV Increment").first.locator("input").first
    cov_inc_input.fill("2.000")
    page.wait_for_timeout(200)

    # 点击 Confirm
    batch_dialog.get_by_role("button").filter(has_text="Confirm").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    # 验证 Confirm 后无错误
    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"COV Batch Update Confirm 应成功（field_errors={field_errors}, msg_errors={msg_errors}）"

    # Step 5: 在 Parameter Config 对话框中点击 Save
    assert param_dialog.is_visible(), "Confirm 后 Parameter Config 对话框应仍然可见"
    param_dialog.get_by_role("button").filter(has_text="Save").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    save_errors = page.locator(".el-form-item__error").count()
    save_msg_errors = page.locator(".el-message--error").count()
    assert save_errors == 0 and save_msg_errors == 0, \
        f"Parameter Config Save 应成功（field_errors={save_errors}, msg_errors={save_msg_errors}）"

    # Step 6: 重新打开 Parameter Config，验证 COV Increment 列值已回填为 2.000
    # 等待对话框关闭
    page.wait_for_timeout(500)

    # 重新找 AcuRev4100 行并打开 Parameter Config
    device_table_fi2 = None
    for fi in page.locator(".el-form-item").all():
        if fi.locator(".el-switch, .el-checkbox").count() > 3:
            device_table_fi2 = fi
            break
    assert device_table_fi2 is not None, "重新打开时未找到设备表格"

    acurev_row2 = device_table_fi2.locator("tbody tr").filter(has_text="AcuRev4100")
    acurev_row2.first.locator(".el-button--primary").first.click()
    page.wait_for_timeout(1000)

    param_dialog2 = page.locator(".el-dialog").filter(has_text="Parameter Config").first
    assert param_dialog2.is_visible(), "重新打开 Parameter Config 对话框失败"

    # 检查 COV Increment 列（cell[3]）值是否与配置一致（2 或 2.0 或 2.000）
    rows2 = param_dialog2.locator("tbody tr").all()
    mismatch = []
    for i, row in enumerate(rows2):
        try:
            # COV Enable 状态
            cov_enable_cell = row.locator("td").nth(2)
            cov_enabled = cov_enable_cell.locator("input[type='checkbox']").first.is_checked()
            if not cov_enabled:
                continue  # 未开启 COV Enable 的行不检查
            # COV Increment 值（cell[3]）
            cov_inc_input = row.locator("td").nth(3).locator("input[type='text']").first
            val = cov_inc_input.input_value().strip()
            # 接受 "2", "2.0", "2.00", "2.000" 等格式
            try:
                numeric_val = float(val) if val else None
            except ValueError:
                numeric_val = None
            if numeric_val != 2.0:
                mismatch.append(f"Row[{i}]: COV Increment='{val}' (期望 2.0)")
        except Exception:
            pass

    assert len(mismatch) == 0, \
        f"COV Increment 回填值与配置不一致：{mismatch}"
