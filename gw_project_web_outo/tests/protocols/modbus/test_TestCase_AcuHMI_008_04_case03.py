import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_device_mirror(page):
    """Navigate to Protocols > Modbus > Device Mirror, click Enable."""
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
    page.get_by_role("menuitem", name="Device Mirror").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # 点击 Enable radio 使功能开启
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0:
        enable_radio.first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _find_editable_row(page):
    """
    遍历所有行，返回第一个 Enable 复选框可点击（非 disabled）的行对象。
    若找不到可用行则断言失败。
    """
    rows = page.locator("tbody tr")
    row_count = rows.count()
    for i in range(row_count):
        row = rows.nth(i)
        checkbox = row.locator(".el-checkbox__original, input[type='checkbox']").first
        if checkbox.count() > 0 and checkbox.is_enabled():
            return row
    pytest.fail(f"在 {row_count} 行中未找到 Enable 复选框可操作的行，请确认设备配置")


def _set_row_slave_id(page, slave_id: int):
    """找到可编辑行，勾选其 Enable 复选框，填写 Slave ID。"""
    row = _find_editable_row(page)

    # 勾选该行的 Enable 复选框（若未勾选）
    checkbox = row.locator(".el-checkbox__original, input[type='checkbox']").first
    if not checkbox.is_checked():
        row.locator(".el-checkbox__inner").first.click()
        page.wait_for_timeout(300)

    # 等待 Slave ID 输入框可用后填值
    slave_id_input = row.locator("input:not(.el-checkbox__original):not([type='checkbox'])").first
    try:
        expect(slave_id_input).to_be_enabled(timeout=3000)
    except Exception:
        pass
    slave_id_input.fill(str(slave_id))
    page.wait_for_timeout(200)


def _click_save(page):
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)


def _get_field_error_text(page) -> str:
    tip = page.locator(".el-form-item__error").first
    if tip.count() > 0 and tip.is_visible():
        return tip.inner_text().strip()
    return ""


def _is_rejected(page) -> bool:
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
    if page.locator(".el-message--error").count() > 0:
        return False
    try:
        expect(page.get_by_text("success", exact=False)).to_be_visible(timeout=3000)
        return True
    except Exception:
        return page.locator(".el-form-item__error").count() == 0


# 用例编号：TestCase_AcuHMI_008_04_case03
# 用例标题：Device Mirror Slave ID边界验证：<2保存失败，>99保存失败，2/99保存成功
# 预置条件：
#   1. 管理员账号登录AcuHMI-1-7网页
# 测试步骤：
#   1. Protocols→Modbus→Device Mirror，点击Enable
#   2. 选择可勾选Enable的设备行，Slave ID设置为1（低于最小值2），点击Save
#   3. Slave ID设置为100（高于最大值99），点击Save
#   4. Slave ID设置为边界最小值2，点击Save
#   5. Slave ID设置为边界最大值99，点击Save
# 预期结果：
#   2. 保存失败，输入框有提示信息
#   3. 保存失败，输入框有提示信息
#   4. 保存成功
#   5. 保存成功
def test_TestCase_AcuHMI_008_04_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 2: Slave ID=1（低于最小值2），应保存失败且有提示信息
    _nav_device_mirror(page)
    _set_row_slave_id(page, 1)
    _click_save(page)
    assert _is_rejected(page), \
        "Slave ID=1 低于最小值2，保存应失败，但系统接受了该值"
    tip = _get_field_error_text(page)
    assert len(tip) > 0, \
        "Slave ID=1 保存失败后，输入框下方应有提示信息，但未检测到"
    print(f"\nSlave ID=1 提示信息：{tip}")

    # Step 3: Slave ID=100（高于最大值99），应保存失败且有提示信息
    _nav_device_mirror(page)
    _set_row_slave_id(page, 100)
    _click_save(page)
    assert _is_rejected(page), \
        "Slave ID=100 高于最大值99，保存应失败，但系统接受了该值"
    tip = _get_field_error_text(page)
    assert len(tip) > 0, \
        "Slave ID=100 保存失败后，输入框下方应有提示信息，但未检测到"
    print(f"Slave ID=100 提示信息：{tip}")

    # Step 4: Slave ID=2（边界最小值），应保存成功
    _nav_device_mirror(page)
    _set_row_slave_id(page, 2)
    _click_save(page)
    assert _is_saved_successfully(page), \
        "Slave ID=2 为边界最小值，保存应成功，但出现错误"

    # Step 5: Slave ID=99（边界最大值），应保存成功
    _nav_device_mirror(page)
    _set_row_slave_id(page, 99)
    _click_save(page)
    assert _is_saved_successfully(page), \
        "Slave ID=99 为边界最大值，保存应成功，但出现错误"
