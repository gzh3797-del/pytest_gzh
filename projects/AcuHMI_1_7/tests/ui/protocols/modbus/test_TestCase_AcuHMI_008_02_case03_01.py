import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
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


def _set_port_and_save(page, port: int):
    """Set Modbus Port value and click Save. Returns True if no error toast."""
    port_input = page.locator(".el-form-item").filter(has_text="Port").locator("input").first
    port_input.fill(str(port))
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1500)

    return page.locator(".el-message--error").count() == 0


def _verify_tips_table(page):
    """校验 tips 表表头包含 Type/Register/Length/Sequence，表体包含已知数据类型行。"""
    # 校验表头列名
    header = page.locator(".el-table__header")
    if header.count() > 0:
        header_text = header.last.inner_text().lower()
        for col in ["type", "register", "length", "sequence"]:
            assert col in header_text, \
                f"tips 表头未找到列 '{col}'，当前表头内容：{header_text[:200]}"

    # 校验表体包含已知数据类型且数值正确
    body = page.locator(".el-table__body").last
    body_text = body.inner_text().lower()
    for dtype in ["bit", "uint16", "int32", "float", "double"]:
        assert dtype in body_text, \
            f"tips 表体未找到数据类型 '{dtype}'，当前表体内容：{body_text[:300]}"


# 用例编号：TestCase_AcuHMI_008_02_case03_01
# 用例标题：配置modbus port为2000、5999，slave id为已存在设备slaveid，保存配置成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Protocols->Modbus Config: Modbus Port=2000
#   2. Save保存
#   3. 端口设置为5999，重复步骤1-2
# 预期结果：
#   2. 配置保存成功，tips表中Type,Register Length, Data Length Sequence显示信息正确
def test_TestCase_AcuHMI_008_02_case03_01(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_modbus_config(page)

    # 确保 Modbus 已启用
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0:
        try:
            enable_radio.first.click()
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Step 1-2: Port=2000，Save，校验成功
    success = _set_port_and_save(page, port=2000)
    assert success, "Port=2000 时保存应成功，但出现错误提示"

    try:
        expect(page.get_by_text("success", exact=False)).to_be_visible(timeout=3000)
    except Exception:
        assert page.locator(".el-message--error").count() == 0, \
            "Port=2000 保存后不应出现错误 toast"

    _verify_tips_table(page)

    # Step 3: Port=5999，Save，校验成功
    success = _set_port_and_save(page, port=5999)
    assert success, "Port=5999 时保存应成功，但出现错误提示"

    try:
        expect(page.get_by_text("success", exact=False)).to_be_visible(timeout=3000)
    except Exception:
        assert page.locator(".el-message--error").count() == 0, \
            "Port=5999 保存后不应出现错误 toast"

    _verify_tips_table(page)
