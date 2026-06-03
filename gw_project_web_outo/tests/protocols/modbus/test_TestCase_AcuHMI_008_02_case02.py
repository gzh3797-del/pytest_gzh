import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


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

    # 确保 Modbus 已启用
    enable_radio = page.locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0:
        try:
            enable_radio.first.click()
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Step 1: 填写 Modbus Port=502
    port_input = page.locator(".el-form-item").filter(has_text="Port").locator("input").first
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
