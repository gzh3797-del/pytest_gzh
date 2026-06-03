import pytest
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_011
# 用例标题: 非法 BBMD 配置保存被阻止
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 启用 BACnet Enable + Foreign Device Function
#   2. 分别输入非法值：BBMD IP=300.1.1.1、BBMD Port=70000、Time To Live=0
#   3. 每次点击 Save
# 预期结果: 每次保存均被阻止，页面显示错误提示


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


def _setup_bacnet_form(page):
    """Enable BACnet + FDF, fill required fields with valid base values."""
    # Enable BACnet
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    if "is-checked" not in (bacnet_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Device Object Name（必填）
    obj_name_inp = page.locator(".el-form-item").filter(has_text="Device Object Name").first.locator("input").first
    if not obj_name_inp.input_value():
        obj_name_inp.fill("TestGW")
        page.wait_for_timeout(100)

    # Enable Foreign Device Function
    fdf_item = page.locator(".el-form-item").filter(has_text="Foreign Device Function").first
    if "is-checked" not in (fdf_item.locator(".el-radio").filter(has_text="Enable").get_attribute("class") or ""):
        fdf_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Devices Selection To Mapping：勾选第一行（必填）
    for fi in page.locator(".el-form-item.is-required").all():
        if fi.locator(".el-checkbox").count() > 3:
            tbody_rows = fi.locator("tbody tr").all()
            if tbody_rows:
                if not tbody_rows[0].locator("input[type='checkbox']").first.is_checked():
                    tbody_rows[0].locator(".el-checkbox").first.click()
                    page.wait_for_timeout(300)
            break


def _save_and_assert_error(page, label: str):
    """Click Save and assert at least one validation error appears."""
    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors > 0 or msg_errors > 0, \
        f"[{label}] 非法值应保存失败并出现错误提示，但未找到错误（field_errors={field_errors}, msg_errors={msg_errors}）"


def test_TestCase_WEB2_033_001_011(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # --- 非法 BBMD IP: 300.1.1.1 ---
    _nav_protocol(page, "BACnet/IP")
    _setup_bacnet_form(page)

    bbmd_ip_inp = page.locator(".el-form-item").filter(has_text="BBMD IP").first.locator("input").first
    bbmd_ip_inp.fill("300.1.1.1")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="BBMD Port").first.locator("input").first.fill("47809")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="Time To Live").first.locator("input").first.fill("60")
    page.wait_for_timeout(100)
    _save_and_assert_error(page, "BBMD IP=300.1.1.1")

    # --- 非法 BBMD Port: 70000（超出 47808-49000）---
    _nav_protocol(page, "BACnet/IP")
    _setup_bacnet_form(page)

    page.locator(".el-form-item").filter(has_text="BBMD IP").first.locator("input").first.fill("192.168.2.100")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="BBMD Port").first.locator("input").first.fill("70000")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="Time To Live").first.locator("input").first.fill("60")
    page.wait_for_timeout(100)
    _save_and_assert_error(page, "BBMD Port=70000")

    # --- 非法 Time To Live: 0（低于下边界 5）---
    _nav_protocol(page, "BACnet/IP")
    _setup_bacnet_form(page)

    page.locator(".el-form-item").filter(has_text="BBMD IP").first.locator("input").first.fill("192.168.2.100")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="BBMD Port").first.locator("input").first.fill("47809")
    page.wait_for_timeout(100)
    page.locator(".el-form-item").filter(has_text="Time To Live").first.locator("input").first.fill("0")
    page.wait_for_timeout(100)
    _save_and_assert_error(page, "Time To Live=0")
