import pytest
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_009
# 用例标题: Time To Live 合法边界值保存正常
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 启用 BACnet Enable，再启用 Foreign Device Function
#   2. 填写必填字段（Device Object Name、Devices Selection、BBMD IP、BBMD Port）
#   3. 输入 Time To Live=5 并保存 → 预期成功
#   4. 输入 Time To Live=1440 并保存 → 预期成功
# 预期结果:
#   3. TTL 下边界值 5 保存成功（无字段校验错误）
#   4. TTL 上边界值 1440 保存成功（无字段校验错误）


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
    """Enable BACnet + FDF, fill all required fields."""
    # Enable BACnet Enable
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    bacnet_enable_radio = bacnet_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (bacnet_enable_radio.get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Device Object Name（必填，空则无法保存）
    obj_name_inp = page.locator(".el-form-item").filter(has_text="Device Object Name").first.locator("input").first
    if not obj_name_inp.input_value():
        obj_name_inp.fill("TestGW")
        page.wait_for_timeout(100)

    # Enable Foreign Device Function
    fdf_item = page.locator(".el-form-item").filter(has_text="Foreign Device Function").first
    fdf_enable_radio = fdf_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (fdf_enable_radio.get_attribute("class") or ""):
        fdf_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # BBMD IP（必填）
    bbmd_ip_inp = page.locator(".el-form-item").filter(has_text="BBMD IP").first.locator("input").first
    bbmd_ip_inp.fill("192.168.2.100")
    page.wait_for_timeout(100)

    # BBMD Port（必填，范围 47808-49000）
    bbmd_port_inp = page.locator(".el-form-item").filter(has_text="BBMD Port").first.locator("input").first
    bbmd_port_inp.fill("47809")
    page.wait_for_timeout(100)

    # Devices Selection To Mapping：勾选第一个设备行（tbody 第一行 checkbox）
    device_table = page.locator(".el-form-item").filter(
        has_text="Devices Selection To Mapping"
    )
    if device_table.count() == 0:
        # 该 form-item 用 div 作为 label，fallback：找含 is-required 且有 checkbox 的 form-item
        for fi in page.locator(".el-form-item.is-required").all():
            if fi.locator(".el-checkbox").count() > 3:
                device_table = fi
                break

    if device_table.count() > 0:
        tbody_rows = device_table.locator("tbody tr").all()
        if tbody_rows:
            first_row_cb = tbody_rows[0].locator(".el-checkbox").first
            if not tbody_rows[0].locator("input[type='checkbox']").first.is_checked():
                first_row_cb.click()
                page.wait_for_timeout(300)


def _save_and_check(page, ttl_value: str):
    """Fill Time To Live, click Save, assert no validation errors."""
    ttl_input = page.locator(".el-form-item").filter(has_text="Time To Live").first.locator("input").first
    ttl_input.fill(ttl_value)
    page.wait_for_timeout(200)

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)

    field_errors = page.locator(".el-form-item__error").count()
    msg_errors = page.locator(".el-message--error").count()
    assert field_errors == 0 and msg_errors == 0, \
        f"Time To Live={ttl_value} 合法值应保存成功，但出现错误（field_errors={field_errors}, msg_errors={msg_errors}）"


def test_TestCase_WEB2_033_001_009(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    # Step 1: 导航并配置表单（BACnet Enable + FDF + 所有必填字段）
    _nav_protocol(page, "BACnet/IP")
    _setup_bacnet_form(page)

    assert page.locator(".el-form-item").filter(has_text="Time To Live").count() > 0, \
        "启用 FDF 后应显示 Time To Live 字段"

    # Step 2: TTL 下边界值 5
    _save_and_check(page, "5")

    # 重新进入页面并配置（Save 后页面刷新）
    _nav_protocol(page, "BACnet/IP")
    _setup_bacnet_form(page)

    # Step 3: TTL 上边界值 1440
    _save_and_check(page, "1440")
