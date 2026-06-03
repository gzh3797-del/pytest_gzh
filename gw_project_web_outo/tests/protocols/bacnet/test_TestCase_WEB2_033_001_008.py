import pytest
from pages.login_page import LoginPage

# 用例编号: TestCase_WEB2_033_001_008
# 用例标题: Enable Foreign Device Function 开关控制字段可配
# 预置条件: web2已正常上电，相关服务正常启动。
# 测试步骤:
#   1. 打开 BACnet/IP 页面，启用 BACnet，查看 Foreign Device Function 字段
#   2. 启用 Foreign Device Function → 验证 BBMD IP / BBMD Port / Time To Live 可编辑
#   3. 关闭 Foreign Device Function → 验证上述字段消失（不可配置）
# 预期结果:
#   2. 启用后三个字段出现且可编辑
#   3. 关闭后三个字段消失


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


def test_TestCase_WEB2_033_001_008(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_protocol(page, "BACnet/IP")

    # 确保 BACnet Enable 已选中（若未选中则点击 Enable）
    bacnet_item = page.locator(".el-form-item").filter(has_text="BACnet Enable").first
    bacnet_enable_radio = bacnet_item.locator(".el-radio").filter(has_text="Enable")
    if "is-checked" not in (bacnet_enable_radio.get_attribute("class") or ""):
        bacnet_item.locator(".el-radio__label").filter(has_text="Enable").click()
        page.wait_for_timeout(500)

    # Step 1: 验证 Foreign Device Function 字段可见
    fdf_item = page.locator(".el-form-item").filter(has_text="Foreign Device Function").first
    assert fdf_item.count() > 0, \
        "BACnet 启用后应可看到 Foreign Device Function 字段"

    # Step 2: 启用 Foreign Device Function
    fdf_item.locator(".el-radio__label").filter(has_text="Enable").click()
    page.wait_for_timeout(500)

    bbmd_ip_item = page.locator(".el-form-item").filter(has_text="BBMD IP")
    bbmd_port_item = page.locator(".el-form-item").filter(has_text="BBMD Port")
    ttl_item = page.locator(".el-form-item").filter(has_text="Time To Live")

    assert bbmd_ip_item.count() > 0, "启用 Foreign Device Function 后应显示 BBMD IP 字段"
    assert bbmd_port_item.count() > 0, "启用 Foreign Device Function 后应显示 BBMD Port 字段"
    assert ttl_item.count() > 0, "启用 Foreign Device Function 后应显示 Time To Live 字段"

    assert bbmd_ip_item.first.locator("input").first.is_enabled(), \
        "BBMD IP 字段应可编辑"
    assert bbmd_port_item.first.locator("input").first.is_enabled(), \
        "BBMD Port 字段应可编辑"
    assert ttl_item.first.locator("input").first.is_enabled(), \
        "Time To Live 字段应可编辑"

    # Step 3: 关闭 Foreign Device Function → 字段消失
    fdf_item.locator(".el-radio__label").filter(has_text="Disable").click()
    page.wait_for_timeout(500)

    assert page.locator(".el-form-item").filter(has_text="BBMD IP").count() == 0, \
        "禁用 Foreign Device Function 后 BBMD IP 字段应消失"
    assert page.locator(".el-form-item").filter(has_text="BBMD Port").count() == 0, \
        "禁用 Foreign Device Function 后 BBMD Port 字段应消失"
    assert page.locator(".el-form-item").filter(has_text="Time To Live").count() == 0, \
        "禁用 Foreign Device Function 后 Time To Live 字段应消失"
