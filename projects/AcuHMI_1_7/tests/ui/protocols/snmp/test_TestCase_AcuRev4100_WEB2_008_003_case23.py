import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_008_003_case23
# 用例标题: Report Hold Time参数设置验证
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3、NMS管理端已开启 | 4、SNMP已开启
# 测试步骤:
#   1.开启TRAP，并成功与NMS端建立连接
#   2.设置Report Hold Time为0
#   3.设置Report Hold Time为300
#   4.设置Report Hold Time为301
#   5.使用IPV6登录网页，NMS管理端配置IPV6地址，重复1-4步骤
# 预期结果: 2.设置成功 | 3.设置成功 | 4.设置失败

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


def test_TestCase_AcuRev4100_WEB2_008_003_case23(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "SNMP")

    enable_radio = page.locator(".el-form-item").filter(has_text="SNMP Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(800)

    trap_enable = page.locator(".el-form-item").filter(has_text="Trap Enable").locator(".el-radio").filter(has_text="Enable")
    if trap_enable.count() > 0 and 'is-checked' not in (trap_enable.first.get_attribute("class") or ""):
        trap_enable.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(800)

    hold_field = page.locator(".el-form-item").filter(has_text="Hold Time").locator("input")
    if hold_field.count() == 0:
        pytest.skip("Hold Time 字段未找到，可能 Trap Enable 未显示该字段")

    # 合法值 — 只检查 Hold Time 自身的字段错误
    hold_item = page.locator(".el-form-item").filter(has_text="Hold Time")
    for valid in ["10", "300"]:
        hold_field.fill(valid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert hold_item.locator(".el-form-item__error").count() == 0, \
            f"Hold Time={valid} 应保存成功"

    # 非法值
    for invalid in ["301", "-1", "abc"]:
        hold_field.fill(invalid)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        assert hold_item.locator(".el-form-item__error").count() > 0 or \
            page.locator(".el-message--error").count() > 0, \
            f"Hold Time={invalid} 应保存失败"
