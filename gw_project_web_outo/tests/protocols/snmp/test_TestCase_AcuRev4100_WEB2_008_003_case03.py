import pytest
import re
from playwright.sync_api import expect
from pages.login_page import LoginPage

# 用例编号: TestCase_AcuRev4100_WEB2_008_003_case03
# 用例标题: 开启SNMP，配置Version为SNMPv2C、端口为16159、RO Community与管理端一致
# 预置条件: 1、服务启动正常，账号登录成功 | 2、环境存在多个设备 | 3、NMS管理端已开启
# 测试步骤:
#   1.登录Setting-Communications-SNMP页面
#   2.开启SNMP，配置Version为SNMPv2C、端口为16159，RO Community为@@@###，点击保存
#   3.点击Download MIB File，下载MIB文件
#   4.NMS管理端导入下载的MIB文件，配置端口和RO Community与WEB2一致
#   5.NMS管理端GET请求接入设备的数据
#   6.使用IPV6登录网页，NMS管理端配置IPV6地址，重复1-5步骤
# 预期结果: 2.参数保存成功 | 3.下载文件格式和内容正确，符合SNMP标准协议 | 5.请求成功，参数内容完整，值正确。支持的参数范围可参考需求

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


def test_TestCase_AcuRev4100_WEB2_008_003_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    _nav_protocol(page, "SNMP")

    # 开启 SNMP，配置 Version=v2c, Port=16159, Community='public'
    enable_radio = page.locator(".el-form-item").filter(has_text="SNMP Enable").locator(".el-radio").filter(has_text="Enable")
    if enable_radio.count() > 0 and 'is-checked' not in (enable_radio.first.get_attribute("class") or ""):
        enable_radio.first.locator(".el-radio__inner").click()
        page.wait_for_timeout(500)

    port_field = page.get_by_label("Port", exact=False).or_(
        page.get_by_placeholder("Enter Port"))

    # Select SNMP Version v2c
    version_select = page.locator(".el-form-item").filter(has_text="Version").locator(".el-select")
    if version_select.count() > 0:
        version_select.click()
        page.wait_for_timeout(200)
        v2c_opt = page.get_by_role("option").filter(has_text=re.compile(r"v2c", re.IGNORECASE))
        if v2c_opt.count() > 0:
            v2c_opt.first.click()
        page.wait_for_timeout(200)
    port_field.fill("16159")

    community_field = page.get_by_placeholder("Enter RO Community").or_(
        page.get_by_label("RO Community", exact=False))
    community_field.fill("@@@###")

    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    assert page.locator(".el-message--error").count() == 0, "保存不应出现错误提示"
    # TODO: 使用 pysnmp 从 NMS 端验证 SNMP MIB 数据成功
