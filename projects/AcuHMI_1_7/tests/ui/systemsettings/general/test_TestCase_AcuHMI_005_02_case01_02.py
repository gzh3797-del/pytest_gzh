# 用例编号: TestCase_AcuHMI_005_02_case01_02
# 用例标题: 以太网1 IP地址为非法值，Mask/Gateway/DNS为非法值，保存配置失败，提示错误信息准确
# 预置条件: 1、服务启动正常，账号登录成功
# 测试步骤:
#   1. IP地址=192.168.1.a，Mask=255.255.255.k，Gateway=192.168.1.l，DNS=8.8.a.w，Save → 失败
#   2. IP=192.168.*.123，Mask=255.255.255.*，Gateway=192.168.1.#，DNS=8.8.8.@，Save → 失败
#   3. IP=192.168.1123，Mask=255.255.2550，Gateway=192.168.11，DNS=8.8.83，Save → 失败
# 预期结果: 所有非法IP/Mask/Gateway/DNS值保存失败，显示错误信息

import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_settings_tab(page, tab: str):
    if "/systemSettings" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=tab).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def test_TestCase_AcuHMI_005_02_case01_02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Network")

    # 切换以太网1为手动/Static模式
    try:
        page.locator(".el-radio").filter(has_text="Static").first.click()
        page.wait_for_timeout(500)
    except Exception:
        try:
            page.locator(".el-radio").filter(has_text="Manual").first.click()
            page.wait_for_timeout(500)
        except Exception:
            pass

    # 三组非法值测试
    invalid_sets = [
        ("192.168.1.a", "255.255.255.k", "192.168.1.l", "8.8.a.w"),
        ("192.168.*.123", "255.255.255.*", "192.168.1.#", "8.8.8.@"),
        ("192.168.1123", "255.255.2550", "192.168.11", "8.8.83"),
    ]
    for ip, mask, gw, dns in invalid_sets:
        page.get_by_placeholder("Enter IP").first.fill(ip)
        page.get_by_placeholder("Enter Mask").first.fill(mask)
        page.get_by_placeholder("Enter Gateway").first.fill(gw)
        page.get_by_placeholder("Enter DNS 1").first.fill(dns)
        page.get_by_role("button", name="Save").click()
        page.wait_for_timeout(500)
        # EP 表单校验错误异步渲染，先 auto-wait 再断言，避免即时 count 取到 0 误判
        page.locator(".el-form-item__error").first.wait_for(state="visible", timeout=5000)
        assert page.locator(".el-form-item__error").count() > 0, \
            f"非法IP组合 ({ip}, {mask}, {gw}, {dns}) 应保存失败并显示错误信息"
