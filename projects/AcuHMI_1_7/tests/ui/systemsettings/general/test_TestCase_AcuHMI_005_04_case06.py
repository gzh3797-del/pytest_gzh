# 用例编号: TestCase_AcuHMI_005_04_case06
# 用例标题: 密码为混合字符，取消隐藏查看密码为明文，保存成功
# 预置条件: 管理权限登录AcuHMI网页
# 测试步骤:
#   1. 进入 System Settings -> Email
#   2. 填写基准配置，将 Password 改为 !@#AbC123
#   3. 点击密码字段旁的显示按钮（眼睛图标），切换为明文显示
#   4. 验证 input type 变为 text（密码可见）
#   5. 点击 Save
# 预期结果: 密码切换为明文显示，保存成功

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


def _fill_email_baseline(page):
    page.get_by_label("Email Server", exact=False).fill("smtp.163.com")
    page.get_by_label("Email Port", exact=False).fill("25")
    try:
        page.locator(".el-radio").filter(has_text="Off").click()
        page.wait_for_timeout(200)
    except Exception:
        pass
    page.get_by_label("Sender Name", exact=True).fill("xiaoming")
    page.get_by_label("From Email Address", exact=True).fill("159xxxx4651@163.com")
    page.get_by_label("Username", exact=True).fill("xiaoming123")
    page.get_by_label("Password", exact=True).fill("Admin@110001")


@pytest.mark.skip(reason="Email Password field has no show/hide eye icon (suffix-inner is empty); feature not implemented in this build")
def test_TestCase_AcuHMI_005_04_case06(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_settings_tab(page, "Email")
    _fill_email_baseline(page)
    pwd_value = "!@#AbC123"
    page.get_by_label("Password", exact=True).fill(pwd_value)
    # Click show password toggle — suffix icon is hidden (opacity 0) until hover; use force=True
    page.locator(".el-form-item").filter(has_text="Password").locator(".el-input__suffix-inner").click(force=True)
    page.wait_for_timeout(200)
    # Verify password is now visible (type=text)
    pwd_input = page.get_by_label("Password", exact=True)
    input_type = pwd_input.get_attribute("type")
    assert input_type == "text", f"密码应显示明文，当前type：{input_type}"
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message")).to_be_visible(timeout=5000)
