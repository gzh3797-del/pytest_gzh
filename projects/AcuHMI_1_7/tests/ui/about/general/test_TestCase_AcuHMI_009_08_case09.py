import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 恰好40个合法字符（上边界值）
_NAME_40 = "a" * 40


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case09
# 用例标题：Name字段输入恰好40个字符（上边界值），保存成功
# 预置条件：1.管理权限登录AcuHMI网页
# 测试步骤：
#   1. 进入About页面
#   2. 在Name字段输入恰好40个合法字符（如40个字母"a"）
#   3. 点击Save按钮
# 预期结果：
#   1. 保存成功，恰好等于最大长度40字符时允许保存
def test_TestCase_AcuHMI_009_08_case09(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 输入恰好40个字符
    name_field = page.get_by_placeholder("Enter Name")
    name_field.clear()
    name_field.fill(_NAME_40)

    # 点击Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # 验证保存成功：出现success提示，页面无错误提示
    assert page.get_by_text("success", exact=False).is_visible() or \
           page.locator(".el-message--success").count() > 0, \
        f"Name输入恰好40字符（上边界值）时应保存成功"
    assert page.locator(".el-form-item__error").count() == 0, \
        "Name输入恰好40字符时页面不应显示验证错误"
