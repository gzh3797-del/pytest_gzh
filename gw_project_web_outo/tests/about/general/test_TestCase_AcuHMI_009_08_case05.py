import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_VALID_NAME = "TestDevice_01"
_VALID_LOCATION = "Room 101"
_VALID_DESCRIPTION = "Test description"


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case05
# 用例标题：填写有效的Name、Location、Description并保存，保存成功且有提示反馈
# 预置条件：1.管理权限登录AcuHMI网页
# 测试步骤：
#   1. 进入About页面
#   2. 在Name字段输入有效名称（字母/数字/下划线/空格，≤40字符，如"TestDevice_01"）
#   3. 在Location字段输入有效位置（≤40字符，如"Room 101"）
#   4. 在Description字段输入有效描述（≤40字符，如"Test description"）
#   5. 点击Save按钮
# 预期结果：
#   1. 点击Save后页面出现保存成功提示（Toast消息）
#   2. 刷新页面后Name、Location、Description字段显示已保存的内容，信息未丢失
def test_TestCase_AcuHMI_009_08_case05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 填写有效的Name、Location、Description
    page.get_by_placeholder("Enter Name").clear()
    page.get_by_placeholder("Enter Name").fill(_VALID_NAME)
    page.get_by_placeholder("Enter Location").clear()
    page.get_by_placeholder("Enter Location").fill(_VALID_LOCATION)
    page.get_by_placeholder("Enter Description").clear()
    page.get_by_placeholder("Enter Description").fill(_VALID_DESCRIPTION)

    # 点击Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # 验证保存成功Toast
    assert page.get_by_text("success", exact=False).is_visible() or \
           page.locator(".el-message--success").count() > 0, \
        "保存后应出现success提示Toast"

    # 刷新页面验证数据持久化
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    _nav_to_about(page)

    assert page.get_by_placeholder("Enter Name").input_value() == _VALID_NAME, \
        f"刷新后Name字段应显示已保存的值 '{_VALID_NAME}'"
    assert page.get_by_placeholder("Enter Location").input_value() == _VALID_LOCATION, \
        f"刷新后Location字段应显示已保存的值 '{_VALID_LOCATION}'"
    assert page.get_by_placeholder("Enter Description").input_value() == _VALID_DESCRIPTION, \
        f"刷新后Description字段应显示已保存的值 '{_VALID_DESCRIPTION}'"
