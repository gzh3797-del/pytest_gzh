import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# Name仅允许字母、数字、下划线及空格，以下包含非法特殊字符
_INVALID_NAME = "Test!@#$%"


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case08
# 用例标题：Name字段输入不合法特殊字符，保存失败并提示格式错误
# 预置条件：1.管理权限登录AcuHMI网页
# 测试步骤：
#   1. 进入About页面
#   2. 在Name字段输入包含不合法特殊字符的内容（如"Test!@#$%"）
#   3. 点击Save按钮
# 预期结果：
#   1. 保存失败
#   2. 页面提示Name格式不符合要求（Name仅允许字母、数字、下划线及空格）
def test_TestCase_AcuHMI_009_08_case08(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 输入含非法特殊字符的Name
    name_field = page.get_by_placeholder("Enter Name")
    name_field.clear()
    name_field.fill(_INVALID_NAME)
    # 触发失焦以激活实时验证
    page.get_by_placeholder("Enter Location").click()
    page.wait_for_timeout(200)

    # 点击Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)

    # 验证保存失败：显示格式错误提示
    assert page.locator(".el-form-item__error").count() > 0 or \
           page.locator(".el-message--error").count() > 0, \
        f"Name含非法字符'{_INVALID_NAME}'时应显示格式错误提示，保存不成功"

    # 确认没有出现保存成功提示（成功 toast 类名 el-message--success，文本 "Device info saved"）
    assert page.locator(".el-message--success").count() == 0, \
        "Name含非法字符时不应出现保存成功提示"
