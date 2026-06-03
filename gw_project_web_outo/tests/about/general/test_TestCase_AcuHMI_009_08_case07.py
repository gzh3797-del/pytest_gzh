import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case07
# 用例标题：Name字段为空时点击Save，保存失败并提示必填
# 预置条件：1.管理权限登录AcuHMI网页
# 测试步骤：
#   1. 进入About页面
#   2. 清空Name输入字段（删除已有内容使输入框为空）
#   3. 点击Save按钮
# 预期结果：
#   1. 保存失败
#   2. 页面提示Name字段不能为空（Name为必填项）
def test_TestCase_AcuHMI_009_08_case07(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 清空Name字段
    name_field = page.get_by_placeholder("Enter Name")
    name_field.clear()
    # 触发失焦以激活验证
    page.get_by_placeholder("Enter Location").click()
    page.wait_for_timeout(200)

    # 点击Save
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(500)

    # 验证保存失败：显示必填错误提示
    assert page.locator(".el-form-item__error").count() > 0 or \
           page.locator(".el-message--error").count() > 0, \
        "Name为空时点击Save应显示必填验证错误提示，保存不成功"

    # 确认没有出现success提示
    assert not page.get_by_text("success", exact=False).is_visible(), \
        "Name为空时不应出现success提示"
