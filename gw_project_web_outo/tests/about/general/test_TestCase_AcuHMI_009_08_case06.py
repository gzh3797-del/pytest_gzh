import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_SAVED_NAME = "PersistDevice_06"
_SAVED_LOCATION = "Lab 06"
_SAVED_DESCRIPTION = "Persist test desc"


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case06
# 用例标题：保存Device Information后刷新页面，配置信息持久化保存
# 预置条件：
#   1. 管理权限登录AcuHMI网页
#   2. 已在About页面成功保存Name/Location/Description
# 测试步骤：
#   1. 在About页面成功保存Name、Location、Description（参考case05）
#   2. 刷新浏览器页面（按F5或重新加载）
#   3. 重新进入About页面，查看各字段内容
# 预期结果：
#   1. 刷新后Name、Location、Description字段显示上次保存的内容
#   2. 内容未被清空或恢复为默认值
def test_TestCase_AcuHMI_009_08_case06(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 保存一组有效数据
    page.get_by_placeholder("Enter Name").clear()
    page.get_by_placeholder("Enter Name").fill(_SAVED_NAME)
    page.get_by_placeholder("Enter Location").clear()
    page.get_by_placeholder("Enter Location").fill(_SAVED_LOCATION)
    page.get_by_placeholder("Enter Description").clear()
    page.get_by_placeholder("Enter Description").fill(_SAVED_DESCRIPTION)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)

    # 确认保存成功
    assert page.get_by_text("success", exact=False).is_visible() or \
           page.locator(".el-message--success").count() > 0, \
        "保存后应出现success提示"

    # 刷新页面（模拟F5）
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # 重新导航到About页面
    _nav_to_about(page)

    # 验证字段值持久化
    assert page.get_by_placeholder("Enter Name").input_value() == _SAVED_NAME, \
        f"刷新后Name字段应持久化显示 '{_SAVED_NAME}'"
    assert page.get_by_placeholder("Enter Location").input_value() == _SAVED_LOCATION, \
        f"刷新后Location字段应持久化显示 '{_SAVED_LOCATION}'"
    assert page.get_by_placeholder("Enter Description").input_value() == _SAVED_DESCRIPTION, \
        f"刷新后Description字段应持久化显示 '{_SAVED_DESCRIPTION}'"
