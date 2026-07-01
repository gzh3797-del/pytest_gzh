import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 两组有效值：运行时挑与当前已保存值不同的一组，规避前端脏检查（相同值会弹 "No change to save"）
_VALID_SET_A = ("TestDevice_01", "Room 101", "Test description")
_VALID_SET_B = ("TestDevice_02", "Room 102", "Test description B")


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

    name_in = page.get_by_placeholder("Enter Name")
    loc_in = page.get_by_placeholder("Enter Location")
    desc_in = page.get_by_placeholder("Enter Description")

    # 选取与当前已保存值不同的一组值，确保 Save 触发真实保存（相同值会弹 "No change to save"）
    current = (name_in.input_value(), loc_in.input_value(), desc_in.input_value())
    name_v, loc_v, desc_v = _VALID_SET_B if current == _VALID_SET_A else _VALID_SET_A

    # 填写有效的Name、Location、Description
    name_in.clear()
    name_in.fill(name_v)
    loc_in.clear()
    loc_in.fill(loc_v)
    desc_in.clear()
    desc_in.fill(desc_v)

    # 点击Save
    page.get_by_role("button", name="Save").click()

    # 验证保存成功Toast："Device info saved"
    expect(page.locator(".el-message--success")).to_contain_text("Device info saved", timeout=8000)

    # 刷新页面验证数据持久化
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    _nav_to_about(page)

    assert page.get_by_placeholder("Enter Name").input_value() == name_v, \
        f"刷新后Name字段应显示已保存的值 '{name_v}'"
    assert page.get_by_placeholder("Enter Location").input_value() == loc_v, \
        f"刷新后Location字段应显示已保存的值 '{loc_v}'"
    assert page.get_by_placeholder("Enter Description").input_value() == desc_v, \
        f"刷新后Description字段应显示已保存的值 '{desc_v}'"
