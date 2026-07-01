import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 两组有效值：运行时挑与当前已保存值不同的一组，规避前端脏检查（相同值会弹 "No change to save"）
_SAVED_SET_A = ("PersistDevice_06", "Lab 06", "Persist test desc")
_SAVED_SET_B = ("PersistDevice_06B", "Lab 06B", "Persist test desc B")


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

    name_in = page.get_by_placeholder("Enter Name")
    loc_in = page.get_by_placeholder("Enter Location")
    desc_in = page.get_by_placeholder("Enter Description")

    # 选取与当前已保存值不同的一组值，确保 Save 触发真实保存（相同值会弹 "No change to save"）
    current = (name_in.input_value(), loc_in.input_value(), desc_in.input_value())
    name_v, loc_v, desc_v = _SAVED_SET_B if current == _SAVED_SET_A else _SAVED_SET_A

    # 保存一组有效数据
    name_in.clear()
    name_in.fill(name_v)
    loc_in.clear()
    loc_in.fill(loc_v)
    desc_in.clear()
    desc_in.fill(desc_v)
    page.get_by_role("button", name="Save").click()

    # 确认保存成功："Device info saved"
    expect(page.locator(".el-message--success")).to_contain_text("Device info saved", timeout=8000)

    # 刷新页面（模拟F5）
    page.reload()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)

    # 重新导航到About页面
    _nav_to_about(page)

    # 验证字段值持久化
    assert page.get_by_placeholder("Enter Name").input_value() == name_v, \
        f"刷新后Name字段应持久化显示 '{name_v}'"
    assert page.get_by_placeholder("Enter Location").input_value() == loc_v, \
        f"刷新后Location字段应持久化显示 '{loc_v}'"
    assert page.get_by_placeholder("Enter Description").input_value() == desc_v, \
        f"刷新后Description字段应持久化显示 '{desc_v}'"
