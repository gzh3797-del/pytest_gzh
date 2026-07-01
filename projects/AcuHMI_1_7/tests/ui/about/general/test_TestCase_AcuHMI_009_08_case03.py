import pytest
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_009_08_case03
# 用例标题：设备描述大于40个字符，保存配置失败，小于等于40个字符，保存配置成功
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. About->Device information设置Description超过40个字符，点击Save
#   2. 设置Description等于40个字符，点击Save
# 预期结果：
#   1. 弹框提示"Description cannot exceed 40 characters"
#   2. 弹框提示"Device info saved"（success Toast）
# 说明：设备前端有脏检查——保存与已保存值相同的内容会弹 warning "No change to save" 而非成功提示，
#       故 Step 2 取与当前已保存值不同的等长(40字符)目标值，确保触发真实保存。
def test_TestCase_AcuHMI_009_08_case03(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_about(page)

    # 进入页面时输入框回显的是设备已保存值，据此选取与之不同的等长目标值（规避前端脏检查）
    desc_field = page.get_by_placeholder("Enter Description")
    baseline = desc_field.input_value()

    # Step 1: 输入超过40个字符的Description，点击Save，验证弹框提示"Description cannot exceed 40 characters"
    desc_field.clear()
    desc_field.fill("c" * 41)
    page.get_by_role("button", name="Save").click()
    expect(page.get_by_text("Description cannot exceed 40 characters", exact=False)).to_be_visible(timeout=3000)

    # Step 2: 输入恰好等于40个字符且与已保存值不同的Description，点击Save，验证成功提示"Device info saved"
    target = "c" * 40 if baseline != "c" * 40 else "z" * 40
    desc_field.clear()
    desc_field.fill(target)
    page.get_by_role("button", name="Save").click()
    expect(page.locator(".el-message--success")).to_contain_text("Device info saved", timeout=8000)
