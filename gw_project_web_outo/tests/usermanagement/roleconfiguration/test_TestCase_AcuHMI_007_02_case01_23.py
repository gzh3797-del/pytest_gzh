import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_VALID_NAME_40 = "A" * 40
_OVER_NAME_41  = "A" * 41


def _nav_to_submenu(page, submenu: str):
    if "/userManagement/" not in page.url:
        page.locator("header span").filter(has_text="AcuHMI").first.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page.get_by_text("User Management").first.click()
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _delete_role_if_exists(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


# 用例编号：TestCase_AcuHMI_007_02_case01_23
# 用例标题：添加角色，角色名称验证
# 预置条件：服务启动正常，账号登录成功，进入用户设置界面
# 测试步骤：
#   1. 角色名称输入特殊字符（如 @!#$%）
#   2. 角色名称输入 40 个字符，保存配置
#   3. 角色名称输入 41 个字符，保存配置
# 预期结果：
#   1. 仅支持输入数字、字母、下划线、空格（特殊字符被过滤或保存失败）
#   2. 保存成功
#   3. 保存失败，提示角色名称长度限制为 40
def test_TestCase_AcuHMI_007_02_case01_23(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_submenu(page, "Role Configuration")

    # Step 1: 特殊字符保存失败（对话框保持打开 或 出现验证错误）
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    name_input = page.get_by_placeholder("Enter Role Name")
    name_input.click()
    page.keyboard.type("@!#$%^")
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(800)
    # 验证：对话框未关闭 或 出现 el-form-item__error
    dialog_still_open = page.get_by_placeholder("Enter Role Name").is_visible()
    error_visible = page.locator(".el-form-item__error").is_visible()
    assert dialog_still_open or error_visible, \
        "特殊字符角色名称应保存失败（对话框保持打开或出现验证错误）"
    page.get_by_role("button", name="Cancel").click()
    page.wait_for_timeout(500)

    # Step 2: 40 字符名称保存成功
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_VALID_NAME_40)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    # 对话框应已关闭（保存成功）
    assert not page.get_by_placeholder("Enter Role Name").is_visible(), \
        f"40 字符角色名称应保存成功（对话框关闭），当前对话框仍可见"

    # Step 3: 41 字符名称保存失败
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(_OVER_NAME_41)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)
    # 保存失败：对话框仍开着 或 有错误提示
    error_visible = page.locator(".el-form-item__error").is_visible()
    dialog_still_open = page.get_by_placeholder("Enter Role Name").is_visible()
    assert error_visible or dialog_still_open, \
        "41 字符角色名称应保存失败（出现错误提示或对话框保持开着）"
    # 关闭对话框
    try:
        page.get_by_role("button", name="Cancel").click(timeout=2000)
        page.wait_for_timeout(500)
    except Exception:
        pass

    # 清理：删除 40 字符名称的角色
    _delete_role_if_exists(page, _VALID_NAME_40)
