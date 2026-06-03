import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

_TEST_USER   = "pwdchg05"
_CUSTOM_ROLE = "editrole05"
_INIT_PWD    = "Admin@110001"
_NEW_PWD     = "Admin@Changed5"


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


def _create_edit_role(page, role_name: str):
    """Create a custom role with User=edit permission."""
    _nav_to_submenu(page, "Role Configuration")
    page.get_by_role("button", name="Add Role").click()
    page.wait_for_timeout(1000)
    page.get_by_placeholder("Enter Role Name").fill(role_name)
    page.locator(".el-form-item").filter(has_text="User").locator(".el-select").click()
    page.wait_for_timeout(300)
    page.get_by_role("option", name="edit").click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_role(page, role_name: str):
    _nav_to_submenu(page, "Role Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=role_name)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _create_user(page, username: str, password: str, role: str):
    _nav_to_submenu(page, "User Configuration")
    page.get_by_role("button", name="Add User").click()
    page.wait_for_timeout(1000)
    page.get_by_label("Username", exact=True).fill(username)
    page.get_by_label("Password", exact=True).fill(password)
    page.get_by_label("Repeat Password", exact=True).fill(password)
    page.get_by_text("--Select Role--", exact=True).click()
    page.get_by_role("option", name=role).click()
    page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    page.wait_for_timeout(1000)


def _delete_user(page, username: str):
    _nav_to_submenu(page, "User Configuration")
    row = page.locator("tbody").get_by_role("row").filter(has_text=username)
    if row.count() == 0:
        return
    row.get_by_role("button").last.click()
    page.get_by_role("button", name="Yes, continue").click()
    page.wait_for_timeout(500)


def _login_as_user(browser, username: str, password: str):
    """Open a new context, log in as the given user, return (ctx, page)."""
    ctx = browser.new_context(ignore_https_errors=True)
    p   = ctx.new_page()
    p.goto(BASE_URL + "/#/login")
    p.wait_for_load_state("networkidle")
    p.get_by_role("textbox", name="Enter User Name").fill(username)
    p.get_by_role("textbox", name="Enter Password").fill(password)
    p.get_by_role("button", name="Sign In").click()
    p.wait_for_load_state("networkidle")
    for btn in ["Accept", "I Accept", "确认"]:
        try:
            p.get_by_role("button", name=btn).click(timeout=2000)
            p.wait_for_load_state("networkidle")
        except Exception:
            pass
    try:
        p.get_by_role("button", name="Cancel").click(timeout=2000)
    except Exception:
        pass
    return ctx, p


def _can_login(browser, username: str, password: str) -> bool:
    ctx = browser.new_context(ignore_https_errors=True)
    p   = ctx.new_page()
    try:
        p.goto(BASE_URL + "/#/login")
        p.wait_for_load_state("networkidle")
        p.get_by_role("textbox", name="Enter User Name").fill(username)
        p.get_by_role("textbox", name="Enter Password").fill(password)
        p.get_by_role("button", name="Sign In").click()
        p.wait_for_load_state("networkidle")
        for btn in ["Accept", "I Accept", "确认"]:
            try:
                p.get_by_role("button", name=btn).click(timeout=2000)
                p.wait_for_load_state("networkidle")
            except Exception:
                pass
        try:
            p.get_by_role("button", name="Cancel").click(timeout=2000)
        except Exception:
            pass
        return "/#/login" not in p.url
    finally:
        ctx.close()


# 用例编号：TestCase_ARM_XXL_002_04_case1_05
# 用例标题：非admin用户修改自己的密码
# 预置条件：管理权限登录 AcuHMI 网页
# 测试步骤：
#   1. Role Configuration 新建 User=edit 权限的自定义角色 editrole05
#   2. User Configuration 新建该角色用户 pwdchg05
#   3. pwdchg05 登录系统
#   4. Password Management 页面选择 pwdchg05 进行编辑，输入新密码并保存
# 预期结果：
#   3. 登录成功，Password Management 行有编辑按钮
#   4. 修改成功（提示 "User password changed"），新密码可登录
# 注：Current Password 通过点击编辑后弹出的"Current User Password"弹框输入，
#     而非编辑表单的独立字段；表单字段为 Username/Password/Repeat Password
def test_TestCase_ARM_XXL_002_04_case1_05(login_page: LoginPage):
    login_page.open()
    login_page.login()
    admin_page = login_page.page
    browser    = admin_page.context.browser

    _create_edit_role(admin_page, _CUSTOM_ROLE)
    _create_user(admin_page, _TEST_USER, _INIT_PWD, role=_CUSTOM_ROLE)
    try:
        # 步骤 3：非 admin 用户登录
        ctx, user_page = _login_as_user(browser, _TEST_USER, _INIT_PWD)
        try:
            assert "/#/login" not in user_page.url, \
                f"用户 {_TEST_USER} 登录失败"

            # 步骤 4：进入 Password Management，找到自己的行并点击编辑
            _nav_to_submenu(user_page, "Password Management")
            row = user_page.locator("tbody").get_by_role("row").filter(has_text=_TEST_USER)
            assert row.get_by_role("button").count() > 0, \
                "非 admin（edit权限）在 Password Management 自己行应有编辑按钮"
            row.get_by_role("button").click()
            user_page.wait_for_timeout(500)

            # 验证"Current User Password"弹框出现（非 admin 修改密码时的身份确认）
            assert user_page.get_by_placeholder("Please input").is_visible(), \
                "非 admin 修改密码时，应弹出 Current User Password 确认弹框"

            # 步骤 4 续：填写"Current User Password"弹框后提交
            user_page.get_by_placeholder("Please input").fill(_INIT_PWD)
            user_page.get_by_role("button", name="Confirm").click()
            user_page.wait_for_timeout(500)

            # 弹框关闭后填写新密码并保存
            user_page.get_by_label("Password", exact=True).fill(_NEW_PWD)
            user_page.get_by_label("Repeat Password", exact=True).fill(_NEW_PWD)
            user_page.get_by_role("button", name="Save").click()
            user_page.wait_for_timeout(1000)
            expect(user_page.get_by_text("password changed", exact=False)).to_be_visible(
                timeout=5000
            )
        finally:
            ctx.close()

        # 验证新密码可登录
        assert _can_login(browser, _TEST_USER, _NEW_PWD), \
            f"修改后新密码登录用户 {_TEST_USER} 失败"
    finally:
        _delete_user(admin_page, _TEST_USER)
        _delete_role(admin_page, _CUSTOM_ROLE)
