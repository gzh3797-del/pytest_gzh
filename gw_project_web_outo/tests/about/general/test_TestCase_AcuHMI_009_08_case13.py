import pytest
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage

# View权限测试用户凭据（系统默认view权限用户）
_VIEW_USERNAME = "view"
_VIEW_PASSWORD = "View@110002"


def _nav_to_about(page):
    page.locator("header span").filter(has_text="About").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _login_as_view(browser) -> tuple:
    """以View权限用户创建新browser context并登录，返回 (context, page)"""
    ctx = browser.new_context(ignore_https_errors=True)
    p = ctx.new_page()
    p.goto(BASE_URL + "/#/login")
    p.wait_for_load_state("networkidle")
    p.get_by_role("textbox", name="Enter User Name").fill(_VIEW_USERNAME)
    p.get_by_role("textbox", name="Enter User Name").press("Tab")
    p.get_by_role("textbox", name="Enter Password").fill(_VIEW_PASSWORD)
    p.get_by_role("button", name="Sign In").click()
    p.wait_for_load_state("networkidle")
    p.wait_for_timeout(1500)
    # 处理 EULA 弹窗（首次登录时出现，点击 Accept）
    try:
        p.get_by_role("button", name="Accept", exact=True).click(timeout=3000)
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(1000)
    except Exception:
        pass
    # 关闭"使用默认密码"确认弹窗（精确匹配 Close 按钮，避免匹配关闭图标）
    try:
        p.get_by_role("button", name="Close", exact=True).click(timeout=3000)
        p.wait_for_load_state("networkidle")
        p.wait_for_timeout(500)
    except Exception:
        pass
    return ctx, p


# 用例编号：TestCase_AcuHMI_009_08_case13
# 用例标题：View权限用户进入About页面，Save按钮不可见且显示无操作权限提示
# 预置条件：1.以View权限用户账号登录AcuHMI网页
# 测试步骤：
#   1. 以View权限用户登录AcuHMI网页
#   2. 进入About页面
#   3. 查看Name/Location/Description输入区域及Save按钮显示状态
# 预期结果：
#   1. 页面不显示Save按钮
#   2. 页面显示无权限提示信息
def test_TestCase_AcuHMI_009_08_case13(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page
    browser = page.context.browser

    ctx, view_page = _login_as_view(browser)
    try:
        assert "/#/login" not in view_page.url, \
            f"View用户 '{_VIEW_USERNAME}' 应登录成功，当前URL: {view_page.url}"

        _nav_to_about(view_page)

        # 验证Save按钮不可见
        save_btn = view_page.get_by_role("button", name="Save")
        assert save_btn.count() == 0 or not save_btn.is_visible(), \
            "View权限用户进入About页面时不应显示Save按钮"

        # 验证显示无权限提示信息（提示文字可能包含"permission"/"权限"/"read only"等）
        permission_hint = (
            view_page.get_by_text("permission", exact=False).count() > 0
            or view_page.get_by_text("权限", exact=False).count() > 0
            or view_page.get_by_text("read only", exact=False).count() > 0
            or view_page.get_by_text("Read Only", exact=False).count() > 0
            or view_page.locator(".permission-tip, .no-permission, [class*='permission']").count() > 0
        )
        assert permission_hint, \
            "View权限用户进入About页面时应显示无操作权限提示信息"
    finally:
        ctx.close()
