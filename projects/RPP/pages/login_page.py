from playwright.sync_api import Page, expect
from projects.RPP.settings import DEFAULT_USERNAME, BASE_URL, DEFAULT_PASSWORD
from projects.RPP.pages.base_page import BasePage


class LoginPage(BasePage):

    def __init__(self, page: Page):
        super().__init__(page)
        self.url = BASE_URL + "/#/login"

    # 打开登录页面 + 等待页面加载完成
    def open(self):
        self.navigate(self.url)
        self.wait_for_load()

    # 打开登录页面 + 登录页面
    def login(self, username: str = DEFAULT_USERNAME, password: str = DEFAULT_PASSWORD):
        self.page.get_by_role("textbox", name="Enter User Name").fill(username)
        self.page.get_by_role("textbox", name="Enter User Name").press("Tab")
        self.page.get_by_role("textbox", name="Enter Password").fill(password)
        self.page.get_by_role("button", name="Sign In").click()
        self.wait_for_load()
        # 若登录后弹出"修改默认密码"提示框，点击 Cancel 关闭，不影响正常登录流程
        try:
            self.page.get_by_role("button", name="Cancel").click(timeout=3000)
        except Exception:
            pass

    def is_logged_in(self) -> bool:
        # RPP 顶部导航组（About/Commissioning/Monitoring/Settings + Logout）出现即表示登录成功。
        # 用 "Logout" 判据：仅登录后存在，避免与登录页标题/其他文本误匹配。
        return self.page.locator(".nav-item, header span").filter(
            has_text="Logout").first.is_visible()
