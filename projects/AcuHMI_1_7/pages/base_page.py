from playwright.sync_api import Page, expect
from projects.AcuHMI_1_7.settings import TIMEOUT


class BasePage:
    def __init__(self, page: Page):
        self.page = page
        self.timeout = TIMEOUT

    def navigate(self, url: str):
        self.page.goto(url, timeout=self.timeout)

    def click(self, selector: str):
        self.page.click(selector, timeout=self.timeout)

    def fill(self, selector: str, value: str):
        self.page.fill(selector, value, timeout=self.timeout)

    def get_text(self, selector: str) -> str:
        return self.page.text_content(selector, timeout=self.timeout)

    def is_visible(self, selector: str) -> bool:
        return self.page.is_visible(selector)

    def wait_for_selector(self, selector: str):
        self.page.wait_for_selector(selector, timeout=self.timeout)

    def take_screenshot(self, name: str):
        from projects.AcuHMI_1_7.settings import get_screenshot_dir
        path = get_screenshot_dir() / f"{name}.png"
        self.page.screenshot(path=str(path))
        return path

    def wait_for_load(self):
        self.page.wait_for_load_state("networkidle", timeout=self.timeout)
