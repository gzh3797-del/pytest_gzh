import logging

from selenium.common import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class BasePage:
    """所有页面的基类，封装通用操作"""
    def __init__(self, driver):
        self.driver = driver
        self.timeout = 10  # 默认等待10秒

    def find_element(self, locator):
        """
        查找元素，包含显式等待
        locator: (By.ID, "username") 这种元组形式
        """
        try:
            return WebDriverWait(self.driver, self.timeout).until(
                EC.presence_of_element_located(locator)
            )
        except Exception as e:
            print(f"元素未找到: {locator} - {e}")
            raise e

    def find_elements(self, locator: tuple) -> list:
        """
        查找所有匹配的元素列表，包含显式等待 (等待至少一个元素出现在DOM中)
        如果元素未找到，返回空列表 []
        locator: (By.CSS_SELECTOR, "div.item") 这种元组形式
        """
        try:
            # 使用 presence_of_all_elements_located 等待至少一个元素出现
            return WebDriverWait(self.driver, self.timeout).until(
                EC.presence_of_all_elements_located(locator)
            )
        except TimeoutException as e:
            logging.info(f"未找到元素: {locator}")
            return []

    def click(self, locator):
        """点击操作"""
        element = self.find_element(locator)
        # 确保元素可点击
        WebDriverWait(self.driver, self.timeout).until(
            EC.element_to_be_clickable(locator)
        )
        element.click()

    def input_text(self, locator, text):
        """输入文本操作"""
        element = self.find_element(locator)
        element.clear()  # 输入前先清空
        element.send_keys(text)