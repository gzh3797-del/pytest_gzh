import logging
import time
from selenium.webdriver.support import expected_conditions as EC
import pytest
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from operation.WEB2.UpdatePage import UpdatePage
from operation.WEB2.LoginPage import LoginPage


class TestUpdate:
    @pytest.mark.w_update
    def test_update(self, driver):
        """
        自动升级
        :param driver:
        :return:
        """
        n = 100
        for x in range(1, n+1):
            update_page = UpdatePage(driver)
            update_page.update()
            WebDriverWait(driver, 300).until(EC.url_contains("login"))
            logging.info(f"第{x}次升级完成")
            login_page = LoginPage(driver)
            # 等待页面加载
            WebDriverWait(driver, 100).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-mask"))
            )
            login_page.login("admin", "admin")
