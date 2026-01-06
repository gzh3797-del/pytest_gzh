import logging
import time
from selenium.webdriver.support import expected_conditions as EC
import pytest
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from operation.WEB2_4100.UpdatePage import UpdatePage
from operation.WEB2_4100.LoginPage import LoginPage


class TestUpdate:
    @pytest.mark.w_update
    def test_update(self, driver):
        """
        自动升级
        :param driver:
        :return:
        """
        n = 200
        for x in range(1, n+1):
            update_page = UpdatePage(driver)
            update_page.File_path = "C:/Users/ZihanGao/Downloads/AcuRev-4100-WEB2-v1.00p06.a2h"
            update_page.update()
            # 等待转到登录界面
            WebDriverWait(driver, 600).until(EC.url_contains("login"))
            logging.info(f"第{x}次升级完成")
            login_page = LoginPage(driver)
            # 等待登录页面加载
            WebDriverWait(driver, 100).until_not(
                EC.presence_of_element_located((By.CLASS_NAME, "el-loading-mask"))
            )
            login_page.login("admin", "admin")
