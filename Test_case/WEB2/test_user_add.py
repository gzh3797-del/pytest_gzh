import logging
import time
import logger
import pytest
from selenium.webdriver.common.by import By
from operation.WEB2.login import LoginPage


class TestUserOperations:
    @pytest.mark.web2_add
    def test_add_user_flow(self, web2_driver):
        login_page = web2_driver
        user_numbers = 5
        data_sequence = [[f'user{i}', i] for i in range(1, user_numbers + 1)]
        for username, password in data_sequence:
            login_page.add_user(username, password)

        time.sleep(0.2)
        # 检查data_sequence中最大的username是否在页面上
        max_user = max(data_sequence, key=lambda x: x[1])[0]
        span_max_user = login_page.driver.find_element(By.XPATH, f"//span[normalize-space(.)='{max_user}']")
        assert span_max_user
        logging.info(f"✅ 成功添加{user_numbers}个用户")

    @pytest.mark.web2_delete
    def test_delete_all_users(self, web2_driver):
        login_page = web2_driver
        login_page.delete_all_users()
