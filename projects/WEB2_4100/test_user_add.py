"""
终端运行
| **批量添加用户** | `pytest -m w_add` | 默认执行 20 个管理员用户添加 |
| **清理用户数据** | `pytest -m w_delete` | 删除所有非admin用户 |
默认URL：192.168.2.209。在 conftest.py 文件的 driver() 中修改
"""
import logging
import time
import pytest
from selenium.webdriver.common.by import By

from projects.WEB2_4100.operation.LoginPage import LoginPage


class TestUserOperations:
    @pytest.mark.w_add
    def test_add_user_flow(self,  driver):
        """
        批量添加用户
        :param driver:
        :return:
        """
        login_page = LoginPage(driver)
        user_numbers = 20
        usernames_passwords = [[f'user{i}', i] for i in range(1, user_numbers + 1)]
        login_page.add_user(usernames_passwords)
        time.sleep(0.2)
        # 检查data_sequence中最大的username是否在页面上
        max_user = max(usernames_passwords, key=lambda x: x[1])[0]
        span_max_user = login_page.driver.find_element(By.XPATH, f"//span[normalize-space(.)='{max_user}']")
        assert span_max_user
        logging.info(f"✅ 成功添加{user_numbers}个用户")

    @pytest.mark.w_delete
    def test_delete_all_users(self,  driver):
        """
        删除所有用户
        :param driver:
        :return:
        """
        login_page = LoginPage(driver)
        login_page.delete_all_users()



