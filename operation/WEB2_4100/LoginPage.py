import logging
import time

from selenium.common import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from api.ui_base_page import BasePage


class LoginPage(BasePage):
    USERNAME_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter User Name']")
    PASSWORD_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Password']")
    LOGIN_BTN = (By.XPATH, "//button[span[text()='Sign In']]")

    SETTINGS_BTN = (By.XPATH, "//span[normalize-space(.)='Settings']")
    USER_MANAGEMENT_BTN = (By.XPATH, "//div[normalize-space(.)='User Management']")
    USER_CONFIGURATION_BTN = (By.XPATH, "//span[normalize-space(.)='User Configuration']")
    ADD_USER_BTN = (By.XPATH, "//span[normalize-space(.)='Add User']")
    NEW_USERNAME_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Username']")
    NEW_PASSWORD_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Password']")
    CONFIRM_PASSWORD_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Repeat Password']")
    SELECT_ROLE_DROPDOWN = (By.XPATH, "//div[./span[text()='--Select Role--']]")
    SELECT_ADMIN_ROLE = (By.XPATH, "//span[normalize-space(.)='admin']")
    SELECT_VIEWER_ROLE = (By.XPATH, "//span[normalize-space(.)='view']")
    Override_Password_Policy = (By.XPATH, "//span[normalize-space(.)='Override Password Policy']")
    SAVE_BTN = (By.XPATH, "//button[span[text()='Save']]")

    DELETE_USERS_BTN = (By.XPATH, "//button[contains(@class, 'el-button--danger') and not(normalize-space("
                                         ".)='Lock')]")
    YES_DELETE_BTN = (By.XPATH, "//span[normalize-space(.)='Yes, continue']")

    def login(self, username, password):
        """执行登录流程"""
        self.input_text(self.USERNAME_INPUT, username)
        self.input_text(self.PASSWORD_INPUT, password)
        self.click(self.LOGIN_BTN)

    def add_user(self, usernames_passwords):
        """点击添加用户按钮"""
        self.click(self.SETTINGS_BTN)
        self.click(self.USER_MANAGEMENT_BTN)
        self.click(self.USER_CONFIGURATION_BTN)
        for username, password in usernames_passwords:
            self.click(self.ADD_USER_BTN)
            self.input_text(self.NEW_USERNAME_INPUT, username)
            self.input_text(self.NEW_PASSWORD_INPUT, password)
            self.input_text(self.CONFIRM_PASSWORD_INPUT, password)
            self.click(self.SELECT_ROLE_DROPDOWN)
            self.click(self.SELECT_ADMIN_ROLE)
            self.click(self.Override_Password_Policy)
            self.click(self.SAVE_BTN)
            print(f"添加用户 {username} 成功")

    def delete_all_users(self):
        """删除所有用户"""
        self.click(self.SETTINGS_BTN)
        self.click(self.USER_MANAGEMENT_BTN)
        self.click(self.USER_CONFIGURATION_BTN)
        # 确认删除所有用户
        click_count = 0
        sleep_time = 0.2  # 每次点击后等待1秒，确保DOM更新

        while True:
            delete_btn = self.find_elements(self.DELETE_USERS_BTN)
            if not delete_btn:
                logging.info(f"✅ 已删除{click_count}个用户。")
                return  # 退出循环

            try:
                # 始终点击找到的第一个元素，因为列表会在每次删除后更新
                element_to_click = delete_btn[0]
                element_to_click.click()
                self.click(self.YES_DELETE_BTN)
                click_count += 1
                logging.info(f"删除成功{click_count}个用户。")

                # 强制等待，等待DOM更新和删除操作完成
                time.sleep(sleep_time)

            except StaleElementReferenceException:
                # 元素在点击前消失（DOM更新太快），我们忽略并进行下一次循环查找
                logging.info(f"⚠️ 元素引用过期，进行下一次查找。")
            except ElementClickInterceptedException:
                # 点击被遮罩或其他元素拦截，等待并重试
                logging.info(f"❌ 警告：点击被拦截，等待并重试...")
                time.sleep(sleep_time + 1)
            except Exception as e:
                logging.info(f"❌ 发生点击错误，退出循环: {e}")
                break


