import logging
import time
from selenium.common import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from api.ui_base_page import BasePage


class UpdatePage(BasePage):
    def __init__(self, driver):
        super().__init__(driver)
        self.logger = logging.getLogger(__name__)

    SETTINGS_BTN = (By.XPATH, "//span[normalize-space(.)='Settings']")
    Firmware_Update_BTN = (By.XPATH, "//div[normalize-space(.)='Firmware Update']")
    Browse_BTN = (By.XPATH, "//span[normalize-space(.)='Browse']")
    Update_File_Input_INPUT = (By.XPATH, "//input[@type='file']")
    Upload_BTN = (By.XPATH, "//span[normalize-space(.)='Upload']")
    Update_BTN = (By.XPATH, "//span[normalize-space(.)='Update']")

    File_path = "C:/Users/ZihanGao/Downloads/AcuRev-4100-WEB2-v1.00p06.a2h"

    def update(self):
        self.click(self.SETTINGS_BTN)
        self.click(self.Firmware_Update_BTN)
        self.find_element(self.Update_File_Input_INPUT).send_keys(self.File_path)
        self.click(self.Upload_BTN)
        self.click(self.Update_BTN)
