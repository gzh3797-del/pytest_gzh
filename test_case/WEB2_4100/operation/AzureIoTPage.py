import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from comm.ui_base_page import BasePage


class AzureIoTPage(BasePage):
    """AcuHMI-1-7 / Protocols / Azure IoT 配置页面"""

    # 顶部导航栏 AcuHMI-1-7 按钮
    DEVICE_SETTINGS_BTN = (By.XPATH, "//div[contains(@class,'nav-item-menu')][.//span[normalize-space(.)='AcuHMI-1-7']]")

    # 左侧导航
    PROTOCOLS_MENU = (By.XPATH, "//li[contains(@class,'left-nav-item') and contains(normalize-space(.),'Protocols')]")
    AZURE_IOT_TAB  = (By.XPATH, "//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Azure IoT')]")

    # 主功能 Enable / Disable（第一组单选）
    ENABLE_RADIO  = (By.XPATH, "(//label[.//span[normalize-space(.)='Enable']])[1]")
    DISABLE_RADIO = (By.XPATH, "(//label[.//span[normalize-space(.)='Disable']])[1]")

    # Connection String 输入框（优先按 label 上下文定位，兜底按 placeholder）
    PRIMARY_CONN_STR_INPUT = (By.XPATH,
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'Primary')]]"
        "//input | "
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'Primary')]]"
        "//textarea")
    SECONDARY_CONN_STR_INPUT = (By.XPATH,
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'Secondary')]]"
        "//input | "
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'Secondary')]]"
        "//textarea")

    # Interval 下拉
    INTERVAL_DROPDOWN = (By.XPATH,
        "//label[contains(normalize-space(.),'Interval')]/following::div[contains(@class,'el-select')][1]")

    # Enable SSL 单选（第二组，位于 Interval 之后）
    SSL_ENABLE_RADIO  = (By.XPATH,
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'SSL')]]"
        "//label[.//span[normalize-space(.)='Enable']][1]")
    SSL_DISABLE_RADIO = (By.XPATH,
        "//div[contains(@class,'el-form-item')][.//label[contains(normalize-space(.),'SSL')]]"
        "//label[.//span[normalize-space(.)='Disable']][1]")

    # 证书文件 input（SSL Enable 后可见）
    CERT_FILE_INPUT = (By.XPATH, "(//input[@type='file'])[1]")
    KEY_FILE_INPUT  = (By.XPATH, "(//input[@type='file'])[2]")

    # 设备表格行（与 AWSIoTPage 相同结构）
    DEVICE_TABLE_ROWS = (By.XPATH,
        "//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]")

    # 操作按钮
    SAVE_BTN            = (By.XPATH, "//button[.//span[normalize-space(.)='Save']]")
    TEST_CONNECTION_BTN = (By.XPATH, "//button[.//span[normalize-space(.)='Test Connection']]")

    # 页面提示
    RESULT_MSG = (By.XPATH,
        "//div[contains(@class,'el-message') or contains(@class,'el-notification__content') "
        "or contains(@class,'el-message-box')]")

    # ------------------------------------------------------------------ #

    def _dismiss_overlay_dialog(self):
        """登录后页面可能残留确认框，点击取消/关闭将其关掉，避免遮挡后续操作。"""
        try:
            dialogs = self.driver.find_elements(
                By.XPATH,
                "//div[contains(@class,'el-overlay-message-box') or contains(@class,'el-message-box')]"
                "[not(contains(@style,'display:none'))]"
            )
            for dlg in dialogs:
                if not dlg.is_displayed():
                    continue
                for xpath in [
                    ".//button[.//span[normalize-space(.)='Yes']]",
                    ".//button[.//span[normalize-space(.)='Continue']]",
                    ".//button[.//span[normalize-space(.)='确认']]",
                    ".//button[contains(@class,'el-button--primary')]",
                    ".//button[contains(@class,'el-message-box__headerbtn')]",
                    ".//button[contains(@class,'el-button')]",
                ]:
                    btns = dlg.find_elements(By.XPATH, xpath)
                    if btns:
                        self.driver.execute_script("arguments[0].click()", btns[0])
                        logging.info("[Web] 已关闭页面残留确认框")
                        time.sleep(0.5)
                        break
        except Exception as e:
            logging.debug(f"[Web] 关闭残留弹框时忽略异常：{e}")

    def navigate_to_azure_iot(self):
        self._dismiss_overlay_dialog()
        self.click(self.DEVICE_SETTINGS_BTN)
        time.sleep(2)
        WebDriverWait(self.driver, 15).until(
            EC.element_to_be_clickable(self.PROTOCOLS_MENU)
        )
        self.click(self.PROTOCOLS_MENU)
        time.sleep(1)
        self.click(self.AZURE_IOT_TAB)
        time.sleep(1.5)

    def is_enabled(self) -> bool:
        try:
            label = self.find_element(self.ENABLE_RADIO)
            return 'is-checked' in (label.get_attribute('class') or '')
        except Exception:
            return False

    def ensure_enabled(self):
        if not self.is_enabled():
            logging.info("当前为 Disable，切换为 Enable")
            self.click(self.ENABLE_RADIO)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.PRIMARY_CONN_STR_INPUT)
            )

    def is_ssl_enabled(self) -> bool:
        try:
            label = self.find_element(self.SSL_ENABLE_RADIO)
            return 'is-checked' in (label.get_attribute('class') or '')
        except Exception:
            return False

    def ensure_ssl_enabled(self):
        if not self.is_ssl_enabled():
            logging.info("SSL 当前为 Disable，切换为 Enable")
            self.click(self.SSL_ENABLE_RADIO)
            time.sleep(1)

    def set_primary_connection_string(self, conn_str: str):
        self.input_text(self.PRIMARY_CONN_STR_INPUT, conn_str)

    def set_secondary_connection_string(self, conn_str: str):
        if conn_str:
            self.input_text(self.SECONDARY_CONN_STR_INPUT, conn_str)

    def set_interval(self, interval_text: str):
        self.click(self.INTERVAL_DROPDOWN)
        option = (By.XPATH,
                  f"//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{interval_text}']")
        el = self.find_element(option)
        self.driver.execute_script("arguments[0].click()", el)

    def upload_cert_file(self, file_path: str):
        self.find_element(self.CERT_FILE_INPUT).send_keys(file_path)
        time.sleep(1)

    def upload_key_file(self, file_path: str):
        self.find_element(self.KEY_FILE_INPUT).send_keys(file_path)
        time.sleep(1)

    def select_unchecked_devices(self):
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        logging.info(f"设备列表共 {len(rows)} 行，检查勾选状态")
        for idx, row in enumerate(rows):
            cb_labels = row.find_elements(By.XPATH, ".//label[contains(@class,'el-checkbox')]")
            if not cb_labels:
                logging.warning(f"  第 {idx + 1} 行：未找到复选框，跳过")
                continue
            cb = cb_labels[0]
            if 'is-checked' in (cb.get_attribute('class') or ''):
                logging.info(f"  第 {idx + 1} 行：已勾选，跳过")
            else:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'})", cb)
                time.sleep(0.2)
                self.driver.execute_script("arguments[0].click()", cb)
                logging.info(f"  第 {idx + 1} 行：已勾选设备")
                time.sleep(0.3)

    def _all_params_selected(self, dialog) -> bool:
        left_items = dialog.find_elements(By.XPATH,
            ".//div[contains(@class,'transfer-left')]"
            "//div[contains(@class,'content')]"
            "//div[contains(@class,'option-item')]")
        logging.info(f"    [check] transfer-left 剩余未选项：{len(left_items)}")
        return len(left_items) == 0

    def _get_visible_dialog(self):
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@class,'el-overlay-dialog')]//div[contains(@class,'el-dialog__body')]")
            )
        )
        wrappers = self.driver.find_elements(
            By.XPATH, "//div[contains(@class,'el-overlay-dialog')]")
        for w in wrappers:
            if w.is_displayed():
                dialog = w.find_elements(By.XPATH, ".//div[contains(@class,'el-dialog')]")
                if dialog:
                    return dialog[0]
        raise RuntimeError("未找到可见的弹窗")

    def _configure_parameter_dialog(self):
        dialog = self._get_visible_dialog()
        try:
            pt_wrapper = dialog.find_element(By.XPATH,
                ".//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Parameter Type')]]"
                "//div[contains(@class,'el-select__wrapper')]")
            self.driver.execute_script("arguments[0].click()", pt_wrapper)
            time.sleep(0.5)
            option_els = self.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-select__popper') or contains(@class,'el-select-dropdown')]"
                "//li[contains(@class,'el-select-dropdown__item')]")
            option_texts = [o.text.strip() for o in option_els if o.text.strip()]
            logging.info(f"Parameter Type 选项：{option_texts}")
            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.3)
            for opt_text in option_texts:
                self.driver.execute_script("arguments[0].click()", pt_wrapper)
                time.sleep(0.5)
                opts = self.driver.find_elements(By.XPATH,
                    "//div[contains(@class,'el-select__popper') or contains(@class,'el-select-dropdown')]"
                    "//li[contains(@class,'el-select-dropdown__item')]")
                for opt in opts:
                    if opt.text.strip() == opt_text:
                        self.driver.execute_script("arguments[0].click()", opt)
                        break
                time.sleep(0.5)
                if self._all_params_selected(dialog):
                    logging.info(f"  [{opt_text}] → 参数已全选，跳过")
                    continue
                all_btns = dialog.find_elements(By.XPATH,
                    ".//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]")
                if all_btns:
                    self.driver.execute_script("arguments[0].click()", all_btns[0])
                    logging.info(f"  [{opt_text}] → 已点击 All")
                    time.sleep(0.3)
                else:
                    logging.warning(f"  [{opt_text}] → All 按钮未找到")
        except Exception as e:
            logging.error(f"配置 Parameter 弹窗时出错：{e}")
        finally:
            confirmed = False
            for xpath in [
                ".//div[contains(@class,'dialog-footer')]//button[contains(@class,'el-button--primary')]",
                ".//footer[contains(@class,'el-dialog__footer')]//button[contains(@class,'el-button--primary')]",
                ".//div[contains(@class,'dialog-footer')]//button[last()]",
            ]:
                btns = dialog.find_elements(By.XPATH, xpath)
                if btns:
                    self.driver.execute_script("arguments[0].click()", btns[0])
                    logging.info("已点击 Confirm，关闭弹窗")
                    confirmed = True
                    break
            if not confirmed:
                try:
                    close = dialog.find_element(
                        By.XPATH, ".//button[contains(@class,'el-dialog__headerbtn')]")
                    self.driver.execute_script("arguments[0].click()", close)
                    logging.info("已点击 × 关闭弹窗")
                    confirmed = True
                except Exception:
                    pass
            if not confirmed:
                # 最终兜底：按 Escape 关闭弹窗
                try:
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    logging.info("已按 Escape 关闭弹窗")
                except Exception:
                    logging.warning("无法关闭弹窗，继续执行")
            time.sleep(0.8)

    def configure_all_devices_parameters(self):
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        logging.info(f"设备列表共 {len(rows)} 行，逐行配置参数选择")
        for idx, row in enumerate(rows):
            btns = row.find_elements(By.XPATH, ".//button")
            if not btns:
                logging.warning(f"  第 {idx + 1} 行未找到按钮，跳过")
                continue
            target_btn = None
            for b in btns:
                txt = (b.text or b.get_attribute('innerText') or '').strip()
                if any(kw in txt for kw in ('Select', 'Parameter', '选择', '参数')):
                    target_btn = b
                    break
            if target_btn is None:
                target_btn = btns[-1]
            label = (target_btn.text or target_btn.get_attribute('innerText') or '').strip()
            logging.info(f"  第 {idx + 1} 行：点击按钮「{label or '(icon)'}」")
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'})", target_btn)
            time.sleep(0.3)
            self.driver.execute_script("arguments[0].click()", target_btn)
            time.sleep(1.5)
            self._configure_parameter_dialog()

    def save(self):
        btn = self.find_element(self.SAVE_BTN)
        self.driver.execute_script("arguments[0].click()", btn)
        time.sleep(3)
        logging.info("已点击 Save")

    def test_connection(self) -> str:
        btn = self.find_element(self.TEST_CONNECTION_BTN)
        self.driver.execute_script("arguments[0].click()", btn)
        logging.info("已点击 Test Connection，等待结果…")
        try:
            msg_el = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located(self.RESULT_MSG))
            result = msg_el.text.strip()
            logging.info(f"Test Connection 结果：{result}")
            return result
        except Exception:
            logging.warning("未在 30 秒内检测到页面反馈")
            return "未检测到结果提示"

    def get_checked_device_names(self) -> list:
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        names = []
        for row in rows:
            cb_labels = row.find_elements(By.XPATH, ".//label[contains(@class,'el-checkbox')]")
            if not cb_labels:
                continue
            if 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                continue
            for td in row.find_elements(By.XPATH, ".//td"):
                if td.find_elements(By.XPATH, ".//*[self::button or self::input]"):
                    continue
                text = td.text.strip().split('\n')[0].strip()
                if text:
                    names.append(text)
                    break
        return names

    def disable(self):
        if self.is_enabled():
            logging.info("切换 Azure IoT 为 Disable")
            self.click(self.DISABLE_RADIO)
            time.sleep(0.5)
        self.save()
        logging.info("Azure IoT 已 Disable 并保存")

    def configure(self, primary_conn_str: str, interval: str = "30 seconds",
                  secondary_conn_str: str = "",
                  enable_ssl: bool = False,
                  cert_file: str = "", key_file: str = ""):
        """
        完整配置流程：
        1. 确保 Enable 已勾选
        2. 填写 Connection String 和 Interval
        3. 配置 SSL（按需开启并上传证书）
        4. 全选设备
        5. 配置各设备参数
        6. 保存并测试连接
        """
        self.ensure_enabled()
        self.set_primary_connection_string(primary_conn_str)
        self.set_secondary_connection_string(secondary_conn_str)
        self.set_interval(interval)
        if enable_ssl:
            self.ensure_ssl_enabled()
            if cert_file:
                self.upload_cert_file(cert_file)
            if key_file:
                self.upload_key_file(key_file)
        self.select_unchecked_devices()
        self.configure_all_devices_parameters()
        self._dismiss_overlay_dialog()
        self.save()
        return self.test_connection()
