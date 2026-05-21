import logging
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from comm.ui_base_page import BasePage


class AWSIoTPage(BasePage):
    """AcuHMI-1-7 / Protocols / AWS IoT 配置页面"""

    # 顶部导航栏 AcuHMI-1-7 按钮（进入设备配置视图）
    DEVICE_SETTINGS_BTN = (By.XPATH, "//div[contains(@class,'nav-item-menu')][.//span[normalize-space(.)='AcuHMI-1-7']]")

    # 左侧导航（设备配置视图中）
    PROTOCOLS_MENU = (By.XPATH, "//li[contains(@class,'left-nav-item') and contains(normalize-space(.),'Protocols')]")
    AWS_IOT_TAB = (By.XPATH, "//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'AWS IoT')]")

    # Enable / Disable 单选（取 label 元素，含 is-checked 类）
    ENABLE_RADIO = (By.XPATH, "//label[.//span[normalize-space(.)='Enable']]")
    DISABLE_RADIO = (By.XPATH, "//label[.//span[normalize-space(.)='Disable']]")

    # 文本输入
    CLIENT_ID_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Client Id']")
    URL_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter URL']")
    TOPIC_INPUT = (By.CSS_SELECTOR, "input[placeholder='Enter Topic']")

    # Interval — element-ui 自定义下拉
    INTERVAL_DROPDOWN = (By.XPATH, "//label[contains(normalize-space(.),'Interval')]/following::div[contains(@class,'el-select')][1]")

    # 隐藏的 file input（直接 send_keys 文件路径）
    CERT_FILE_INPUT = (By.XPATH, "(//input[@type='file'])[1]")
    KEY_FILE_INPUT = (By.XPATH, "(//input[@type='file'])[2]")

    # Devices Selection 全选复选框（表头）
    SELECT_ALL_CHECKBOX = (By.XPATH, "//div[contains(@class,'el-table')]//thead//label[contains(@class,'el-checkbox__original') or contains(@class,'el-checkbox')]")

    # 设备表格行
    DEVICE_TABLE_ROWS = (By.XPATH,
        "//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]")

    # 操作按钮
    SAVE_BTN = (By.XPATH, "//button[.//span[normalize-space(.)='Save']]")
    TEST_CONNECTION_BTN = (By.XPATH, "//button[.//span[normalize-space(.)='Test Connection']]")

    # 页面结果提示（element-ui message / notification）
    RESULT_MSG = (By.XPATH, "//div[contains(@class,'el-message') or contains(@class,'el-notification__content') or contains(@class,'el-message-box')]")

    # ------------------------------------------------------------------ #

    def navigate_to_aws_iot(self):
        # 第一步：点击顶部导航 AcuHMI-1-7 进入设备配置视图
        self.click(self.DEVICE_SETTINGS_BTN)
        time.sleep(2)  # 等待 Vue Router 路由跳转完成
        # 第二步：等待侧边 Protocols 可点击并点击
        WebDriverWait(self.driver, 15).until(
            EC.element_to_be_clickable(self.PROTOCOLS_MENU)
        )
        self.click(self.PROTOCOLS_MENU)
        time.sleep(1)
        # 第三步：点击 AWS IoT 标签，等待 Tab 内容渲染
        self.click(self.AWS_IOT_TAB)
        time.sleep(1.5)

    def is_enabled(self) -> bool:
        """检查当前是否已处于 Enable 状态"""
        try:
            label = self.find_element(self.ENABLE_RADIO)
            return 'is-checked' in (label.get_attribute('class') or '')
        except Exception:
            return False

    def ensure_enabled(self):
        """若当前为 Disable，切换到 Enable 后等待表单出现"""
        if not self.is_enabled():
            logging.info("当前为 Disable，切换为 Enable")
            self.click(self.ENABLE_RADIO)
            # 等待输入框出现（表单在 Enable 后才展开）
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.CLIENT_ID_INPUT)
            )

    def set_client_id(self, client_id: str):
        self.input_text(self.CLIENT_ID_INPUT, client_id)

    def set_url(self, url: str):
        self.input_text(self.URL_INPUT, url)

    def set_topic(self, topic: str):
        self.input_text(self.TOPIC_INPUT, topic)

    def set_interval(self, interval_text: str):
        """interval_text 为下拉可见文字，如 '30 seconds'"""
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
        """逐行检查设备复选框，仅勾选尚未勾选的行"""
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        logging.info(f"设备列表共 {len(rows)} 行，检查勾选状态")
        for idx, row in enumerate(rows):
            cb_labels = row.find_elements(
                By.XPATH, ".//label[contains(@class,'el-checkbox')]"
            )
            if not cb_labels:
                logging.warning(f"  第 {idx + 1} 行：未找到复选框，跳过")
                continue
            cb = cb_labels[0]
            cls = cb.get_attribute('class') or ''
            if 'is-checked' in cls:
                logging.info(f"  第 {idx + 1} 行：已勾选，跳过")
            else:
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'})", cb
                )
                time.sleep(0.2)
                self.driver.execute_script("arguments[0].click()", cb)
                logging.info(f"  第 {idx + 1} 行：已勾选设备")
                time.sleep(0.3)

    def _all_params_selected(self, dialog) -> bool:
        """
        检查 transfer 左侧 Not Selected 区域是否为空。
        左侧无 option-item 说明所有参数已移到右侧（全选）。
        """
        left_items = dialog.find_elements(By.XPATH,
            ".//div[contains(@class,'transfer-left')]"
            "//div[contains(@class,'content')]"
            "//div[contains(@class,'option-item')]"
        )
        logging.info(f"    [check] transfer-left 剩余未选项：{len(left_items)}")
        return len(left_items) == 0

    def _get_visible_dialog(self):
        """返回当前可见的 el-dialog 元素（兼容 Element Plus el-overlay-dialog）"""
        # Element Plus 用 el-overlay-dialog 作为外层 wrapper
        WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@class,'el-overlay-dialog')]//div[contains(@class,'el-dialog__body')]")
            )
        )
        wrappers = self.driver.find_elements(
            By.XPATH, "//div[contains(@class,'el-overlay-dialog')]"
        )
        for w in wrappers:
            if w.is_displayed():
                dialog = w.find_elements(By.XPATH, ".//div[contains(@class,'el-dialog')]")
                if dialog:
                    return dialog[0]
        raise RuntimeError("未找到可见的弹窗")

    def _configure_parameter_dialog(self):
        """
        在 Parameter Selection 弹窗中遍历 Parameter Type 下拉的所有选项，
        每个选项选中后点击 All，最后点击 Confirm 关闭弹窗。
        """
        dialog = self._get_visible_dialog()
        try:
            # Parameter Type 是 el-select 下拉框，找到其 wrapper
            pt_wrapper = dialog.find_element(By.XPATH,
                ".//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Parameter Type')]]"
                "//div[contains(@class,'el-select__wrapper')]"
            )

            # 打开下拉，采集所有选项文本
            self.driver.execute_script("arguments[0].click()", pt_wrapper)
            time.sleep(0.5)

            option_els = self.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-select__popper') or contains(@class,'el-select-dropdown')]"
                "//li[contains(@class,'el-select-dropdown__item')]"
            )
            option_texts = [o.text.strip() for o in option_els if o.text.strip()]
            logging.info(f"Parameter Type 选项：{option_texts}")

            # 关闭下拉
            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(0.3)

            for opt_text in option_texts:
                # 每次重新打开下拉选择
                self.driver.execute_script("arguments[0].click()", pt_wrapper)
                time.sleep(0.5)
                opts = self.driver.find_elements(By.XPATH,
                    "//div[contains(@class,'el-select__popper') or contains(@class,'el-select-dropdown')]"
                    "//li[contains(@class,'el-select-dropdown__item')]"
                )
                for opt in opts:
                    if opt.text.strip() == opt_text:
                        self.driver.execute_script("arguments[0].click()", opt)
                        break
                time.sleep(0.5)

                # 若参数已全部选中，跳过
                if self._all_params_selected(dialog):
                    logging.info(f"  [{opt_text}] → 参数已全选，跳过")
                    continue

                # 点击 All 按钮（参数列表全选）
                all_btns = dialog.find_elements(By.XPATH,
                    ".//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]"
                )
                if all_btns:
                    self.driver.execute_script("arguments[0].click()", all_btns[0])
                    logging.info(f"  [{opt_text}] → 已点击 All")
                    time.sleep(0.3)
                else:
                    logging.warning(f"  [{opt_text}] → All 按钮未找到")

        except Exception as e:
            logging.error(f"配置 Parameter 弹窗时出错：{e}")
        finally:
            # 确保弹窗始终被关闭
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
                        By.XPATH, ".//button[contains(@class,'el-dialog__headerbtn')]"
                    )
                    self.driver.execute_script("arguments[0].click()", close)
                    logging.info("已点击 × 关闭弹窗")
                except Exception:
                    logging.warning("无法关闭弹窗，继续执行")
            time.sleep(0.8)

    def _debug_dom_after_click(self):
        """点击按钮后，用 JS 打印 DOM 中所有 dialog/modal 相关元素的信息"""
        info = self.driver.execute_script("""
            var els = document.querySelectorAll('[class*="dialog"],[class*="modal"],[class*="popup"]');
            var result = [];
            for (var i = 0; i < Math.min(els.length, 20); i++) {
                var el = els[i];
                var cs = window.getComputedStyle(el);
                result.push({
                    tag: el.tagName,
                    cls: el.className,
                    display: cs.display,
                    visibility: cs.visibility,
                    outerHTML: el.outerHTML.substring(0, 300)
                });
            }
            return result;
        """)
        for item in info:
            logging.info(f"[DOM] tag={item['tag']} cls={item['cls']} "
                         f"display={item['display']} visibility={item['visibility']}")
            logging.info(f"      HTML={item['outerHTML']}")

    def configure_all_devices_parameters(self):
        """
        遍历设备列表每一行，点击 Parameter Selection 按钮，
        在弹窗中对所有 Parameter Type 分组选中全部参数。
        """
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        logging.info(f"设备列表共 {len(rows)} 行，逐行配置参数选择")
        for idx, row in enumerate(rows):
            btns = row.find_elements(By.XPATH, ".//button")
            if not btns:
                logging.warning(f"  第 {idx + 1} 行未找到按钮，跳过")
                continue
            # 优先找包含 "Select" / "Parameter" 文字的按钮
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
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'})", target_btn
            )
            time.sleep(0.3)
            self.driver.execute_script("arguments[0].click()", target_btn)
            time.sleep(1.5)
            self._configure_parameter_dialog()

    def save(self):
        self.click(self.SAVE_BTN)
        time.sleep(1)
        logging.info("已点击 Save")

    def test_connection(self) -> str:
        """点击 Test Connection 并等待页面反馈，返回提示文字"""
        self.click(self.TEST_CONNECTION_BTN)
        logging.info("已点击 Test Connection，等待结果…")
        try:
            msg_el = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located(self.RESULT_MSG)
            )
            result = msg_el.text.strip()
            logging.info(f"Test Connection 结果：{result}")
            return result
        except Exception:
            logging.warning("未在 30 秒内检测到页面反馈")
            return "未检测到结果提示"

    def get_checked_device_names(self) -> list:
        """返回设备表格中所有已勾选行的设备名称列表。"""
        rows = self.driver.find_elements(*self.DEVICE_TABLE_ROWS)
        names = []
        for row in rows:
            cb_labels = row.find_elements(By.XPATH, ".//label[contains(@class,'el-checkbox')]")
            if not cb_labels:
                continue
            if 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                continue
            # 找到不含 button / input 的文本单元格作为设备名
            for td in row.find_elements(By.XPATH, ".//td"):
                if td.find_elements(By.XPATH, ".//*[self::button or self::input]"):
                    continue
                text = td.text.strip().split('\n')[0].strip()
                if text:
                    names.append(text)
                    break
        return names

    def disable(self):
        """点击 Disable 单选并保存，关闭 AWS IoT 推送。"""
        if self.is_enabled():
            logging.info("切换 AWS IoT 为 Disable")
            self.click(self.DISABLE_RADIO)
            time.sleep(0.5)
        self.save()
        logging.info("AWS IoT 已 Disable 并保存")

    def configure(self, client_id: str, url: str, topic: str,
                  cert_file: str, key_file: str,
                  interval: str = "30 seconds"):
        """
        完整配置流程：
        1. 确保 Enable 已勾选
        2. 填写连接参数
        3. 上传证书文件
        4. 全选设备
        5. 保存
        6. 测试连接并返回结果
        """
        self.ensure_enabled()
        self.set_client_id(client_id)
        self.set_url(url)
        self.set_topic(topic)
        self.set_interval(interval)
        self.upload_cert_file(cert_file)
        self.upload_key_file(key_file)
        self.select_unchecked_devices()
        self.configure_all_devices_parameters()
        self.save()
        result = self.test_connection()
        return result
