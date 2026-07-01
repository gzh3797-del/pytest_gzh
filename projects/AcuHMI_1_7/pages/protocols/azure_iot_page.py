import logging
import time

from pages.base_page import BasePage
from config.settings import BASE_URL


class AzureIoTPage(BasePage):
    """AcuHMI-1-7 Azure IoT 配置页面（Playwright 实现）"""

    # ── XPath / CSS 定位器常量 ─────────────────────────────────────────────────

    # 顶部导航栏 AcuHMI-1-7 按钮
    DEVICE_SETTINGS_BTN = "xpath=//div[contains(@class,'nav-item-menu')][.//span[normalize-space(.)='AcuHMI-1-7']]"

    # 左侧导航
    PROTOCOLS_MENU = "xpath=//li[contains(@class,'left-nav-item') and contains(normalize-space(.),'Protocols')]"
    AZURE_IOT_TAB  = "xpath=//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Azure IoT')]"

    # Enable / Disable 单选
    ENABLE_RADIO  = "xpath=//label[.//span[normalize-space(.)='Enable']]"
    DISABLE_RADIO = "xpath=//label[.//span[normalize-space(.)='Disable']]"

    # 文本输入
    PRIMARY_CONN_STR_INPUT   = "css=input[placeholder='Enter Primary Connection String']"
    SECONDARY_CONN_STR_INPUT = "css=input[placeholder='Enter Secondary Connection String']"

    # Interval 下拉
    INTERVAL_DROPDOWN = "xpath=//label[contains(normalize-space(.),'Interval')]/following::div[contains(@class,'el-select')][1]"

    # 设备表格行
    DEVICE_TABLE_ROWS = "xpath=//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]"

    # 操作按钮
    SAVE_BTN            = "xpath=//button[.//span[normalize-space(.)='Save']]"
    TEST_CONNECTION_BTN = "xpath=//button[.//span[normalize-space(.)='Test Connection']]"

    # 结果提示
    RESULT_MSG = "xpath=//div[contains(@class,'el-message') or contains(@class,'el-notification__content') or contains(@class,'el-message-box')]"

    # Virtual Devices 相关
    DEVICES_NAV_BTN      = "xpath=//div[contains(@class,'nav-item-menu') and .//span[normalize-space(.)='Devices']]"
    VIRTUAL_DEVICES_MENU = "xpath=//li[contains(normalize-space(.),'Virtual Devices')] | //a[contains(normalize-space(.),'Virtual Devices')]"
    READING_TAB          = "xpath=//*[contains(@class,'el-tabs__item') and contains(normalize-space(.),'Reading')]"
    READING_TABLE_ROWS   = "xpath=//div[contains(@class,'el-table__body-wrapper')]//tbody/tr[contains(@class,'el-table__row')]"

    # ── 内部辅助 ──────────────────────────────────────────────────────────────

    def _dismiss_overlay_dialog(self):
        """关闭页面可能残留的确认弹框，避免遮挡后续操作。"""
        try:
            dialogs = self.page.locator(
                "xpath=//div[contains(@class,'el-overlay-message-box') or contains(@class,'el-message-box')]"
            ).all()
            for dlg in dialogs:
                if not dlg.is_visible():
                    continue
                for xpath in [
                    ".//button[.//span[normalize-space(.)='Yes']]",
                    ".//button[.//span[normalize-space(.)='Continue']]",
                    ".//button[.//span[normalize-space(.)='确认']]",
                    ".//button[contains(@class,'el-button--primary')]",
                    ".//button[contains(@class,'el-message-box__headerbtn')]",
                    ".//button[contains(@class,'el-button')]",
                ]:
                    btns = dlg.locator(f"xpath={xpath}").all()
                    if btns:
                        btns[0].evaluate("el => el.click()")
                        logging.info("[Web] 已关闭页面残留确认框")
                        self.page.wait_for_timeout(500)
                        break
        except Exception as e:
            logging.debug(f"[Web] 关闭残留弹框时忽略异常：{e}")

    def _clear_and_fill(self, selector: str, value: str):
        """清空并填写输入框。"""
        loc = self.page.locator(selector)
        loc.triple_click()
        loc.fill(value)

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def navigate_to_azure_iot(self):
        """导航到 Azure IoT 配置页面。"""
        self._dismiss_overlay_dialog()
        # 第一步：点击顶部导航 AcuHMI-1-7 进入设备配置视图
        self.page.locator(self.DEVICE_SETTINGS_BTN).click()
        self.page.wait_for_timeout(2000)
        # 第二步：等待侧边 Protocols 可点击并点击
        self.page.wait_for_selector(self.PROTOCOLS_MENU, state="visible", timeout=15000)
        self.page.locator(self.PROTOCOLS_MENU).click()
        self.page.wait_for_timeout(1000)
        # 第三步：点击 Azure IoT 标签，等待 Tab 内容渲染
        self.page.locator(self.AZURE_IOT_TAB).click()
        self.page.wait_for_timeout(1500)

    def is_enabled(self) -> bool:
        """检查当前是否已处于 Enable 状态。"""
        try:
            label = self.page.locator(self.ENABLE_RADIO)
            cls = label.get_attribute('class') or ''
            return 'is-checked' in cls
        except Exception:
            return False

    def ensure_enabled(self):
        """若当前为 Disable，切换到 Enable 后等待表单出现。"""
        if not self.is_enabled():
            logging.info("当前为 Disable，切换为 Enable")
            self.page.locator(self.ENABLE_RADIO).click()
            # 等待输入框出现（表单在 Enable 后才展开）
            self.page.wait_for_selector(self.PRIMARY_CONN_STR_INPUT, state="visible", timeout=10000)

    def set_primary_conn_str(self, conn_str: str):
        self._clear_and_fill(self.PRIMARY_CONN_STR_INPUT, conn_str)

    def set_secondary_conn_str(self, conn_str: str):
        self._clear_and_fill(self.SECONDARY_CONN_STR_INPUT, conn_str)

    def set_interval(self, interval_text: str):
        """interval_text 为下拉可见文字，如 '30 seconds'。"""
        logging.info(f"设置 Interval：{interval_text}")
        self.page.locator(self.INTERVAL_DROPDOWN).evaluate("el => el.click()")
        self.page.wait_for_timeout(400)
        option_xpath = (
            f"xpath=//li[contains(@class,'el-select-dropdown__item') "
            f"and normalize-space(.)='{interval_text}']"
        )
        self.page.wait_for_selector(option_xpath, state="visible", timeout=5000)
        self.page.locator(option_xpath).evaluate("el => el.click()")
        self.page.wait_for_timeout(300)

    def select_unchecked_devices(self):
        """逐行检查设备复选框，仅勾选尚未勾选的行。"""
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"设备列表共 {len(rows)} 行，检查勾选状态")
        for idx, row in enumerate(rows):
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                logging.warning(f"  第 {idx + 1} 行：未找到复选框，跳过")
                continue
            cb = cb_labels[0]
            cls = cb.get_attribute('class') or ''
            if 'is-checked' in cls:
                logging.info(f"  第 {idx + 1} 行：已勾选，跳过")
            else:
                cb.evaluate("el => el.scrollIntoView({block:'center'})")
                self.page.wait_for_timeout(200)
                cb.evaluate("el => el.click()")
                logging.info(f"  第 {idx + 1} 行：已勾选设备")
                self.page.wait_for_timeout(300)

    def _transfer_left_count(self, dialog_locator) -> int:
        """返回 transfer 左侧（未选）区域的条目数，多个 XPath fallback 取最大值。"""
        for xpath in (
            "xpath=(.//*[contains(@class,'el-transfer-panel')])[1]"
            "//*[contains(@class,'el-transfer-panel__item')]",
            "xpath=.//div[contains(@class,'transfer-left')]"
            "//div[contains(@class,'option-item')]",
            "xpath=(.//*[contains(@class,'el-transfer-panel')])[1]"
            "//label[contains(@class,'el-checkbox')]",
        ):
            try:
                cnt = len(dialog_locator.locator(xpath).all())
                if cnt > 0:
                    return cnt
            except Exception:
                continue
        return 0

    def _all_params_selected(self, dialog_locator) -> bool:
        """检查 transfer 左侧 Not Selected 区域是否为空。"""
        cnt = self._transfer_left_count(dialog_locator)
        logging.info(f"    [check] transfer-left 剩余未选项：{cnt}")
        return cnt == 0

    def _dismiss_message_box(self):
        """关闭可能出现的 el-message-box 确认框，不关闭参数配置弹窗。"""
        try:
            msg_boxes = self.page.locator(
                "xpath=//div[contains(@class,'el-overlay-message-box')]"
            ).all()
            for mb in msg_boxes:
                if not mb.is_visible():
                    continue
                ok_btns = mb.locator(
                    "xpath=.//button[contains(@class,'el-button--primary')]"
                ).all()
                if ok_btns:
                    ok_btns[-1].evaluate("el => el.click()")
                    logging.info("已关闭 el-message-box 确认框")
                    self.page.wait_for_timeout(300)
                    return
                self.page.keyboard.press("Escape")
                self.page.wait_for_timeout(300)
        except Exception:
            pass

    def _close_pt_dropdown_safely(self, dialog, pt_wrapper=None):
        """收起 Parameter Type 下拉：再次点击 wrapper 触发 el-select toggle 关闭。
        绝不使用 Escape——Escape 会同时关闭 el-dialog（close-on-press-escape）。
        """
        if pt_wrapper is not None:
            try:
                expanded = pt_wrapper.get_attribute("aria-expanded", timeout=300) or "false"
                if expanded.lower() == "true":
                    pt_wrapper.click()
                    self.page.wait_for_timeout(300)
                    return
            except Exception:
                pass
        try:
            body = dialog.locator("xpath=.//div[contains(@class,'el-dialog__body')]").first
            body.click(position={"x": 5, "y": 5})
            self.page.wait_for_timeout(300)
        except Exception:
            pass

    def _get_visible_dialog(self):
        """返回当前可见的 el-dialog locator。"""
        self.page.wait_for_selector(
            "xpath=//div[contains(@class,'el-overlay-dialog')]//div[contains(@class,'el-dialog__body')]",
            state="visible",
            timeout=10000,
        )
        wrappers = self.page.locator(
            "xpath=//div[contains(@class,'el-overlay-dialog')]"
        ).all()
        for w in wrappers:
            if w.is_visible():
                dialog = w.locator("xpath=.//div[contains(@class,'el-dialog')]")
                if dialog.count() > 0:
                    return dialog.first
        raise RuntimeError("未找到可见的弹窗")

    def _configure_parameter_dialog(self):
        """
        在 Parameter Selection 弹窗中配置参数：
        - 有 Parameter Type 下拉（普通设备）：逐个类型全选
        - 无 Parameter Type 下拉（虚拟设备）：直接点 All 一次性全选
        最后点击 Confirm 关闭弹窗。
        """
        dialog = None
        try:
            dialog = self._get_visible_dialog()

            # 判断是否有 Parameter Type 下拉（普通设备）
            pt_wrappers = dialog.locator(
                "xpath=.//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Parameter Type')]]"
                "//div[contains(@class,'el-select__wrapper')]"
            ).all()

            if not pt_wrappers:
                # 虚拟设备：无分组，直接点击 All
                logging.info("  未检测到 Parameter Type 下拉（虚拟设备），直接点击 All")
                all_btns = dialog.locator(
                    "xpath=.//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]"
                ).all()
                if all_btns:
                    all_btns[0].evaluate("el => el.click()")
                    logging.info("  [All] → 已点击 All，全选参数")
                    self.page.wait_for_timeout(300)
                else:
                    logging.warning("  [All] → 未找到 All 按钮")
                return

            # 普通设备：逐 Parameter Type 全选
            pt_wrapper = pt_wrappers[0]

            # 打开下拉，采集所有选项文本（Playwright 原生 click，正确触发 Vue 事件）
            pt_wrapper.click()
            self.page.wait_for_timeout(500)

            all_opt_els = self.page.locator(
                "xpath=//li[contains(@class,'el-select-dropdown__item')]"
            ).all()
            option_texts = [
                o.inner_text().strip()
                for o in all_opt_els
                if o.is_visible() and o.inner_text().strip()
            ]
            logging.info(f"Parameter Type 选项：{option_texts}")

            # 不提前关闭下拉——i=0 直接使用已打开的 popper；i>0 由选项点击后自动关闭再重新打开
            for i, opt_text in enumerate(option_texts):
                self._dismiss_message_box()

                if i > 0:
                    pt_wrapper.click()
                    self.page.wait_for_timeout(800)

                clicked = False
                opt_scan = self.page.locator(
                    "xpath=//li[contains(@class,'el-select-dropdown__item')]"
                ).all()
                for o in opt_scan:
                    try:
                        if o.is_visible() and o.inner_text().strip() == opt_text:
                            o.click()
                            clicked = True
                            break
                    except Exception:
                        continue
                if not clicked:
                    logging.warning(f"  [{opt_text}] → 未找到可见选项，关闭下拉跳过")
                    self._close_pt_dropdown_safely(dialog, pt_wrapper)
                self.page.wait_for_timeout(500)

                # 若参数已全部选中，跳过
                if self._all_params_selected(dialog):
                    logging.info(f"  [{opt_text}] → 参数已全选，跳过")
                    continue

                # 点击 All 按钮（参数列表全选）
                all_btns = dialog.locator(
                    "xpath=.//button[normalize-space(.)='All' or .//span[normalize-space(.)='All']]"
                ).all()
                if all_btns:
                    all_btns[0].evaluate("el => el.click()")
                    logging.info(f"  [{opt_text}] → 已点击 All")
                    self.page.wait_for_timeout(500)
                    self._dismiss_message_box()
                else:
                    logging.warning(f"  [{opt_text}] → All 按钮未找到")

        except Exception as e:
            logging.error(f"配置 Parameter 弹窗时出错：{e}")
        finally:
            _CONF_XPATHS = [
                "xpath=.//div[contains(@class,'dialog-footer')]//button[contains(@class,'el-button--primary')][last()]",
                "xpath=.//div[contains(@class,'el-dialog__footer')]//button[contains(@class,'el-button--primary')][last()]",
                "xpath=.//footer//button[contains(@class,'el-button--primary')][last()]",
                "xpath=.//button[normalize-space(.)='Confirm']",
                "xpath=.//button[.//span[normalize-space(.)='Confirm']]",
            ]
            confirmed = False
            _search = None
            try:
                if dialog is not None and dialog.is_visible():
                    _search = dialog
            except Exception:
                pass
            # 若 dialog 引用失效，从 DOM 取最后一个可见弹窗
            if _search is None:
                try:
                    all_dlgs = self.page.locator(
                        "xpath=//div[@aria-label='AzureIoT Parameter Config']"
                    ).all()
                    for d in reversed(all_dlgs):
                        if d.is_visible():
                            _search = d
                            break
                except Exception:
                    pass
            if _search is not None:
                for conf_xpath in _CONF_XPATHS:
                    try:
                        btns = _search.locator(conf_xpath).all()
                        if btns:
                            btns[-1].evaluate("el => el.click()")
                            logging.info("已点击 Confirm，关闭弹窗")
                            confirmed = True
                            break
                    except Exception:
                        continue
            # ESC 兜底
            _dlg_still_open = False
            if confirmed and _search is not None:
                try:
                    _dlg_still_open = _search.is_visible()
                except Exception:
                    _dlg_still_open = True
            if not confirmed or _dlg_still_open:
                try:
                    self.page.keyboard.press("Escape")
                    logging.info("已按 ESC 关闭弹窗")
                except Exception:
                    logging.warning("无法关闭弹窗，继续执行")
            # 等待弹窗消失（最多 3 秒）
            try:
                deadline = time.time() + 3
                while time.time() < deadline:
                    dlgs = self.page.locator(
                        "xpath=//div[@aria-label='AzureIoT Parameter Config']"
                    ).all()
                    if not any(d.is_visible() for d in dlgs):
                        logging.info("弹窗已关闭")
                        break
                    self.page.wait_for_timeout(200)
                else:
                    logging.warning("弹窗关闭等待超时，尝试继续执行")
            except Exception:
                pass
            self.page.wait_for_timeout(200)

    def configure_all_devices_parameters(self, checked_only: bool = False):
        """
        遍历设备列表，点击 Parameter Selection 按钮配置参数。
        checked_only=True 时只处理已勾选的设备行。
        """
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        logging.info(f"设备列表共 {len(rows)} 行，逐行配置参数选择（checked_only={checked_only}）")
        for idx, row in enumerate(rows):
            if checked_only:
                cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
                if not cb_labels or 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                    logging.info(f"  第 {idx + 1} 行：未勾选，跳过参数配置")
                    continue
            btns = row.locator("xpath=.//button").all()
            if not btns:
                logging.warning(f"  第 {idx + 1} 行未找到按钮，跳过")
                continue
            # 优先找包含 "Select" / "Parameter" 文字的按钮
            target_btn = None
            for b in btns:
                txt = (b.inner_text() or '').strip()
                if any(kw in txt for kw in ('Select', 'Parameter', '选择', '参数')):
                    target_btn = b
                    break
            if target_btn is None:
                target_btn = btns[-1]
            label = (target_btn.inner_text() or '').strip()
            logging.info(f"  第 {idx + 1} 行：点击按钮「{label or '(icon)'}」")
            target_btn.evaluate("el => el.scrollIntoView({block:'center'})")
            self.page.wait_for_timeout(300)
            target_btn.evaluate("el => el.click()")
            self.page.wait_for_timeout(1500)
            self._configure_parameter_dialog()

    def save(self):
        btn = self.page.locator(self.SAVE_BTN)
        btn.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)
        btn.evaluate("el => el.click()")
        self.page.wait_for_timeout(3000)
        logging.info(f"已点击 Save，当前 URL：{self.page.url}")

    def test_connection(self) -> str:
        """点击 Test Connection 并等待 Test Result 出现，返回提示文字。"""
        self.page.wait_for_selector(self.TEST_CONNECTION_BTN, state="visible", timeout=30000)
        btn = self.page.locator(self.TEST_CONNECTION_BTN)
        btn.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)
        btn.evaluate("el => el.click()")
        logging.info("已点击 Test Connection，等待 Test Result…")

        # 轮询最多 180 秒，每秒检查一次
        deadline = time.time() + 180
        while time.time() < deadline:
            self.page.wait_for_timeout(1000)
            # 查找含 "Test Result" 的元素
            candidates = self.page.locator(
                "xpath=//*[contains(normalize-space(.),'Test Result')]"
            ).all()
            for el in candidates:
                try:
                    txt = el.inner_text().strip()
                    if not txt or "test result" not in txt.lower():
                        continue
                    if len(txt) > 200:
                        continue
                    if "connecting" in txt.lower():
                        continue
                    logging.info(f"Test Connection 结果：{txt}")
                    return txt
                except Exception:
                    continue
            # 兜底：Element UI 标准消息
            result_els = self.page.locator(self.RESULT_MSG).all()
            for el in result_els:
                try:
                    txt = el.inner_text().strip()
                    if (txt
                            and "connecting" not in txt.lower()
                            and "loading" not in txt.lower()
                            and "save" not in txt.lower()):
                        logging.info(f"Test Connection 结果（兜底）：{txt}")
                        return txt
                except Exception:
                    continue
        logging.warning("未在 180 秒内检测到 Test Result 弹窗")
        return "未检测到结果提示"

    def get_checked_device_names(self) -> list:
        """返回设备表格中所有已勾选行的设备名称列表。"""
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()
        names = []
        for row in rows:
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                continue
            if 'is-checked' not in (cb_labels[0].get_attribute('class') or ''):
                continue
            # 找到不含 button / input 的文本单元格作为设备名
            for td in row.locator("xpath=.//td").all():
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split('\n')[0].strip()
                if text:
                    names.append(text)
                    break
        return names

    def select_only_device(self, device_name: str = ""):
        """取消所有已勾选设备，再单独勾选名称匹配的设备。
        device_name 为空时，自动从列表中随机选取一台物理（非虚拟）设备。"""
        import random
        rows = self.page.locator(self.DEVICE_TABLE_ROWS).all()

        if not device_name:
            physical_names = []
            for row in rows:
                cells = row.locator("xpath=.//td").all()
                if any("Virtual" in (c.inner_text() or "") for c in cells):
                    continue
                for td in cells:
                    if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                        continue
                    text = td.inner_text().strip().split("\n")[0].strip()
                    if text:
                        physical_names.append(text)
                        break
            if not physical_names:
                logging.warning("select_only_device: 未找到可用的物理设备，跳过选择")
                return
            device_name = random.choice(physical_names)
            logging.info(f"select_only_device: 自动随机选取物理设备 {device_name!r}")
            print(f"\n[Setup] 自动随机选取物理设备: {device_name!r}")

        logging.info(f"select_only_device: 目标={device_name!r}，共 {len(rows)} 行")
        for idx, row in enumerate(rows):
            cb_labels = row.locator("xpath=.//label[contains(@class,'el-checkbox')]").all()
            if not cb_labels:
                continue
            cb = cb_labels[0]
            is_checked = "is-checked" in (cb.get_attribute("class") or "")
            row_name = ""
            for td in row.locator("xpath=.//td").all():
                if td.locator("xpath=.//*[self::button or self::input]").count() > 0:
                    continue
                text = td.inner_text().strip().split("\n")[0].strip()
                if text:
                    row_name = text
                    break
            should_check = device_name.lower() in row_name.lower()
            if should_check != is_checked:
                cb.evaluate("el => el.scrollIntoView({block:'center'})")
                self.page.wait_for_timeout(200)
                cb.evaluate("el => el.click()")
                action = "勾选" if should_check else "取消"
                logging.info(f"  第 {idx+1} 行 [{row_name}]：{action}")
            else:
                logging.info(f"  第 {idx+1} 行 [{row_name}]：无需操作")

    def disable(self):
        """点击 Disable 单选并保存，关闭 Azure IoT 推送。已是 Disable 状态则跳过。"""
        # 先关闭可能残留的弹窗
        try:
            confirm_btn = self.page.locator(
                "xpath=//div[contains(@class,'el-overlay-message-box')]//button[contains(@class,'el-button--primary')]"
            )
            if confirm_btn.count() > 0 and confirm_btn.first.is_visible():
                confirm_btn.first.evaluate("el => el.click()")
                self.page.wait_for_timeout(500)
        except Exception:
            pass
        if self.is_enabled():
            logging.info("切换 Azure IoT 为 Disable")
            self.page.locator(self.DISABLE_RADIO).evaluate("el => el.click()")
            self.page.wait_for_timeout(500)
            self.save()
            logging.info("Azure IoT 已 Disable 并保存")
        else:
            logging.info("Azure IoT 已是 Disable 状态，跳过")

    def configure(self, primary_conn_str: str, secondary_conn_str: str = '',
                  interval: str = '30 seconds') -> str:
        """
        完整配置流程：
        1. 确保 Enable 已勾选
        2. 填写连接参数
        3. 全选设备
        4. 保存
        5. 测试连接并返回结果
        """
        self.ensure_enabled()
        self.set_primary_conn_str(primary_conn_str)
        if secondary_conn_str:
            self.set_secondary_conn_str(secondary_conn_str)
        self.set_interval(interval)
        self.select_unchecked_devices()
        self.configure_all_devices_parameters()
        self.save()
        result = self.test_connection()
        return result

    def get_virtual_device_readings(self, device_name: str) -> dict:
        """
        导航到 Devices → Virtual Devices → 指定设备 → Reading 页面，
        读取所有参数名和对应值/单位。

        返回: {param_name: {"value": str, "unit": str}}
        """
        # 1. 进入 Devices 视图
        devices_btn = self.page.locator(self.DEVICES_NAV_BTN)
        self.page.wait_for_selector(self.DEVICES_NAV_BTN, state="visible", timeout=10000)
        devices_btn.evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)

        # 2. 点击 Virtual Devices
        self.page.wait_for_selector(self.VIRTUAL_DEVICES_MENU, state="visible", timeout=10000)
        self.page.locator(self.VIRTUAL_DEVICES_MENU).first.evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)

        # 3. 点击目标设备名
        _dev_xpaths = [
            f"xpath=//div[contains(@class,'link-url') and normalize-space(.)='{device_name}']",
            f"xpath=//td[contains(normalize-space(.),'{device_name}')]//div[contains(@class,'link-url')]",
            f"xpath=//td[contains(normalize-space(.),'{device_name}')]//a",
            f"xpath=//a[normalize-space(.)='{device_name}']",
        ]
        dev_link = None
        deadline = time.time() + 10
        while time.time() < deadline:
            for xp in _dev_xpaths:
                els = self.page.locator(xp).all()
                if els:
                    dev_link = els[0]
                    break
            if dev_link is not None:
                break
            self.page.wait_for_timeout(500)
        if dev_link is None:
            raise RuntimeError(f"未找到设备 {device_name!r} 的链接")

        dev_link.evaluate("el => el.scrollIntoView({block:'center'})")
        self.page.wait_for_timeout(300)
        dev_link.evaluate("el => el.click()")
        self.page.wait_for_timeout(2500)

        # 4. 点击 Reading 标签
        _reading_xpaths = [
            "xpath=//*[contains(@class,'el-tabs__item') and contains(normalize-space(.),'Reading')]",
            "xpath=//li[contains(@class,'tab') and contains(normalize-space(.),'Reading')]",
            "xpath=//div[contains(@class,'tab') and contains(normalize-space(.),'Reading')]",
            "xpath=//*[@role='tab' and contains(normalize-space(.),'Reading')]",
            "xpath=//span[normalize-space(.)='Reading']/ancestor::*[@role='tab' or contains(@class,'tab')][1]",
        ]
        reading_tab = None
        deadline = time.time() + 15
        while time.time() < deadline:
            for xp in _reading_xpaths:
                els = self.page.locator(xp).all()
                if els:
                    reading_tab = els[0]
                    break
            if reading_tab is not None:
                break
            self.page.wait_for_timeout(500)
        if reading_tab is None:
            raise RuntimeError("未找到 Reading 标签")

        reading_tab.evaluate("el => el.click()")
        self.page.wait_for_timeout(1500)

        # 5. 读取表格
        params = {}
        rows = self.page.locator(self.READING_TABLE_ROWS).all()
        for row in rows:
            cells = row.locator("xpath=.//td").all()
            if len(cells) < 2:
                continue
            param_name = cells[0].inner_text().strip()
            raw_val = cells[1].inner_text().strip()
            if not param_name:
                continue
            parts = raw_val.split(None, 1)
            params[param_name] = {
                "value": parts[0] if parts else "",
                "unit": parts[1] if len(parts) > 1 else "",
            }
        logging.info(f"get_virtual_device_readings({device_name!r}): 读取到 {len(params)} 个参数")
        return params
