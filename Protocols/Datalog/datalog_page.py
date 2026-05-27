# -*- coding: utf-8 -*-
"""
datalog_page.py — DataLog 页面自动化（基于截图实际 UI 结构）

导航路径：
  左侧导航 Data Log
    → 顶部 tab "Post Channels ▼" 下拉 → Post Channel 1 / 2 / 3
    → 顶部 tab "Data Loggers ▼"  下拉 → Data Logger  1 / 2 / 3

Post Channel 字段（按 Post Method 不同动态显示）：
  ┌ FTP
  │   FTP URL（FTP:// + IP/域名）、FTP Port、Enable anonymous mode、
  │   FTP User Name、FTP password
  ├ SFTP
  │   SFTP URL（SFTP:// + IP/域名）、SFTP Port、
  │   SFTP User Name、SFTP password
  └ HTTP/HTTPS
      Post Name Fixed（Yes/No）、Post File Name、
      Authentication Required（Yes/No）、
      HTTP/HTTPS URL（HTTP:// 或 HTTPS:// + IP/域名）、HTTP/HTTPS Port、
      HTTP/HTTPS Meter ID、Include Header（Yes/No）
"""
import logging
import time
from dataclasses import dataclass, field
from typing import Optional

from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait

from comm.ui_base_page import BasePage

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 配置数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class PostChannelConfig:
    """
    Post Channel 配置参数。

    protocol: 内部协议标识 "FTP" | "SFTP" | "HTTP" | "HTTPS"
      - "FTP"  / "SFTP"  → 填 FTP / SFTP 字段
      - "HTTP" / "HTTPS" → UI 下拉选 "HTTP/HTTPS"，scheme 决定 URL 前缀
    """
    protocol: str          # "FTP" | "SFTP" | "HTTP" | "HTTPS"
    host: str              # IP 或域名（不含协议前缀）
    port: int

    # FTP / SFTP 共用
    username: str = ""
    password: str = ""

    # FTP 专用
    anonymous_mode: bool = False

    # HTTP/HTTPS 专用
    post_name_fixed: bool = False   # Post Name Fixed: Yes/No
    post_file_name: str = ""        # Post File Name（固定文件名时填）
    auth_required: bool = False     # Authentication Required: Yes/No
    meter_id: str = ""              # HTTP/HTTPS Meter ID
    include_header: bool = True     # Include Header: Yes/No

    @property
    def ui_method(self) -> str:
        """UI Post Method 下拉的显示文字。"""
        return "HTTP/HTTPS" if self.protocol in ("HTTP", "HTTPS") else self.protocol

    @property
    def http_scheme(self) -> str:
        """HTTP/HTTPS URL 前缀下拉的显示文字。"""
        return "HTTPS://" if self.protocol == "HTTPS" else "HTTP://"


@dataclass
class DataLoggerConfig:
    """
    Data Logger 配置参数（对应截图实际字段）。

    channel_index      : 关联 Post Channel 编号（1 / 2 / 3）
    timestamp_format   : Timestamp Format radio
                         "Local Time String" | "UTC Seconds" | "ISO8601 Format"
    log_file_name_format: Log File Name Format radio
                         "UTC Timestamp" | "Time Interval Format"
    log_file_format    : Log File Format 下拉，"csv" 或 "json"
    log_file_name_prefix: Log File Name Prefix 输入框（留空不修改）
    log_file_length    : Log File Length 下拉文字，如 "1 minute"
    log_interval       : Log Interval 下拉文字，如 "1 minute"
    device_names       : 勾选的设备名列表，空列表 = 全选（点击 header checkbox）
    """
    channel_index: int
    enabled: bool = True
    timestamp_format: str = "UTC Seconds"             # Timestamp Format
    log_file_name_format: str = "Time Interval Format" # Log File Name Format
    log_file_format: str = "csv"                      # Log File Format 下拉
    log_file_name_prefix: str = ""                    # 留空不修改
    log_file_length: str = "5 minute"                 # Log File Length
    log_interval: str = "1 minute"                    # Log Interval
    device_names: list = field(default_factory=list)  # 空 = 全选


@dataclass
class DataLogParamConfig:
    """
    Data Log Parameter Config 页面的配置。

    device_names : 要配置的设备名列表（与 UI Device 下拉选项文字精确匹配）
    param_types  : 要选择的参数类型列表（与 UI Parameter Type 选项文字精确匹配）
                   空列表 = 遍历下拉中所有可用类型并全选
    """
    device_names: list = field(default_factory=list)
    param_types: list = field(default_factory=list)   # 空 = 全部类型全选


# ─────────────────────────────────────────────────────────────────────────────
# 基础操作 mixin
# ─────────────────────────────────────────────────────────────────────────────

class _PageBase(BasePage):
    """共用的底层操作。"""

    # ── 顶部右侧设备菜单（Element Plus nav-item-menu，与 AWSIoTPage 一致） ──────
    TOP_DEVICE_BTN = (By.XPATH,
        "//div[contains(@class,'nav-item-menu')][.//span[normalize-space(.)='AcuHMI-1-7']]")

    # 左侧导航 Data Log（Element Plus left-nav-item）
    DATA_LOG_MENU = (By.XPATH,
        "//li[contains(@class,'left-nav-item') and contains(normalize-space(.),'Data Log')]")

    def _js_click(self, el):
        self.driver.execute_script("arguments[0].click()", el)

    def _safe_click(self, locator, name: str = ""):
        try:
            el = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(locator))
            self._js_click(el)
            log.debug("点击：%s", name)
        except TimeoutException:
            log.warning("元素未找到：%s  %s", name, locator)

    def _fill(self, locator, value: str, name: str = ""):
        """清空并输入文字（用 presence 定位，JS click 后 Ctrl+A 替换，避免下拉遮挡）。"""
        try:
            el = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            self._js_click(el)
            el.send_keys(Keys.CONTROL, 'a')
            el.send_keys(str(value))
            log.debug("填写 %s = %s", name, value)
        except TimeoutException:
            log.warning("输入框未找到：%s  %s", name, locator)

    def _select_by_text(self, locator, text: str, name: str = ""):
        """操作 <select> 原生下拉，按文字选择。"""
        try:
            el = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            Select(el).select_by_visible_text(text)
            log.debug("下拉 %s 选择：%s", name, text)
            time.sleep(0.3)
        except (TimeoutException, NoSuchElementException) as e:
            log.warning("下拉 %s 选 '%s' 失败：%s", name, text, e)

    def _select_el_by_text(self, trigger_locator, text: str, name: str = ""):
        """
        操作 Element Plus el-select 容器 div，点击展开后再选目标 option。
        失败时按 Escape 关闭残留下拉框，避免遮挡后续元素。
        """
        try:
            container = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(trigger_locator))
            # 优先点击 el-select__wrapper（内层可点击区域），回退到容器本身
            inner = container.find_elements(By.XPATH,
                ".//div[contains(@class,'el-select__wrapper')]")
            self._js_click(inner[0] if inner else container)
            time.sleep(0.5)
            # 尝试多种 XPath 匹配下拉选项（Element Plus v2 teleports to body）
            opt_patterns = [
                f"//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{text}']",
                f"//li[contains(@class,'el-select-dropdown__item') and .//span[normalize-space(.)='{text}']]",
                f"//ul[contains(@class,'el-select-dropdown__list')]//li[contains(normalize-space(.),'{text}')]",
                f"//div[contains(@class,'el-popper')]//li[contains(normalize-space(.),'{text}')]",
            ]
            for pat in opt_patterns:
                try:
                    opt = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, pat)))
                    self._js_click(opt)
                    log.debug("el-select %s 选择：%s", name, text)
                    time.sleep(0.4)
                    return
                except (TimeoutException, AttributeError):
                    continue
            # 选项未找到 — 记录实际可见选项以辅助调试
            try:
                visible_opts = self.driver.find_elements(By.XPATH,
                    "//li[contains(@class,'el-select-dropdown__item')]")
                opt_texts = [o.text.strip() for o in visible_opts if o.text.strip()]
                log.warning("el-select %s 选 '%s' 失败，可见选项：%s，按 Escape 关闭下拉",
                            name, text, opt_texts)
            except Exception:
                log.warning("el-select %s 选 '%s' 失败，按 Escape 关闭下拉", name, text)
        except (TimeoutException, AttributeError) as e:
            log.warning("el-select %s 容器未找到（'%s'）：%s", name, text, e)
        finally:
            # 无论成功/失败都尝试关闭可能残留的下拉框
            try:
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.2)
            except Exception:
                pass

    def _click_radio(self, label_text: str, context_el=None):
        """点击包含 label_text 文字的 radio 标签或其对应的 input。"""
        root = context_el or self.driver
        for xpath in [
            # Element Plus 风格：<label class="el-radio"><span>text</span></label>
            f".//label[.//span[normalize-space(.)='{label_text}']]",
            # 标准 HTML：<label>text</label>
            f".//label[normalize-space(.)='{label_text}']",
            f".//label[contains(normalize-space(.),'{label_text}')]",
            f".//span[normalize-space(.)='{label_text}']/ancestor::label[1]",
            f".//input[@type='radio'][following-sibling::*"
            f"[normalize-space(.)='{label_text}']]",
        ]:
            els = root.find_elements(By.XPATH, xpath)
            if els:
                try:
                    self._js_click(els[0])
                    log.debug("单选：%s", label_text)
                    time.sleep(0.3)
                    return
                except Exception:
                    continue
        log.warning("未找到 radio：%s", label_text)

    def _dismiss_overlay_dialog(self):
        """关闭登录后可能残留的 el-overlay-message-box 确认弹窗。"""
        from selenium.common.exceptions import StaleElementReferenceException
        try:
            dialogs = self.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-overlay-message-box') "
                "or contains(@class,'el-message-box')]")
            for dlg in dialogs:
                try:
                    if not dlg.is_displayed():
                        continue
                except StaleElementReferenceException:
                    continue
                for xpath in [
                    ".//button[.//span[normalize-space(.)='Yes']]",
                    ".//button[.//span[normalize-space(.)='Continue']]",
                    ".//button[contains(@class,'el-button--primary')]",
                    ".//button[contains(@class,'el-button')]",
                ]:
                    try:
                        btns = dlg.find_elements(By.XPATH, xpath)
                        if btns:
                            self._js_click(btns[0])
                            log.info("已关闭确认弹窗")
                            time.sleep(0.5)
                            break
                    except StaleElementReferenceException:
                        break
        except Exception as e:
            log.debug("dismiss_overlay_dialog: %s", e)

    def _navigate_data_log(self):
        """关闭登录后弹窗，为后续 URL 直达导航做准备。"""
        self._dismiss_overlay_dialog()
        time.sleep(0.8)

    def _navigate_url(self, url_hash: str):
        """两步导航：先到 Data Log 根路由建立 Vue 上下文，再跳转目标子页面。"""
        current = self.driver.current_url
        base = current.split('#')[0]
        # Step 1: 始终先进入 Data Log 根路由，确保 Vue router 完成该模块的初始化
        self.driver.get(f"{base}#/dataLog")
        time.sleep(1.5)
        # Step 2: 再跳转目标子页面
        self.driver.get(f"{base}#{url_hash}")
        time.sleep(2.0)
        actual = self.driver.current_url
        log.info("URL 导航完成：%s", actual)

    def _click_radio_in_group(self, group_label: str, option_text: str) -> bool:
        """
        在包含 group_label 的表单项后，点击文字含 option_text 的 el-radio label。
        Element Plus 将 radio input 隐藏，必须点击外层 el-radio label 而非 input。
        返回 True 表示成功找到并点击，False 表示未找到。
        """
        for xpath in [
            # Element Plus: el-radio label 跟随 form-item label
            f"//label[contains(@class,'el-form-item__label') and contains(normalize-space(.), '{group_label}')]"
            f"/following::label[contains(@class,'el-radio') and normalize-space(.)='{option_text}'][1]",
            f"//label[contains(@class,'el-form-item__label') and contains(normalize-space(.), '{group_label}')]"
            f"/following::label[contains(@class,'el-radio') and contains(normalize-space(.), '{option_text}')][1]",
            # 同一 form-item 容器内
            f"//*[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.), '{group_label}')]]"
            f"//label[contains(@class,'el-radio') and contains(normalize-space(.), '{option_text}')]",
            # 兜底：找包含 el-radio__original value 的 label
            f"//label[contains(@class,'el-form-item__label') and contains(normalize-space(.), '{group_label}')]"
            f"/following::label[contains(@class,'el-radio')][.//span[normalize-space(.)='{option_text}']][1]",
        ]:
            els = self.driver.find_elements(By.XPATH, xpath)
            if els:
                self._js_click(els[0])
                log.debug("  Radio [%s] → %s", group_label, option_text)
                time.sleep(0.3)
                return True
        log.warning("未找到 radio 组 '%s' 中的选项 '%s'", group_label, option_text)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Post Channel 页面
# ─────────────────────────────────────────────────────────────────────────────

class PostChannelPage(_PageBase):
    """
    配置 Post Channel 1 / 2 / 3。
    每个 Channel 均可独立选择 FTP / SFTP / HTTP/HTTPS。
    """

    # ── 顶部 tab：Post Channels（Element Plus el-menu-item / el-sub-menu）────────
    POST_CHANNELS_TAB = (By.XPATH,
        "//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Post Channels')]"
        " | //div[contains(@class,'el-sub-menu__title') and contains(normalize-space(.),'Post Channels')]")

    @staticmethod
    def _channel_item(n: int):
        """Post Channel N 菜单项（Element Plus el-menu-item）。"""
        return (By.XPATH,
            f"//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Post Channel {n}')]")

    # ── Post Channel 表单字段 ─────────────────────────────────────────────────

    # Enable / Disable — 点击 el-radio label（实际 input 被 Element Plus 隐藏）
    ENABLE_RADIO  = (By.XPATH, "//label[contains(@class,'el-radio') and .//input[@value='true']]")
    DISABLE_RADIO = (By.XPATH, "//label[contains(@class,'el-radio') and .//input[@value='false']]")

    # Post Method — el-select（点击触发 div 后选 option）
    POST_METHOD_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and normalize-space(.)='Post Method']"
        "/following::div[contains(@class,'el-select')][1]")

    # ── FTP 字段（placeholder 定位）────────────────────────────────────────────
    FTP_URL_INPUT  = (By.XPATH, "//input[@placeholder='Enter FTP URL']")
    FTP_PORT_INPUT = (By.XPATH, "//input[@placeholder='Enter FTP Port']")
    FTP_ANON_CB    = (By.XPATH, "//input[@type='checkbox']")
    FTP_USER_INPUT = (By.XPATH, "//input[@placeholder='Enter FTP User Name']")
    FTP_PASS_INPUT = (By.XPATH, "//input[@placeholder='Enter FTP Password']")

    # ── SFTP 字段（placeholder 定位）──────────────────────────────────────────
    SFTP_URL_INPUT  = (By.XPATH, "//input[@placeholder='Enter SFTP URL']")
    SFTP_PORT_INPUT = (By.XPATH, "//input[@placeholder='Enter SFTP Port']")
    SFTP_USER_INPUT = (By.XPATH, "//input[@placeholder='Enter SFTP User Name']")
    SFTP_PASS_INPUT = (By.XPATH, "//input[@placeholder='Enter SFTP Password']")

    # ── HTTP/HTTPS 字段（placeholder 定位）────────────────────────────────────
    # Post Name Fixed / Auth Required / Include Header → el-radio（Yes/No），非 el-select
    # HTTP/HTTPS URL 协议前缀 → el-select，通过 URL 输入框的 preceding 关系定位
    HTTP_SCHEME_SELECT = (By.XPATH,
        "//input[@placeholder='Enter HTTP/HTTPS URL']/preceding::div[contains(@class,'el-select')][1]")
    HTTP_URL_INPUT       = (By.XPATH, "//input[@placeholder='Enter HTTP/HTTPS URL']"
                                       " | //input[@placeholder='Enter URL']")
    HTTP_PORT_INPUT      = (By.XPATH, "//input[@placeholder='Enter HTTP/HTTPS Port']"
                                       " | //input[@placeholder='Enter Port']")
    HTTP_METER_ID_INPUT  = (By.XPATH, "//input[@placeholder='Enter Meter ID']"
                                       " | //input[@placeholder='Enter HTTP/HTTPS Meter ID']")
    HTTP_FILE_NAME_INPUT = (By.XPATH, "//input[@placeholder='Enter Post File Name']"
                                       " | //input[@placeholder='Enter File Name']")

    # ── 按钮 ──────────────────────────────────────────────────────────────────
    TEST_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Test Post Channel']]"
        " | //button[normalize-space(.)='Test Post Channel']")
    SAVE_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Save']]"
        " | //button[normalize-space(.)='Save']")

    # ── 导航 ─────────────────────────────────────────────────────────────────

    def navigate_to_channel(self, n: int):
        """导航到 Post Channel n（1/2/3）页面。"""
        self._navigate_data_log()
        self._navigate_url(f"/dataLog/postChannels/postChannel{n}")
        log.info("已进入 Post Channel %d 页面", n)

    # ── 字段填写 ──────────────────────────────────────────────────────────────

    def _set_enable(self, enabled: bool):
        """点击 el-radio 标签（实际 input 被 Element Plus 隐藏，用 presence 定位）。"""
        locator = self.ENABLE_RADIO if enabled else self.DISABLE_RADIO
        try:
            el = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            self._js_click(el)
            log.debug("设置 Enable=%s", enabled)
        except TimeoutException:
            log.warning("Enable/Disable radio 未找到，locator=%s", locator)

    def _set_post_method(self, ui_method: str):
        """选择 Post Method（el-select）。"""
        self._select_el_by_text(self.POST_METHOD_SELECT, ui_method, "Post Method")
        time.sleep(0.8)   # 等待字段切换动画

    def _fill_ftp(self, cfg: PostChannelConfig):
        self._fill(self.FTP_URL_INPUT,  cfg.host,        "FTP URL")
        self._fill(self.FTP_PORT_INPUT, str(cfg.port),   "FTP Port")
        if cfg.anonymous_mode:
            try:
                cb = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located(self.FTP_ANON_CB))
                if not cb.is_selected():
                    self._js_click(cb)
                log.info("已勾选 Enable anonymous mode")
            except TimeoutException:
                log.warning("Enable anonymous mode 复选框未找到")
        else:
            self._fill(self.FTP_USER_INPUT, cfg.username, "FTP User Name")
            self._fill(self.FTP_PASS_INPUT, cfg.password, "FTP Password")

    def _fill_sftp(self, cfg: PostChannelConfig):
        self._fill(self.SFTP_URL_INPUT,  cfg.host,        "SFTP URL")
        self._fill(self.SFTP_PORT_INPUT, str(cfg.port),   "SFTP Port")
        self._fill(self.SFTP_USER_INPUT, cfg.username,    "SFTP User Name")
        self._fill(self.SFTP_PASS_INPUT, cfg.password,    "SFTP Password")

    def _fill_http(self, cfg: PostChannelConfig):
        # Post Name Fixed — el-radio（Yes / No）
        self._click_radio_in_group("Post Name Fixed",
                                   "Yes" if cfg.post_name_fixed else "No")
        time.sleep(0.3)
        if cfg.post_name_fixed and cfg.post_file_name:
            self._fill(self.HTTP_FILE_NAME_INPUT, cfg.post_file_name, "Post File Name")

        # Authentication Required — el-radio（Yes / No）
        self._click_radio_in_group("Authentication Required",
                                   "Yes" if cfg.auth_required else "No")

        # HTTP/HTTPS URL — 先选协议前缀（el-select，通过 URL 输入框 preceding 定位），再填 host
        self._select_el_by_text(self.HTTP_SCHEME_SELECT, cfg.http_scheme, "HTTP/HTTPS scheme")
        self._fill(self.HTTP_URL_INPUT,  cfg.host,           "HTTP/HTTPS URL")
        self._fill(self.HTTP_PORT_INPUT, str(cfg.port),      "HTTP/HTTPS Port")
        self._fill(self.HTTP_METER_ID_INPUT, cfg.meter_id or "", "HTTP/HTTPS Meter ID")

        # Include Header — el-radio（Yes / No）
        self._click_radio_in_group("Include Header",
                                   "Yes" if cfg.include_header else "No")

    def _click_yes_no_after_label(self, label_text: str, choose_yes: bool):
        """
        找到包含 label_text 的 label，点击其后第一个 Yes 或 No radio。
        """
        target = "Yes" if choose_yes else "No"
        # 找到 label 元素所在的 form-group/row 容器，再在其内找 Yes/No
        for container_xpath in [
            f"//label[contains(normalize-space(.), '{label_text}')]"
            f"/following-sibling::*[1]",
            f"//label[contains(normalize-space(.), '{label_text}')]"
            f"/parent::*/following-sibling::*[1]",
            f"//*[contains(@class,'form-group') or contains(@class,'row')]"
            f"[.//label[contains(normalize-space(.), '{label_text}')]]",
        ]:
            containers = self.driver.find_elements(By.XPATH, container_xpath)
            for container in containers:
                for radio_xpath in [
                    f".//label[normalize-space(.)='{target}']",
                    f".//label[contains(normalize-space(.), '{target}')]",
                    f".//input[@type='radio'][following-sibling::*[normalize-space(.)='{target}']]",
                ]:
                    els = container.find_elements(By.XPATH, radio_xpath)
                    if els:
                        self._js_click(els[0])
                        log.debug("  %s → %s", label_text, target)
                        time.sleep(0.3)
                        return
        # 兜底：在整页找 label 后最近的 Yes/No
        log.warning("未能精确定位 %s 的 %s radio，尝试兜底", label_text, target)
        for xpath in [
            f"//label[contains(normalize-space(.), '{label_text}')]"
            f"/following::label[normalize-space(.)='{target}'][1]",
            f"//label[contains(normalize-space(.), '{label_text}')]"
            f"/following::input[@type='radio']"
            f"[following-sibling::*[normalize-space(.)='{target}']][1]",
        ]:
            els = self.driver.find_elements(By.XPATH, xpath)
            if els:
                self._js_click(els[0])
                log.debug("  兜底 %s → %s", label_text, target)
                time.sleep(0.3)
                return

    # ── 主流程 ────────────────────────────────────────────────────────────────

    def configure_channel(self, n: int, cfg: PostChannelConfig,
                          enabled: bool = True, test: bool = True) -> str:
        """
        配置第 n 个 Post Channel，返回 Test Post Channel 结果文字（未测试时为空串）。

        当 enabled=False 时：先以 Enable 状态填写所有字段并保存（确保字段可见），
        再切换到 Disabled 二次保存。这样服务器信息被记住，后续只需 enable_channel() 即可。
        """
        log.info("配置 Post Channel %d：%s  %s:%d", n, cfg.protocol, cfg.host, cfg.port)
        self.navigate_to_channel(n)

        # 始终先 Enable，确保所有协议字段可见可填
        self._set_enable(True)
        self._set_post_method(cfg.ui_method)

        if cfg.protocol == "FTP":
            self._fill_ftp(cfg)
        elif cfg.protocol == "SFTP":
            self._fill_sftp(cfg)
        else:  # HTTP / HTTPS
            self._fill_http(cfg)

        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(2)
        log.info("Post Channel %d 已保存（Enabled）", n)

        result = ""
        if test and enabled:
            result = self._test_channel()

        if not enabled:
            # 配置完成后 Disable，让用例按需 enable_channel() 启用
            self._set_enable(False)
            self._safe_click(self.SAVE_BTN, "Save")
            time.sleep(1.5)
            log.info("Post Channel %d 已禁用（配置保留）", n)

        return result

    def enable_channel(self, n: int):
        """
        仅启用 Post Channel n（不修改其他配置），用于用例执行前按需开启。
        """
        log.info("启用 Post Channel %d", n)
        self.navigate_to_channel(n)
        self._set_enable(True)
        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(1.5)
        log.info("Post Channel %d 已启用", n)

    def disable_channel(self, n: int):
        """
        仅禁用 Post Channel n（不修改其他配置）。
        """
        log.info("禁用 Post Channel %d", n)
        self.navigate_to_channel(n)
        self._set_enable(False)
        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(1.5)
        log.info("Post Channel %d 已禁用", n)

    def _test_channel(self) -> str:
        """
        点击 Test Post Channel，等待弹窗结果：
        - 成功（含 success / Success / connected）：关闭弹窗后返回结果文字。
        - 失败（含 fail / Fail / error / Error）：关闭弹窗后抛出 RuntimeError，终止测试。
        - 30 秒内未出现弹窗：记录警告后返回空串（不中断）。

        处理顺序：① 原生 browser alert → ② el-message-box → ③ toast/notification
        """
        from selenium.common.exceptions import UnexpectedAlertPresentException
        from selenium.webdriver.support.ui import WebDriverWait

        self._safe_click(self.TEST_BTN, "Test Post Channel")
        log.info("已点击 Test Post Channel，等待结果…")
        time.sleep(2)

        # ── ① 原生浏览器 Alert ────────────────────────────────────────────────
        try:
            alert = self.driver.switch_to.alert
            result = alert.text.strip()
            log.info("Test Post Channel 原生 Alert：%s", result)
            result_lower = result.lower()
            is_fail = any(kw in result_lower for kw in ("fail", "error", "failed"))
            alert.accept()
            log.info("已点击原生 Alert OK")
            time.sleep(0.5)
            if is_fail:
                raise RuntimeError(
                    f"Test Post Channel 失败，请检查服务器配置。Alert 内容：{result!r}"
                )
            return result
        except Exception as e:
            if "RuntimeError" in type(e).__name__:
                raise
            pass  # 无原生 Alert，继续尝试 HTML 弹窗

        # ── ② Element Plus el-message-box ────────────────────────────────────
        msgbox_locator = (By.XPATH,
            "//div[contains(@class,'el-message-box__wrapper') and not(contains(@style,'display: none'))]"
            " | //div[@role='dialog' and contains(@class,'el-overlay')]"
        )
        try:
            box = WebDriverWait(self.driver, 28).until(
                EC.presence_of_element_located(msgbox_locator))
            result = box.text.strip()
            log.info("Test Post Channel 弹窗内容：%s", result)

            result_lower = result.lower()
            is_success = any(kw in result_lower for kw in
                             ("success", "connected", "test success", "pass"))
            is_fail = any(kw in result_lower for kw in
                          ("fail", "error", "failed", "test fail"))

            # 如有 Detail 按钮，先点击获取详细错误信息
            detail_text = ""
            try:
                detail_btn = box.find_element(By.XPATH,
                    ".//button[normalize-space(.)='Detail']"
                    " | .//button[.//span[normalize-space(.)='Detail']]")
                self._js_click(detail_btn)
                time.sleep(1.0)
                detail_text = box.text.strip()
                log.info("Test Post Channel Detail 内容：%s", detail_text)
            except NoSuchElementException:
                pass

            # 点击 OK / Confirm / 确定 按钮关闭弹窗
            ok_xpath = (
                ".//button[contains(@class,'el-button--primary')]"
                " | .//button[normalize-space(.)='OK']"
                " | .//button[normalize-space(.)='Confirm']"
                " | .//button[normalize-space(.)='确定']"
            )
            try:
                ok_btn = box.find_element(By.XPATH, ok_xpath)
                self._js_click(ok_btn)
                log.info("已点击弹窗 OK 按钮")
                time.sleep(0.8)
            except NoSuchElementException:
                pass

            full_result = detail_text if detail_text else result
            if is_fail and not is_success:
                raise RuntimeError(
                    f"Test Post Channel 失败，请检查服务器配置。详情：{full_result!r}"
                )
            return full_result

        except TimeoutException:
            log.warning("Test Post Channel 30 秒内未见弹窗（原生 Alert 和 el-message-box 均未匹配）")
            return ""

    def configure_all(
        self,
        channel_configs: dict[int, PostChannelConfig],
        enabled: bool = True,
        test: bool = True,
    ) -> dict[int, str]:
        """配置所有 Post Channel，返回 {channel_n: test_result}。"""
        results = {}
        for n, cfg in sorted(channel_configs.items()):
            results[n] = self.configure_channel(n, cfg, enabled=enabled, test=test)
        return results


# ─────────────────────────────────────────────────────────────────────────────
# Data Logger 页面（基于截图实际 UI）
# ─────────────────────────────────────────────────────────────────────────────

class DataLoggerPage(_PageBase):
    """
    配置 Data Logger 1 / 2 / 3。

    导航：Data Log 左侧菜单 → 顶部 "Data Loggers ▼" 下拉 → Data Loggers N
    （注意：下拉菜单及面包屑中为 "Data Loggers N"，复数形式）
    """

    # ── 顶部 tab：Data Loggers（Element Plus el-menu-item / el-sub-menu）─────────
    DATA_LOGGERS_TAB = (By.XPATH,
        "//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Data Loggers')]"
        " | //div[contains(@class,'el-sub-menu__title') and contains(normalize-space(.),'Data Loggers')]")

    @staticmethod
    def _logger_item(n: int):
        """Data Loggers N 菜单项（Element Plus el-menu-item）。"""
        return (By.XPATH,
            f"//li[contains(@class,'el-menu-item') and contains(normalize-space(.),'Data Loggers {n}')]")

    # ── Post Channel 下拉（el-select）────────────────────────────────────────
    POST_CHANNEL_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Post Channel')]"
        "/following::div[contains(@class,'el-select')][1]")

    # ── Log File Format（el-select：csv / json）──────────────────────────────
    LOG_FILE_FORMAT_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Format')]"
        "/following::div[contains(@class,'el-select')][1]")

    # ── Log File Name Prefix <input> ─────────────────────────────────────────
    LOG_FILE_NAME_PREFIX_INPUT = (By.XPATH,
        "//input[@placeholder='Enter Log File Name Prefix']"
        " | //label[contains(normalize-space(.),'Log File Name Prefix')]/following::input[1]")

    # ── Log File Length（el-select，如 "1 minute"）────────────────────────────
    LOG_FILE_LENGTH_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Length')]"
        "/following::div[contains(@class,'el-select')][1]")

    # ── Log Interval（el-select，如 "1 minute"）──────────────────────────────
    LOG_INTERVAL_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log Interval')]"
        "/following::div[contains(@class,'el-select')][1]")

    # ── Timestamp Format / Log File Name Format（el-select 或 radio） ─────────
    TIMESTAMP_FORMAT_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Timestamp Format')]"
        "/following::div[contains(@class,'el-select')][1]")
    LOG_FILE_NAME_FORMAT_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and contains(normalize-space(.),'Log File Name Format')]"
        "/following::div[contains(@class,'el-select')][1]")

    # ── Devices Selection 表格 ────────────────────────────────────────────────
    DEVICE_HEADER_CB  = (By.XPATH, "//thead//input[@type='checkbox']")
    DEVICE_TBODY_ROWS = (By.XPATH, "//tbody/tr")

    # ── 保存按钮 ──────────────────────────────────────────────────────────────
    SAVE_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Save']]"
        " | //button[normalize-space(.)='Save']")

    # ── 导航 ─────────────────────────────────────────────────────────────────

    def navigate_to_logger(self, n: int):
        """导航到 Data Loggers N 页面。"""
        self._navigate_data_log()
        self._navigate_url(f"/dataLog/dataLogger/dataLogger{n}")
        log.info("已进入 Data Loggers %d 页面", n)

    # ── 各字段操作 ────────────────────────────────────────────────────────────

    def _set_enable(self, enabled: bool):
        """点击 Enable/Disable 的 el-radio 标签（实际 input 被 Element Plus 隐藏）。"""
        locator = (By.XPATH, "//label[contains(@class,'el-radio') and .//input[@value='true']]") \
                  if enabled \
                  else (By.XPATH, "//label[contains(@class,'el-radio') and .//input[@value='false']]")
        try:
            el = WebDriverWait(self.driver, 8).until(EC.presence_of_element_located(locator))
            self._js_click(el)
        except TimeoutException:
            log.warning("Enable/Disable radio 未找到")

    def _set_timestamp_format(self, fmt: str):
        """
        选择 Timestamp Format（radio 优先，el-select 兜底）。
        HMI-1-7 UI 使用 radio button 组，el-select 定位器在该网关会找到错误的下拉。
        """
        if not self._click_radio_in_group("Timestamp Format", fmt):
            self._select_el_by_text(self.TIMESTAMP_FORMAT_SELECT, fmt, "Timestamp Format")

    def _set_log_file_name_format(self, fmt: str):
        """
        选择 Log File Name Format（radio 优先，el-select 兜底）。
        同 _set_timestamp_format 说明。
        """
        if not self._click_radio_in_group("Log File Name Format", fmt):
            self._select_el_by_text(self.LOG_FILE_NAME_FORMAT_SELECT, fmt, "Log File Name Format")

    def _select_devices(self, target_names: list):
        """
        勾选设备。
        target_names 为空 → 点击表头全选 checkbox，全部勾选。
        target_names 非空 → 只勾选 Device Name 列包含目标名字的行，其余取消。
        """
        # 等待表格渲染
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(self.DEVICE_TBODY_ROWS))
        except TimeoutException:
            log.warning("Devices Selection 表格未找到")
            return

        if not target_names:
            # 全选：点击 header checkbox（若未勾选）
            try:
                header_cb = self.driver.find_element(*self.DEVICE_HEADER_CB)
                if not header_cb.is_selected():
                    self._js_click(header_cb)
                    log.info("Devices Selection：已点击全选 checkbox")
                else:
                    log.info("Devices Selection：全选 checkbox 已是选中状态")
                time.sleep(0.3)
            except NoSuchElementException:
                log.warning("未找到全选 checkbox，改为逐行勾选")
                target_names = []   # 走下面逐行逻辑，空名单 = 全勾
            else:
                return

        # 逐行操作
        rows = self.driver.find_elements(*self.DEVICE_TBODY_ROWS)
        log.info("Devices Selection：共 %d 行，逐行处理", len(rows))
        for idx, row in enumerate(rows):
            # 取第一列文字作为设备名
            cells = row.find_elements(By.XPATH, ".//td")
            device_name = cells[0].text.strip() if cells else row.text.strip().split("\n")[0]

            should_check = (not target_names) or any(
                t.lower() in device_name.lower() for t in target_names)

            cb_els = row.find_elements(By.XPATH, ".//input[@type='checkbox']")
            if not cb_els:
                continue
            cb = cb_els[0]
            is_checked = cb.is_selected()

            if should_check and not is_checked:
                self.driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'})", cb)
                self._js_click(cb)
                log.info("  ✓ %s", device_name)
                time.sleep(0.15)
            elif not should_check and is_checked:
                self._js_click(cb)
                log.info("  ✗ %s（取消）", device_name)
                time.sleep(0.15)
            else:
                log.debug("  - %s（状态正确，跳过）", device_name)

    # ── 主配置流程 ────────────────────────────────────────────────────────────

    def configure_logger(self, n: int, cfg: DataLoggerConfig):
        """配置第 n 个 Data Logger。"""
        log.info(
            "配置 Data Loggers %d：channel=%d  format=%s  length=%s  interval=%s",
            n, cfg.channel_index, cfg.log_file_format,
            cfg.log_file_length, cfg.log_interval,
        )
        self.navigate_to_logger(n)

        # 始终先 Enable，确保所有字段可见可填（即使最终目标是 Disabled 状态）
        self._set_enable(True)

        # 等待完整表单渲染（Post Channel select 出现即表示字段已全部可见）
        try:
            WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(self.POST_CHANNEL_SELECT))
            log.debug("Data Loggers %d 表单已渲染", n)
        except TimeoutException:
            log.warning("Data Loggers %d Post Channel 下拉未在 Enable 后及时出现，继续尝试", n)

        # Post Channel 下拉（el-select：Post Channel N）
        self._select_el_by_text(
            self.POST_CHANNEL_SELECT,
            f"Post Channel {cfg.channel_index}",
            "Post Channel",
        )

        # Timestamp Format
        self._set_timestamp_format(cfg.timestamp_format)

        # Log File Name Format
        self._set_log_file_name_format(cfg.log_file_name_format)

        # Log File Format（el-select：csv / json）
        self._select_el_by_text(
            self.LOG_FILE_FORMAT_SELECT, cfg.log_file_format, "Log File Format")

        # Log File Name Prefix（若配置了则修改，否则保留页面默认）
        if cfg.log_file_name_prefix:
            self._fill(
                self.LOG_FILE_NAME_PREFIX_INPUT,
                cfg.log_file_name_prefix,
                "Log File Name Prefix",
            )

        # Log File Length（el-select：1 minute / 5 minutes / ...）
        self._select_el_by_text(
            self.LOG_FILE_LENGTH_SELECT, cfg.log_file_length, "Log File Length")

        # Log Interval（el-select）
        self._select_el_by_text(
            self.LOG_INTERVAL_SELECT, cfg.log_interval, "Log Interval")

        # Devices Selection
        self._select_devices(cfg.device_names)

        # Save（Enable 状态下保存，确保配置写入）
        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(2)

        # 若目标为 Disabled 状态，再次点击 Disable 并保存
        if not cfg.enabled:
            self._set_enable(False)
            self._safe_click(self.SAVE_BTN, "Save")
            time.sleep(1.5)
            log.info("Data Loggers %d 已配置并禁用", n)
        else:
            log.info("Data Loggers %d 已保存", n)

    def configure_all(self, logger_configs: dict[int, DataLoggerConfig]):
        """配置所有 Data Logger。"""
        for n, cfg in sorted(logger_configs.items()):
            self.configure_logger(n, cfg)

    def disable_logger(self, n: int):
        """将 Data Logger N 置为 Disabled 并保存，用于收到推送数据后关闭。"""
        log.info("禁用 Data Loggers %d", n)
        self.navigate_to_logger(n)
        self._set_enable(False)
        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(1.5)
        log.info("Data Loggers %d 已禁用", n)

    def disable_all(self, logger_configs: dict[int, DataLoggerConfig]):
        """禁用所有已配置的 Data Logger。"""
        for n in sorted(logger_configs.keys()):
            self.disable_logger(n)


# ─────────────────────────────────────────────────────────────────────────────
# Data Log Parameter Config 页面
# ─────────────────────────────────────────────────────────────────────────────

class DataLogParamConfigPage(_PageBase):
    """
    Data Log Parameter Config 页面 —— 为每台设备的每种参数类型全选参数。

    导航：Data Log → Data Loggers ▼ → Data Log Parameter Config
    URL : #/dataLog/dataLoggers/dataLogParameterConfig（如与实际不符，修改 _URL_HASH）

    操作流程（每台设备）：
      1. Device 下拉选设备
      2. 遍历所有（或指定的）Parameter Type：
           a. Parameter Type 下拉选类型
           b. 点击 All 按钮 → 该类型全部参数移入 Selected
      3. 点击 Save
    注：切换 Parameter Type 时 Selected 面板保留已选项（各类型可累积），
        最后一次 Save 提交全部选中的参数。
    """

    _URL_HASH = "/dataLog/dataLoggers/dataLogParameterConfig"

    # Device 下拉（el-select，placeholder "--Select Device--"）
    DEVICE_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and normalize-space(.)='Device']"
        "/following::div[contains(@class,'el-select')][1]")

    # Parameter Type 下拉（el-select，placeholder "--Select Parameter Type--"）
    PARAM_TYPE_SELECT = (By.XPATH,
        "//label[contains(@class,'el-form-item__label') and normalize-space(.)='Parameter Type']"
        "/following::div[contains(@class,'el-select')][1]")

    # "All" 按钮（将左侧全部移入右侧 Selected）
    ALL_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='All']]"
        " | //button[normalize-space(.)='All']")

    # Save 按钮
    SAVE_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Save']]"
        " | //button[normalize-space(.)='Save']")

    # 菜单导航定位器（Data Loggers 子菜单 → Data Log Parameter Config 菜单项）
    _DATA_LOGGERS_SUB_MENU = (By.XPATH,
        "//div[contains(@class,'el-sub-menu__title') and contains(normalize-space(.),'Data Loggers')]"
        " | //li[contains(@class,'el-sub-menu')][.//span[contains(normalize-space(.),'Data Loggers')]]"
        "//div[contains(@class,'el-sub-menu__title')]")
    _PARAM_CONFIG_MENU_ITEM = (By.XPATH,
        "//li[contains(@class,'el-menu-item') and "
        "contains(normalize-space(.),'Data Log Parameter Config')]")

    def navigate(self):
        """
        导航到 Data Log Parameter Config 页面。
        先尝试 URL hash 直达，5s 内未找到 Device 下拉则改用菜单点击导航。
        """
        self._navigate_data_log()
        self._navigate_url(self._URL_HASH)

        # 验证页面是否加载正确（找到 Device 下拉为准）
        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(self.DEVICE_SELECT))
            log.info("已进入 Data Log Parameter Config 页面（URL 导航）")
            return
        except TimeoutException:
            log.warning("URL hash 导航未找到 Device 下拉，改用菜单导航")

        self._navigate_via_menu()

    def _navigate_via_menu(self):
        """通过顶部菜单 Data Loggers → Data Log Parameter Config 导航（URL 导航备用）。"""
        current = self.driver.current_url
        base = current.split('#')[0]
        self.driver.get(f"{base}#/dataLog")
        time.sleep(1.5)

        # 展开 Data Loggers 子菜单
        try:
            el = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(self._DATA_LOGGERS_SUB_MENU))
            self._js_click(el)
            time.sleep(0.8)
            log.debug("已展开 Data Loggers 菜单")
        except TimeoutException:
            log.warning("Data Loggers 子菜单未找到，放弃菜单导航")
            return

        # 点击 Data Log Parameter Config 菜单项
        try:
            el = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(self._PARAM_CONFIG_MENU_ITEM))
            self._js_click(el)
            time.sleep(2.0)
            log.info("已通过菜单导航到 Data Log Parameter Config 页面")
        except TimeoutException:
            log.warning("Data Log Parameter Config 菜单项未找到")

    def _get_dropdown_options(self, locator) -> list:
        """展开 el-select，收集所有未禁用选项的文字后关闭下拉，返回列表。"""
        options = []
        try:
            container = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            inner = container.find_elements(By.XPATH,
                ".//div[contains(@class,'el-select__wrapper')]")
            self._js_click(inner[0] if inner else container)
            time.sleep(0.8)
            items = self.driver.find_elements(By.XPATH,
                "//li[contains(@class,'el-select-dropdown__item')"
                " and not(contains(@class,'disabled'))]")
            options = [el.text.strip() for el in items if el.text.strip()]
        except Exception as e:
            log.warning("获取下拉选项失败：%s", e)
        finally:
            try:
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.2)
            except Exception:
                pass
        return options

    def configure_device(self, device_name: str, param_types: list):
        """
        为单台设备配置：选设备 → 逐类型全选 → 保存。
        param_types 为空时自动读取下拉中所有可用类型。
        """
        log.info("Parameter Config：配置设备 [%s]", device_name)
        self._select_el_by_text(self.DEVICE_SELECT, device_name, "Device")
        time.sleep(1.0)

        types_to_use = list(param_types) if param_types else \
            self._get_dropdown_options(self.PARAM_TYPE_SELECT)

        if not types_to_use:
            log.warning("  设备 [%s] 无可用参数类型，跳过", device_name)
            return

        log.info("  参数类型：%s", types_to_use)
        for pt in types_to_use:
            self._select_el_by_text(self.PARAM_TYPE_SELECT, pt, "Parameter Type")
            time.sleep(0.8)
            self._safe_click(self.ALL_BTN, f"All [{pt}]")
            time.sleep(0.5)
            log.info("  已全选：%s", pt)

        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(1.5)
        log.info("  设备 [%s] 参数配置已保存", device_name)

    def configure_all(self, cfg: DataLogParamConfig):
        """
        遍历 cfg.device_names 中所有设备，依次完成参数配置。
        每台设备都重新 navigate（避免前一台操作残留影响下拉状态）。
        """
        for idx, device_name in enumerate(cfg.device_names):
            if idx > 0:
                # 再次导航以重置页面状态
                self.navigate()
            self.configure_device(device_name, cfg.param_types)

    def configure_all_devices(self, param_types: list = None):
        """
        读取页面 Device 下拉中所有可用设备，逐台全参数配置。
        param_types 为空时自动读取每台设备支持的所有参数类型。
        不在设备间重复导航，直接在当前页面切换 Device 下拉。
        """
        self.navigate()
        device_names = self._get_dropdown_options(self.DEVICE_SELECT)
        if not device_names:
            log.warning("Data Log Parameter Config：Device 下拉无可用设备")
            return
        log.info("发现 %d 台设备：%s", len(device_names), device_names)
        for device_name in device_names:
            self.configure_device(device_name, param_types or [])
