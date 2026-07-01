# -*- coding: utf-8 -*-
"""
mqtt_page.py — MQTT 配置页面自动化（兼容 WEB2、HMI1-7 等多项目）

配置页面 5 个区块（Tab 或表单分组，具体结构由设备 UI 决定）：
  General    → Broker Address / Port / Client ID / Keep Alive / Timeout / Clean Session
  Credential → Username / Password
  SSL-TLS    → Enable / CA File / Cert File / Key File
  LWT        → Topic / QoS
  Topic      → Base Topic / Interval / Retained

跨项目兼容策略：
  - URL Hash 可通过构造参数 url_root 覆盖（默认 "/mqtt"）
  - 导航优先 URL Hash 直达，失败后自动降级为左侧菜单点击
  - 所有定位器基于 Label 文字 + 相对 XPath，不依赖 id/name 属性
  - Tab 切换用文字匹配；若页面无独立 Tab，字段均在当前页直接定位
  - 不引用 Protocols/ 目录内任何其他模块，自包含

使用示例（从仓库根目录执行）：
  from Protocols.MQTT.mqtt_page import MQTTPage, MQTTConfig
  from Protocols.MQTT.mqtt_page import MQTTGeneralConfig, MQTTTopicConfig

  page = MQTTPage(driver)
  page.navigate()
  page.configure(MQTTConfig(
      general=MQTTGeneralConfig(broker_address="mqtt.accu.com", broker_port=1883, client_id="gw-001"),
      topic=MQTTTopicConfig(base_topic="acurev/data", interval="30 seconds"),
  ))
"""

from __future__ import annotations

import logging
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from comm.ui_base_page import BasePage

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 配置数据结构
# ─────────────────────────────────────────────────────────────────────────────

_IP_PATTERN = re.compile(r"^\d{1,3}(\.\d{1,3}){3}$")


@dataclass
class MQTTGeneralConfig:
    """
    General 区块：基本连接参数。

    broker_address : Broker 域名，设备 UI 只接受域名格式（如 www.accu.com），
                     不接受 IP 地址。若传入 IP 格式，实例化时会输出 WARNING。
    broker_port    : 0~65535，默认 1883；SSL 模式下通常改为 8883。
    client_id      : ≤40 字符。
    keep_alive     : 10~600 秒。
    timeout        : 3~120 秒。
    """
    broker_address: str = ""
    broker_port: int = 1883
    client_id: str = ""
    keep_alive: int = 60
    timeout: int = 10
    clean_session: bool = True

    def __post_init__(self):
        if self.broker_address and _IP_PATTERN.match(self.broker_address):
            log.warning(
                "MQTTGeneralConfig.broker_address 传入了 IP 地址 '%s'，"
                "设备 UI 的 Broker Address 字段只接受域名（如 www.accu.com），"
                "填写后可能被设备拒绝或截断。",
                self.broker_address,
            )


@dataclass
class MQTTCredentialConfig:
    """Credential 区块：用户名 / 密码认证。"""
    username: str = ""
    password: str = ""


@dataclass
class MQTTSSLConfig:
    """
    SSL-TLS 区块：证书文件路径（必须是绝对路径，send_keys 直接上传）。

    enabled=False 时仅禁用 SSL，不上传任何证书。
    各文件路径为空时跳过对应上传步骤（保留页面现有证书）。
    """
    enabled: bool = True
    ca_file: str = ""        # CA 根证书（.crt）
    cert_file: str = ""      # 客户端证书（.crt）
    key_file: str = ""       # 客户端私钥（.key）


@dataclass
class MQTTLWTConfig:
    """LWT（Last Will and Testament）区块。"""
    topic: str = ""
    qos: int = 0             # 0 / 1 / 2


@dataclass
class MQTTTopicConfig:
    """Topic 区块：数据发布配置。"""
    base_topic: str = ""
    interval: str = "30 seconds"   # 与 UI 下拉文字完全一致
    retained: bool = False
    qos: int = 0                   # 发布 QoS（0/1/2），在 Topic and Parameter Selection tab
    devices: list = field(default_factory=list)
    # devices 示例：["Basic Module"]
    # 非空时：勾选对应行的 checkbox，并点击其参数选择按钮选择全部参数（All → Confirm）


@dataclass
class MQTTConfig:
    """
    MQTT 完整配置聚合。

    各区块为 None 时跳过对应配置步骤（保留设备现有配置）。
    enabled=False 时：先填写全部字段并保存，再禁用（保留配置供后续测试开启）。
    """
    enabled: bool = True
    general: Optional[MQTTGeneralConfig] = None
    credential: Optional[MQTTCredentialConfig] = None
    ssl: Optional[MQTTSSLConfig] = None
    lwt: Optional[MQTTLWTConfig] = None
    topic: Optional[MQTTTopicConfig] = None


@dataclass
class MQTTDualModeConfig:
    """
    双模式测试配置：一次性描述不加密（plain）和加密（SSL/TLS）两种模式所需的全部参数。

    两种模式共用 broker_address / client_id / 凭据 / topic 等字段；
    差异仅为端口和 SSL 证书，由本类统一管理。

    使用示例：
        dual = MQTTDualModeConfig(
            broker_address="mqtt.accu.com",
            client_id="gw-001",
            plain_port=1883,
            ssl_port=8883,
            ssl_ca_file="C:/certs/ca.crt",
            ssl_cert_file="C:/certs/client.crt",
            ssl_key_file="C:/certs/client.key",
        )
        results = page.test_both_modes(dual)
        assert results["plain"]["pass"]
        assert results["ssl"]["pass"]
    """

    # ── 两种模式共用 ──────────────────────────────────────────────────────────
    broker_address: str = ""          # 域名，如 mqtt.accu.com（设备只接受域名）
    client_id: str = ""
    keep_alive: int = 60
    timeout: int = 10
    clean_session: bool = True

    # 可选：凭据（两种模式共用同一组用户名/密码）
    username: str = ""
    password: str = ""

    # 可选：Topic 发布
    base_topic: str = ""
    interval: str = "30 seconds"
    retained: bool = False

    # 可选：LWT
    lwt_topic: str = ""
    lwt_qos: int = 0

    # ── 不加密专有 ────────────────────────────────────────────────────────────
    plain_port: int = 1883

    # ── 加密专有 ──────────────────────────────────────────────────────────────
    ssl_port: int = 8883
    ssl_ca_file: str = ""             # CA 根证书绝对路径
    ssl_cert_file: str = ""           # 客户端证书绝对路径
    ssl_key_file: str = ""            # 客户端私钥绝对路径

    # ── 测试结束后保留的状态 ──────────────────────────────────────────────────
    final_mode: str = "ssl"           # "ssl" 或 "plain"

    def __post_init__(self):
        if self.broker_address and _IP_PATTERN.match(self.broker_address):
            log.warning(
                "MQTTDualModeConfig.broker_address 传入了 IP '%s'，"
                "设备 Broker Address 字段只接受域名。",
                self.broker_address,
            )

    def _make_general(self, port: int) -> MQTTGeneralConfig:
        return MQTTGeneralConfig(
            broker_address=self.broker_address,
            broker_port=port,
            client_id=self.client_id,
            keep_alive=self.keep_alive,
            timeout=self.timeout,
            clean_session=self.clean_session,
        )

    def _make_credential(self) -> Optional[MQTTCredentialConfig]:
        if self.username or self.password:
            return MQTTCredentialConfig(username=self.username, password=self.password)
        return None

    def _make_topic(self) -> Optional[MQTTTopicConfig]:
        if self.base_topic:
            return MQTTTopicConfig(
                base_topic=self.base_topic,
                interval=self.interval,
                retained=self.retained,
            )
        return None

    def _make_lwt(self) -> Optional[MQTTLWTConfig]:
        if self.lwt_topic:
            return MQTTLWTConfig(topic=self.lwt_topic, qos=self.lwt_qos)
        return None

    def to_plain_config(self) -> MQTTConfig:
        """生成不加密模式的完整 MQTTConfig（SSL 禁用，端口 plain_port）。"""
        return MQTTConfig(
            enabled=True,
            general=self._make_general(self.plain_port),
            credential=self._make_credential(),
            ssl=MQTTSSLConfig(enabled=False),
            lwt=self._make_lwt(),
            topic=self._make_topic(),
        )

    def to_ssl_config(self) -> MQTTConfig:
        """生成加密模式的完整 MQTTConfig（SSL 启用，端口 ssl_port）。"""
        return MQTTConfig(
            enabled=True,
            general=self._make_general(self.ssl_port),
            credential=self._make_credential(),
            ssl=MQTTSSLConfig(
                enabled=True,
                ca_file=self.ssl_ca_file,
                cert_file=self.ssl_cert_file,
                key_file=self.ssl_key_file,
            ),
            lwt=self._make_lwt(),
            topic=self._make_topic(),
        )


# ─────────────────────────────────────────────────────────────────────────────
# MQTT 页面自动化类
# ─────────────────────────────────────────────────────────────────────────────

class MQTTPage(BasePage):
    """
    MQTT 配置页面自动化，兼容使用 Element Plus + Vue Router 的多款网关产品。

    参数
    ----
    driver   : Selenium WebDriver 实例
    url_root : MQTT 页面根路由 hash（不含 '#'），默认 "/mqtt"。
               若 WEB2 或 HMI1-7 的实际路由不同，构造时传入覆盖：
                 MQTTPage(driver, url_root="/protocols/mqtt")
    """

    # ── 通用按钮 ─────────────────────────────────────────────────────────────
    SAVE_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Save']]"
        " | //button[normalize-space(.)='Save']")
    TEST_BTN = (By.XPATH,
        "//button[.//span[normalize-space(.)='Test MQTT']]"
        " | //button[normalize-space(.)='Test MQTT']"
        " | //button[.//span[normalize-space(.)='Test Connection']]"
        " | //button[normalize-space(.)='Test Connection']")

    # ── Enable / Disable（兼容 el-radio value='true'/'false' 和 span 文字两种写法） ──
    ENABLE_RADIO = (By.XPATH,
        "//label[contains(@class,'el-radio') and .//input[@value='true']]"
        " | //label[.//span[normalize-space(.)='Enable']][1]")
    DISABLE_RADIO = (By.XPATH,
        "//label[contains(@class,'el-radio') and .//input[@value='false']]"
        " | //label[.//span[normalize-space(.)='Disable']][1]")

    # ── MQTT 页面专有特征（Broker 相关 label，避免误判其他含 Save 的配置页） ──
    MQTT_PAGE_MARKER = (By.XPATH,
        "//div[contains(@class,'el-form-item')]"
        "[.//label[contains(normalize-space(.),'Broker')]]"
        " | //label[contains(normalize-space(.),'Broker Address')]"
        " | //label[contains(normalize-space(.),'Broker Port')]")

    # ── MQTT 子菜单/标签/内容区入口 ──────────────────────────────────────────
    # 覆盖四类场景：
    #   1. 侧边栏 el-menu-item（WEB2 等 Element Plus 侧边菜单）
    #   2. el-tabs__item（页面内 Tab 切换）
    #   3. 水平导航栏 el-sub-menu__title（HMI1-7 Protocols 总览页顶部 Tab）
    #   4. 主内容区 <a> router-link / 协议卡片 / nav-item
    MQTT_MENU_ITEM = (By.XPATH,
        # 侧边栏 el-menu-item
        "//li[contains(@class,'el-menu-item') and normalize-space(.)='MQTT']"
        " | //li[contains(@class,'el-menu-item') and .//span[normalize-space(.)='MQTT']]"
        # Tab
        " | //div[contains(@class,'el-tabs__item') and normalize-space(.)='MQTT']"
        " | //div[contains(@class,'el-tabs__item') and .//span[normalize-space(.)='MQTT']]"
        # nav-item class
        " | //*[contains(@class,'nav-item')]"
        "   [normalize-space(.)='MQTT' or .//span[normalize-space(.)='MQTT']]"
        # 水平导航栏 el-sub-menu（HMI1-7）——点击 title div 触发导航/展开
        " | //div[contains(@class,'el-sub-menu__title')]"
        "   [normalize-space(.)='MQTT' or .//span[normalize-space(.)='MQTT']]"
        # 主内容区 <a> 链接
        " | //a[not(ancestor::*[contains(@class,'el-menu')])]"
        "   [normalize-space(.)='MQTT' or .//span[normalize-space(.)='MQTT']]"
        # 主内容区 el-card 或协议卡片 div
        " | //div[not(ancestor::*[contains(@class,'el-menu')])]"
        "   [contains(@class,'el-card') or contains(@class,'protocol') or contains(@class,'card')]"
        "   [normalize-space(.)='MQTT' or .//span[normalize-space(.)='MQTT']"
        "    or .//h3[normalize-space(.)='MQTT'] or .//p[normalize-space(.)='MQTT']]")

    def __init__(self, driver, url_root: str = "/mqtt"):
        super().__init__(driver)
        self._url_root = url_root.lstrip("/")

    # ─────────────────────────────────────────────────────────────────────────
    # Element Plus 底层操作（自包含，不依赖外部基类扩展）
    # ─────────────────────────────────────────────────────────────────────────

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
        """定位输入框 → JS 设值 + 触发 input/change 事件，兼容 v-model 和 v-model.lazy。

        流程：
          1. JS click 聚焦
          2. 用原生 property setter 设置 value（绕过 Vue 劫持直接赋值）
          3. dispatchEvent('input')  → 触发 v-model 实时更新
          4. dispatchEvent('change') → 触发 v-model.lazy 更新
        """
        try:
            el = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            self._js_click(el)
            # 通过原生 value setter + 双事件确保 Vue 响应式更新
            self.driver.execute_script(
                "var setter = Object.getOwnPropertyDescriptor("
                "    window.HTMLInputElement.prototype, 'value').set;"
                "setter.call(arguments[0], arguments[1]);"
                "arguments[0].dispatchEvent(new Event('input',    {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change',   {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('blur',     {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('focusout', {bubbles:true}));",
                el, str(value)
            )
            # 若 JS setter 未生效（某些 el-input 场景），降级为 send_keys
            actual = el.get_attribute("value") or ""
            if actual != str(value):
                log.debug("_fill JS setter 未生效（actual=%r expected=%r），改用 send_keys", actual, value)
                ActionChains(self.driver).click(el).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
                el.send_keys(Keys.DELETE)
                if str(value):
                    el.send_keys(str(value))
                el.send_keys(Keys.TAB)
                time.sleep(0.15)
            log.debug("填写 %s = %s", name, value)
        except TimeoutException:
            log.warning("输入框未找到：%s  %s", name, locator)

    def _fill_by_label(self, label_text: str, value: str):
        """
        通过 label 文字定位同一 el-form-item 内的文本输入框并填写。
        排除 type=file / checkbox / radio，只匹配 text/number/password 类输入框。
        """
        locator = (By.XPATH,
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_text}')]]"
            f"//input[not(@type='file') and not(@type='checkbox') and not(@type='radio')]"
            f" | "
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_text}')]]"
            f"//textarea"
        )
        self._fill(locator, value, label_text)

    def _get_el_select_value(self, trigger_locator) -> str:
        """读取 el-select 当前已选中的显示文字（不打开下拉）。"""
        try:
            container = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located(trigger_locator))
            for xpath in [
                ".//span[contains(@class,'el-select__selected-item')]",
                ".//input[contains(@class,'el-select__input')]",
                ".//div[contains(@class,'el-select__wrapper')]//span",
            ]:
                els = container.find_elements(By.XPATH, xpath)
                if els:
                    val = (els[0].text or els[0].get_attribute("value") or "").strip()
                    if val:
                        return val
        except Exception:
            pass
        return ""

    def _select_el_by_text(self, trigger_locator, text: str, name: str = ""):
        """
        操作 Element Plus el-select：点击 wrapper 展开下拉 → 点击目标 option。
        当前值已与目标一致时跳过，避免不必要的下拉操作。
        完成或失败后按 Escape 关闭残留下拉，避免遮挡后续元素。
        """
        current = self._get_el_select_value(trigger_locator)
        if current.lower() == text.lower():
            log.debug("el-select %s 当前已是 '%s'，跳过", name, text)
            return
        try:
            container = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(trigger_locator))
            inner = container.find_elements(By.XPATH,
                ".//div[contains(@class,'el-select__wrapper')]")
            trigger = inner[0] if inner else container
            # 使用 ActionChains 真实鼠标点击（生成完整的 mousedown/click 事件链），
            # 避免 js_click 在全套运行时因历史事件状态导致下拉无法展开
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'})", trigger)
            time.sleep(0.1)
            # 找到 input 元素用于 aria-expanded 检测
            _inp_els = container.find_elements(By.XPATH, ".//input")
            _aria_el = _inp_els[0] if _inp_els else trigger

            def _is_expanded():
                return _aria_el.get_attribute("aria-expanded") == "true"

            ActionChains(self.driver).move_to_element(trigger).click().perform()
            time.sleep(0.6)

            if not _is_expanded():
                # 首次点击未展开，重试（点 input 元素）
                log.warning("el-select %s 首次点击未展开，重试 input 点击", name)
                if _inp_els:
                    ActionChains(self.driver).move_to_element(_inp_els[0]).click().perform()
                else:
                    ActionChains(self.driver).move_to_element(trigger).click().perform()
                time.sleep(0.8)

            if not _is_expanded():
                log.warning("el-select %s 两次点击均未展开", name)
            else:
                # 下拉已展开，在全局 li 中搜索（popper 挂在 body 下）
                # 用 JS 检测可见性（比 Selenium is_displayed() 更可靠）
                def _js_visible(el):
                    try:
                        return self.driver.execute_script(
                            "var s=window.getComputedStyle(arguments[0]);"
                            "return s.display!=='none' && s.visibility!=='hidden' "
                            "&& parseFloat(s.opacity)>0 "
                            "&& arguments[0].offsetWidth>0 && arguments[0].offsetHeight>0;",
                            el)
                    except Exception:
                        return False

                for pat in [
                    f"//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{text}']",
                    f"//li[contains(@class,'el-select-dropdown__item') and .//span[normalize-space(.)='{text}']]",
                    f"//li[contains(@class,'el-select-dropdown__item') and contains(normalize-space(.),'{text}')]",
                ]:
                    opts = self.driver.find_elements(By.XPATH, pat)
                    clickable = [o for o in opts if _js_visible(o)]
                    if clickable:
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({block:'nearest'})", clickable[0])
                        time.sleep(0.1)
                        self._js_click(clickable[0])
                        log.debug("el-select %s 选择：%s", name, text)
                        time.sleep(0.4)
                        return
            # 回退：全局 XPath
            for pat in [
                f"//li[contains(@class,'el-select-dropdown__item') and normalize-space(.)='{text}']",
                f"//div[contains(@class,'el-popper')]//li[contains(normalize-space(.),'{text}')]",
            ]:
                try:
                    opt = WebDriverWait(self.driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, pat)))
                    self._js_click(opt)
                    log.debug("el-select %s 选择(全局)：%s", name, text)
                    time.sleep(0.4)
                    return
                except (TimeoutException, AttributeError):
                    continue
            log.warning("el-select %s 未找到选项 '%s'", name, text)
        except (TimeoutException, AttributeError) as e:
            log.warning("el-select %s 容器未找到（'%s'）：%s", name, text, e)
        finally:
            try:
                self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.2)
            except Exception:
                pass

    def _select_el_by_label(self, label_text: str, option_text: str):
        """通过 label 文字找到同一 form-item 内的 el-select 并选择指定选项。"""
        locator = (By.XPATH,
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_text}')]]"
            f"//div[contains(@class,'el-select')]")
        self._select_el_by_text(locator, option_text, label_text)

    def _set_toggle(self, label_text: str, enabled: bool):
        """
        设置 el-switch 或 radio 组（Enable/Disable、True/False、Yes/No 等）。
        优先尝试 el-switch（aria-checked），再尝试 el-radio label。
        """
        form_item_xpath = (
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_text}')]]"
        )

        # ── 尝试 el-switch ──────────────────────────────────────────────────
        switches = self.driver.find_elements(By.XPATH,
            f"{form_item_xpath}//button[contains(@class,'el-switch')]"
            f" | {form_item_xpath}//div[contains(@class,'el-switch')]")
        if switches:
            sw = switches[0]
            aria = (sw.get_attribute("aria-checked") or "").lower()
            is_on = aria == "true"
            if enabled != is_on:
                self._js_click(sw)
                log.debug("Switch %s → %s", label_text, enabled)
                time.sleep(0.3)
            return

        # ── 尝试 el-radio（value 属性匹配 true/false/Enable/Disable 等） ────
        value_map = {
            True:  ["true", "1", "Enable", "Yes"],
            False: ["false", "0", "Disable", "No"],
        }
        for val in value_map[enabled]:
            for xpath in [
                f"{form_item_xpath}//label[contains(@class,'el-radio') and .//input[@value='{val}']]",
                f"{form_item_xpath}//label[.//span[normalize-space(.)='{val}']]",
            ]:
                els = self.driver.find_elements(By.XPATH, xpath)
                if els:
                    # 已是目标状态（is-checked）则跳过，避免重复点击
                    if "is-checked" in (els[0].get_attribute("class") or ""):
                        log.debug("Radio %s 已是 %s，跳过", label_text, val)
                        return
                    self._js_click(els[0])
                    log.debug("Radio %s → %s", label_text, val)
                    time.sleep(0.3)
                    return

        log.warning("未能定位 toggle/radio：%s", label_text)

    def _read_input_by_label(self, label_text: str) -> str:
        """读取 label 对应输入框的当前 value（不修改）。字段不存在时返回空串。"""
        locator = (By.XPATH,
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_text}')]]"
            f"//input[not(@type='file') and not(@type='checkbox') and not(@type='radio')]"
        )
        els = self.driver.find_elements(*locator)
        if els:
            return (els[0].get_attribute("value") or "").strip()
        return ""

    def read_config(self) -> dict:
        """
        回读当前页面上各区块的配置值，返回 dict。
        用于校验已配置状态或做配置前后对比。

        返回格式示例：
        {
          "enabled": True,
          "general": {
              "broker_address": "mqtt.accu.com",
              "broker_port": "1883",
              "client_id": "gw-001",
              "keep_alive": "60",
              "timeout": "10",
          },
          "credential": {"username": "admin"},
          "lwt": {"topic": "lwt/topic", "qos": "0"},
          "topic": {"base_topic": "acurev/data", "interval": "30 seconds"},
          "ssl_enabled": False,
        }
        """
        self._switch_tab("General")
        result: dict = {"enabled": self._is_enabled()}
        result["general"] = {
            "broker_address": self._read_input_by_label("Broker Address"),
            "broker_port": self._read_input_by_label("Broker Port"),
            "client_id":   self._read_input_by_label("Client ID"),
            "keep_alive":  self._read_input_by_label("Keep Alive"),
            "timeout":     self._read_input_by_label("Timeout"),
        }

        self._switch_tab("User Credential")
        result["credential"] = {
            "username": self._read_input_by_label("Username"),
        }

        self._switch_tab("Last Will and Testament")
        lwt_on = self.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Last Will')]]"
            "//label[contains(@class,'is-checked')]"
            "[.//input[@value='true'] or .//span[normalize-space(.)='Enable']]")
        lwt_enabled = len(lwt_on) > 0
        result["lwt"] = {
            "enabled": lwt_enabled,
            "topic": self._read_input_by_label("Topic") if lwt_enabled else "",
            "qos":   self._get_el_select_value(
                (By.XPATH,
                 "//div[contains(@class,'el-form-item')]"
                 "[.//label[contains(normalize-space(.),'Qos') or contains(normalize-space(.),'QoS')]]"
                 "//div[contains(@class,'el-select')]")) if lwt_enabled else "",
        }

        self._switch_tab("Topic and Parameter Selection")
        # ── 读取 Retained 状态（Yes=True / No=False） ────────────────────────
        _retained_fi = (
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Retained')]]"
        )
        _retained_yes = self.driver.find_elements(By.XPATH,
            f"{_retained_fi}//label[contains(@class,'is-checked')]"
            "[.//input[@value='true'] or .//span[normalize-space(.)='Yes']]")
        _retained_switch = self.driver.find_elements(By.XPATH,
            f"{_retained_fi}//button[contains(@class,'el-switch')][@aria-checked='true']"
            f" | {_retained_fi}//div[contains(@class,'el-switch')][@aria-checked='true']")
        _retained = len(_retained_yes) > 0 or len(_retained_switch) > 0

        result["topic"] = {
            "base_topic": self._read_input_by_label("Base Topic") or self._read_input_by_label("Topic"),
            "qos":        self._get_el_select_value(
                (By.XPATH,
                 "//div[contains(@class,'el-form-item')]"
                 "[.//label[contains(normalize-space(.),'Qos')]]"
                 "//div[contains(@class,'el-select')]")),
            "interval":   self._get_el_select_value(
                (By.XPATH,
                 "//div[contains(@class,'el-form-item')]"
                 "[.//label[contains(normalize-space(.),'Interval')]]"
                 "//div[contains(@class,'el-select')]")),
            "retained":       _retained,
            "payload_format": self._get_el_select_value(
                (By.XPATH,
                 "//div[contains(@class,'el-form-item')]"
                 "[.//label[contains(normalize-space(.),'Payload')]]"
                 "//div[contains(@class,'el-select')]")),
        }

        self._switch_tab("SSL/TLS")
        # 检查 SSL enable 状态（el-switch 或 radio）
        ssl_fi = (
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'SSL')]]"
        )
        ssl_switch = self.driver.find_elements(By.XPATH,
            f"{ssl_fi}//button[contains(@class,'el-switch')]")
        if ssl_switch:
            aria = (ssl_switch[0].get_attribute("aria-checked") or "").lower()
            result["ssl_enabled"] = aria == "true"
        else:
            # 兼容 value='true' 和 span 文字 'Enable' 两种写法
            ssl_on = self.driver.find_elements(By.XPATH,
                f"{ssl_fi}//label[contains(@class,'is-checked')]"
                "[.//input[@value='true'] or .//span[normalize-space(.)='Enable']]")
            result["ssl_enabled"] = len(ssl_on) > 0

        return result

    def _dismiss_overlay_dialog(self):
        """关闭登录后可能残留的 el-overlay-message-box 确认弹窗。"""
        try:
            dialogs = self.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-overlay-message-box') "
                "or contains(@class,'el-message-box')]")
            for dlg in dialogs:
                try:
                    if not dlg.is_displayed():
                        continue
                except Exception:
                    continue
                for xpath in [
                    ".//button[.//span[normalize-space(.)='Yes']]",
                    ".//button[.//span[normalize-space(.)='Continue']]",
                    ".//button[contains(@class,'el-button--primary')]",
                    ".//button[contains(@class,'el-button')]",
                ]:
                    btns = dlg.find_elements(By.XPATH, xpath)
                    if btns:
                        self._js_click(btns[0])
                        log.info("已关闭确认弹窗")
                        time.sleep(0.5)
                        break
        except Exception as e:
            log.debug("dismiss_overlay_dialog: %s", e)

    # ─────────────────────────────────────────────────────────────────────────
    # 导航
    # ─────────────────────────────────────────────────────────────────────────

    def navigate(self):
        """
        导航到 MQTT 配置页面。
        优先 URL Hash 直达；若关键元素 5 秒内未出现，降级为左侧菜单点击。
        """
        self._dismiss_overlay_dialog()
        current = self.driver.current_url
        base = current.split('#')[0]
        self.driver.get(f"{base}#{self._url_root}")
        time.sleep(1.5)

        if not self._page_loaded(timeout=5):
            log.warning("URL 导航未到达 MQTT 页面，降级为菜单导航")
            self._navigate_via_menu()
        else:
            log.info("已进入 MQTT 配置页面（URL 导航：#/%s）", self._url_root)

    def _page_loaded(self, timeout: int = 5) -> bool:
        """
        检测是否已在 MQTT 配置页面。
        使用 Broker 相关 label 作为专有特征，避免误判其他含 Save 按钮的配置页。
        """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(self.MQTT_PAGE_MARKER))
            return True
        except TimeoutException:
            return False

    # ── DFS 导航常量 ────────────────────────────────────────────────────────
    _NAV_SKIP = frozenset({"logout", "service", "about", "help"})
    # 列表中位置越靠前优先级越高（用索引作为评分，越小越先试）
    _NAV_PRIORITY = (
        "protocol", "mqtt", "communication", "setting",
        "network", "advanced", "commission",
    )

    def _navigate_via_menu(self):
        """
        深度优先自动扫描所有菜单/导航项，定位 MQTT 配置页面。

        · 同时扫描 nav-item（顶部）和 el-menu-item（横向/侧边）两类入口
        · 含关键字（Protocol / MQTT / Communication 等）的项优先尝试
        · 点击每项后立即向下递归（DFS），确保在当前视图中继续寻找子菜单，
          而不是先把同层其他项全部尝试完（BFS 会导致视图已切走）
        · 最多递归 3 层；每次重新查找元素，避免 stale reference
        """
        tried: set = set()

        def _priority_score(text: str) -> int:
            tl = text.lower()
            for i, kw in enumerate(self._NAV_PRIORITY):
                if kw in tl:
                    return i
            return len(self._NAV_PRIORITY)

        def _dfs(depth: int) -> bool:
            if depth >= 4:
                return False

            candidates = self._collect_nav_candidates(tried)
            candidates.sort(key=_priority_score)

            log.info("第 %d 层：%d 个候选项，前几项=%s",
                     depth + 1, len(candidates), candidates[:5])

            for txt in candidates:
                tried.add(txt)
                el = self._find_nav_el(txt)
                if el is None:
                    continue
                log.info("  %s点击：%s", "  " * depth, txt)
                try:
                    self._js_click(el)
                except Exception:
                    continue
                time.sleep(1.5)

                # 已直接进入 MQTT 配置页面（Broker label 出现）
                if self._page_loaded(timeout=2):
                    log.info("  → 已进入 MQTT 配置页面（%s）", txt)
                    return True

                # 出现了 MQTT 子菜单入口，点击进入
                mqtt_el = self._find_mqtt_menu_item()
                if mqtt_el is not None:
                    log.info("  → 发现 MQTT 子菜单，点击进入")
                    try:
                        self._js_click(mqtt_el)
                        time.sleep(1.5)
                    except Exception:
                        pass
                    # el-sub-menu__title 的 JS click 不会触发 hover 弹出菜单；
                    # HMI1-7 将 MQTT 子项（General/Credential/…）直接展开在侧边栏，
                    # 只需点击侧边栏的 "General" 即可进入 MQTT 配置页。
                    if not self._page_loaded(timeout=2):
                        sidebar_gen = self.driver.find_elements(By.XPATH,
                            "//li[contains(@class,'el-menu-item')"
                            " and normalize-space(.)='General']")
                        if sidebar_gen:
                            log.info("  → 侧边栏 MQTT General，点击进入")
                            self._js_click(sidebar_gen[0])
                            time.sleep(1.5)
                    return True

                # 向下递归，在当前视图中寻找更深的菜单
                if _dfs(depth + 1):
                    return True

            return False

        if not _dfs(0):
            log.warning("DFS 扫描（4 层）未能定位 MQTT 配置页面")

    def _collect_nav_candidates(self, tried: set) -> list:
        """收集当前页面所有可点击菜单/导航/标签项的文字（排除已尝试和工具类项）。"""
        texts, seen = [], set()
        for xpath in [
            "//*[contains(@class,'nav-item')]",
            "//li[contains(@class,'el-menu-item')]",
            "//div[contains(@class,'el-tabs__item')]",
            "//div[contains(@class,'el-submenu__title')]",
        ]:
            for el in self.driver.find_elements(By.XPATH, xpath):
                try:
                    raw = el.text or el.get_attribute("innerText") or ""
                    txt = raw.strip().split("\n")[0].strip()
                    if (txt and txt not in tried and txt not in seen
                            and txt.lower() not in self._NAV_SKIP):
                        texts.append(txt)
                        seen.add(txt)
                except Exception:
                    pass
        return texts

    def _find_nav_el(self, text: str):
        """按文字重新查找导航/菜单/标签元素（每次重查，避免 stale reference）。"""
        for xpath in [
            f"//*[contains(@class,'nav-item') and normalize-space(.)='{text}']",
            f"//li[contains(@class,'el-menu-item') and normalize-space(.)='{text}']",
            f"//div[contains(@class,'el-tabs__item') and normalize-space(.)='{text}']",
            f"//div[contains(@class,'el-submenu__title') and normalize-space(.)='{text}']",
        ]:
            els = self.driver.find_elements(By.XPATH, xpath)
            if els:
                return els[0]
        return None

    def _find_mqtt_menu_item(self):
        """检查当前页面是否出现了 MQTT 菜单入口（出现即表示 MQTT 在下一层）。"""
        els = self.driver.find_elements(*self.MQTT_MENU_ITEM)
        return els[0] if els else None

    # ─────────────────────────────────────────────────────────────────────────
    # Tab / Section 切换
    # ─────────────────────────────────────────────────────────────────────────

    def _switch_tab(self, tab_name: str) -> bool:
        """
        尝试切换到指定名称的 Tab / Section。
        支持 el-menu-item、el-tabs__item、button tab 等多种 UI 结构。
        若页面无独立 Tab（单 Form 布局），返回 False，调用方继续在当前页填写。
        """
        for pat in [
            f"//li[contains(@class,'el-menu-item') and normalize-space(.)='{tab_name}']",
            f"//li[contains(@class,'el-menu-item') and .//span[normalize-space(.)='{tab_name}']]",
            f"//div[contains(@class,'el-tabs__item') and normalize-space(.)='{tab_name}']",
            f"//div[contains(@class,'el-tabs__item') and .//span[normalize-space(.)='{tab_name}']]",
            f"//button[contains(@class,'tab') and normalize-space(.)='{tab_name}']",
        ]:
            els = self.driver.find_elements(By.XPATH, pat)
            if els:
                self._js_click(els[0])
                time.sleep(0.5)
                log.debug("切换 Tab：%s", tab_name)
                return True
        log.debug("未找到 Tab '%s'，字段在当前页直接填写", tab_name)
        return False

    # ─────────────────────────────────────────────────────────────────────────
    # Enable / Disable
    # ─────────────────────────────────────────────────────────────────────────

    def _is_enabled(self) -> bool:
        """
        读取 MQTT 总开关状态。
        调用前应先 _switch_tab("General")，确保 MQTT Enable 字段可见。
        先尝试以 label 含 'MQTT' 的 form-item 精确定位，
        降级时在当前可见范围内找第一个 Enable/Disable radio 组。
        """
        try:
            # 精确：MQTT Enable form-item（label 含 "MQTT"，排除 SSL/LWT 等其他 Enable 组）
            mqtt_fi = (
                "//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(normalize-space(.)),'MQTT')]]"
            )
            # Enable is-checked（value='true' 或 span='Enable'）
            on_els = self.driver.find_elements(By.XPATH,
                f"{mqtt_fi}//label[contains(@class,'el-radio') and contains(@class,'is-checked')]"
                "[.//input[@value='true'] or .//span[normalize-space(.)='Enable']]")
            if on_els:
                return True
            off_els = self.driver.find_elements(By.XPATH,
                f"{mqtt_fi}//label[contains(@class,'el-radio') and contains(@class,'is-checked')]"
                "[.//input[@value='false'] or .//span[normalize-space(.)='Disable']]")
            if off_els:
                return False

            # 降级：页面首个 Enable/Disable radio 组（value='true'/'false'）
            on2 = self.driver.find_elements(By.XPATH,
                "//label[contains(@class,'el-radio') and contains(@class,'is-checked')"
                " and .//input[@value='true']]")
            if on2:
                return True
            off2 = self.driver.find_elements(By.XPATH,
                "//label[contains(@class,'el-radio') and contains(@class,'is-checked')"
                " and .//input[@value='false']]")
            if off2:
                return False
        except Exception:
            pass
        return False

    def _set_enable(self, enabled: bool):
        """设置 MQTT 总开关；已处于目标状态时跳过，避免触发不必要的 re-render。
        先切换到 General tab 确保 MQTT Enable radio 可见，防止在其他 Tab 上调用时
        _is_enabled() 返回错误结果或 WebDriverWait 白白等待 8s 超时。
        """
        self._switch_tab("General")
        if self._is_enabled() == enabled:
            log.debug("MQTT Enable 已是 %s，跳过", enabled)
            return
        locator = self.ENABLE_RADIO if enabled else self.DISABLE_RADIO
        try:
            el = WebDriverWait(self.driver, 8).until(
                EC.presence_of_element_located(locator))
            self._js_click(el)
            log.debug("MQTT Enable → %s", enabled)
            time.sleep(0.5)
        except TimeoutException:
            log.warning("Enable/Disable radio 未找到，locator=%s", locator)

    # ─────────────────────────────────────────────────────────────────────────
    # 各区块配置
    # ─────────────────────────────────────────────────────────────────────────

    def _configure_general(self, cfg: MQTTGeneralConfig):
        self._switch_tab("General")
        if cfg.broker_address:
            self._fill_by_label("Broker Address", cfg.broker_address)
        if cfg.broker_port:
            self._fill_by_label("Broker Port", str(cfg.broker_port))
        if cfg.client_id:
            self._fill_by_label("Client ID", cfg.client_id)
        self._fill_by_label("Keep Alive", str(cfg.keep_alive))
        self._fill_by_label("Timeout", str(cfg.timeout))
        self._set_toggle("Clean Session", cfg.clean_session)
        log.info("General：address=%s  port=%d  client_id=%s",
                 cfg.broker_address, cfg.broker_port, cfg.client_id)

    def _configure_credential(self, cfg: MQTTCredentialConfig):
        self._switch_tab("User Credential")
        # 无论是否为空都填写，确保 _restore_base() 能清空凭据
        self._fill_by_label("Username", cfg.username)
        self._fill_by_label("Password", cfg.password)
        log.info("Credential：username=%s", cfg.username)

    def _configure_ssl(self, cfg: MQTTSSLConfig):
        self._switch_tab("SSL/TLS")
        time.sleep(0.3)

        # SSL enable/disable toggle（在 SSL/TLS 区块内）
        self._set_toggle("SSL", cfg.enabled)
        if not cfg.enabled:
            log.info("SSL/TLS 已禁用")
            return

        time.sleep(0.5)  # 等待证书字段展开

        # 证书上传：按 label 关键字定位 <input type="file">，send_keys 绝对路径
        for kw, path in [("CA", cfg.ca_file), ("Cert", cfg.cert_file), ("Key", cfg.key_file)]:
            if path:
                self._upload_cert_file(kw, path)

        log.info("SSL-TLS：已配置 %s 证书",
                 "/".join(k for k, p in [("CA", cfg.ca_file), ("Cert", cfg.cert_file),
                                          ("Key", cfg.key_file)] if p))

    def _upload_cert_file(self, label_keyword: str, file_path: str):
        """
        通过 label 关键字定位 <input type="file"> 并 send_keys 文件路径。
        若 label 精确定位失败，按 CA→Cert→Key 顺序使用全局第 1/2/3 个 file input。
        """
        label_patterns = [
            f"//div[contains(@class,'el-form-item')]"
            f"[.//label[contains(normalize-space(.),'{label_keyword}')]]"
            f"//input[@type='file']",
            f"//label[contains(normalize-space(.),'{label_keyword}')]"
            f"/following::input[@type='file'][1]",
        ]
        for pat in label_patterns:
            els = self.driver.find_elements(By.XPATH, pat)
            if els:
                try:
                    self.driver.execute_script(
                        "arguments[0].removeAttribute('style')", els[0])
                    els[0].send_keys(file_path)
                    log.info("  %s File 已上传：%s", label_keyword, file_path)
                    time.sleep(0.8)
                    return
                except Exception as e:
                    log.warning("  %s File 上传失败（%s），尝试下一 locator", label_keyword, e)

        # 降级：全局 file input 按 CA/Cert/Key 顺序（index 1/2/3）
        idx = {"CA": 1, "Cert": 2, "Key": 3}.get(label_keyword, 1)
        fallback = self.driver.find_elements(By.XPATH, "//input[@type='file']")
        if idx <= len(fallback):
            try:
                self.driver.execute_script(
                    "arguments[0].removeAttribute('style')", fallback[idx - 1])
                fallback[idx - 1].send_keys(file_path)
                log.info("  %s File 降级上传（第 %d 个 file input）：%s",
                         label_keyword, idx, file_path)
                time.sleep(0.8)
            except Exception as e:
                log.warning("  %s File 降级上传失败：%s", label_keyword, e)
        else:
            log.warning("  未找到 %s File 的 file input，跳过", label_keyword)

    def _configure_lwt(self, cfg: MQTTLWTConfig):
        self._switch_tab("Last Will and Testament")
        lwt_enabled = bool(cfg.topic)
        self._set_toggle("Last Will Enable", lwt_enabled)
        if lwt_enabled:
            time.sleep(0.5)  # 等待 topic 字段展开
            for label in ["Last Will Topic", "LWT Topic", "Topic"]:
                locator = (By.XPATH,
                    f"//div[contains(@class,'el-form-item')]"
                    f"[.//label[contains(normalize-space(.),'{label}')]]"
                    f"//input[not(@type='file') and not(@type='checkbox') and not(@type='radio')]")
                els = self.driver.find_elements(*locator)
                if els:
                    self._fill(locator, cfg.topic, label)
                    break
            # LWT QoS（仅在 LWT 启用后可见；label 文字为 "Qos"，选项格式 "Qos 0"/"QoS 0"/"0"）
            _lwt_qos_locator = (By.XPATH,
                "//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Qos') or contains(normalize-space(.),'QoS')]]"
                "//div[contains(@class,'el-select')]")
            _lwt_qos_str = str(cfg.qos)
            for _qt in [f"Qos {cfg.qos}", f"QoS {cfg.qos}", _lwt_qos_str]:
                self._select_el_by_text(_lwt_qos_locator, _qt, f"LWT-Qos({_qt})")
                if _lwt_qos_str in self._get_el_select_value(_lwt_qos_locator):
                    break
        log.info("LWT：enabled=%s  topic=%s  qos=%d", lwt_enabled, cfg.topic, cfg.qos)

    def _configure_topic(self, cfg: MQTTTopicConfig):
        self._switch_tab("Topic and Parameter Selection")
        if cfg.base_topic:
            for label in ["Base Topic", "Topic"]:
                locator = (By.XPATH,
                    f"//div[contains(@class,'el-form-item')]"
                    f"[.//label[contains(normalize-space(.),'{label}')]]"
                    f"//input[not(@type='file') and not(@type='checkbox') and not(@type='radio')]")
                els = self.driver.find_elements(*locator)
                if els:
                    self._fill(locator, cfg.base_topic, label)
                    break
        # 发布 QoS（label 为 "Qos"；选项格式因设备而异：'Qos 0'/'QoS 0'/'0'，逐一尝试）
        _qos_locator = (By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Qos') or contains(normalize-space(.),'QoS')]]"
            "//div[contains(@class,'el-select')]")
        _qos_current = self._get_el_select_value(_qos_locator)
        _qos_str = str(cfg.qos)
        if _qos_str not in _qos_current:
            for _qt in [f"Qos {cfg.qos}", f"QoS {cfg.qos}", _qos_str]:
                self._select_el_by_text(_qos_locator, _qt, f"Qos({_qt})")
                if _qos_str in self._get_el_select_value(_qos_locator):
                    break
        self._select_el_by_label("Interval", cfg.interval)
        self._set_toggle("Retained", cfg.retained)
        log.info("Topic：base_topic=%s  qos=%d  interval=%s  retained=%s",
                 cfg.base_topic, cfg.qos, cfg.interval, cfg.retained)

        # 设备选择与参数配置
        if cfg.devices:
            self._configure_devices_selection(cfg.devices)

    def _configure_devices_selection(self, devices: list):
        """
        在 Topic and Parameter Selection Tab 中：
          1. 勾选每个设备名对应行的 checkbox
          2. 点击该行的参数选择按钮（图标按钮）
          3. 在弹窗中遍历 Parameter Type 下拉的全部选项，每个都点击 All（全选参数）
          4. 最后点击 Confirm 关闭弹窗

        参数
        ----
        devices : 设备名列表，如 ["Basic Module"]
                  设备名需与页面 Device Name 列完全一致
        """
        for device_name in devices:
            # ── 找行 ──────────────────────────────────────────────────────────
            row_xp = (f"//tr[td[contains(normalize-space(.),'{device_name}')]]"
                      f" | //div[contains(@class,'el-table__row')]"
                      f"[contains(normalize-space(.),'{device_name}')]")
            rows = self.driver.find_elements(By.XPATH, row_xp)
            if not rows:
                log.warning("Devices Selection：未找到设备行 '%s'", device_name)
                continue
            row = rows[0]

            # ── 勾选 checkbox ─────────────────────────────────────────────────
            cbs = row.find_elements(By.XPATH,
                ".//label[contains(@class,'el-checkbox')]")
            if cbs:
                cb = cbs[0]
                if "is-checked" not in (cb.get_attribute("class") or ""):
                    self.driver.execute_script("arguments[0].click()", cb)
                    time.sleep(0.5)
                    log.info("Devices Selection：已勾选 '%s'", device_name)
                else:
                    log.info("Devices Selection：'%s' 已处于勾选状态", device_name)
            else:
                log.warning("Devices Selection：未在行内找到 checkbox，设备='%s'", device_name)

            # ── 点击参数选择按钮 ──────────────────────────────────────────────
            btns = row.find_elements(By.XPATH, ".//button")
            if not btns:
                log.warning("Devices Selection：未在行内找到参数按钮，设备='%s'", device_name)
                continue
            self.driver.execute_script("arguments[0].click()", btns[0])
            time.sleep(1.5)

            # ── 等弹窗出现 ────────────────────────────────────────────────────
            dlg_xp = ("//div[contains(@class,'el-dialog')"
                      " and contains(@class,'c_parameter_config')]")
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, dlg_xp)))
            except TimeoutException:
                log.warning("Devices Selection：参数配置弹窗未出现，设备='%s'", device_name)
                continue
            dlg = self.driver.find_element(By.XPATH, dlg_xp)

            # ── 遍历 Parameter Type 所有选项，每种类型全选参数 ────────────────
            # 找弹窗内的 el-select（Parameter Type 下拉）
            pt_select = dlg.find_element(By.XPATH, ".//div[contains(@class,'el-select')]")
            # 用 aria-controls 获取对应下拉 list 的 id，避免与页面其他 el-select 混淆
            pt_input = pt_select.find_element(By.XPATH,
                ".//input[@aria-controls]")
            dropdown_list_id = pt_input.get_attribute("aria-controls")

            # 展开下拉，读取所有选项文字
            self.driver.execute_script("arguments[0].click()", pt_select)
            time.sleep(0.6)
            if dropdown_list_id:
                opt_xp = f"//ul[@id='{dropdown_list_id}']/li"
            else:
                opt_xp = ("//div[contains(@class,'el-select-dropdown')]"
                          "[last()]//li[contains(@class,'el-select-dropdown__item')]")
            opt_texts = [
                (self.driver.execute_script(
                    "return arguments[0].textContent", el) or "").strip()
                for el in self.driver.find_elements(By.XPATH, opt_xp)
            ]
            # 收起下拉
            self.driver.execute_script("arguments[0].click()", pt_select)
            time.sleep(0.4)

            log.info("Devices Selection：Parameter Type 选项 %s，设备='%s'",
                     opt_texts, device_name)

            for opt_text in opt_texts:
                if not opt_text:
                    continue
                # 展开并选中该选项
                self.driver.execute_script("arguments[0].click()", pt_select)
                time.sleep(0.5)
                target_opt = None
                for el in self.driver.find_elements(By.XPATH, opt_xp):
                    t = (self.driver.execute_script(
                        "return arguments[0].textContent", el) or "").strip()
                    if t == opt_text:
                        target_opt = el
                        break
                if target_opt:
                    self.driver.execute_script("arguments[0].click()", target_opt)
                    time.sleep(0.6)
                else:
                    log.warning("Devices Selection：未找到 Parameter Type 选项 '%s'", opt_text)
                    continue

                # 点击 All（全选当前 Parameter Type 下的参数）
                all_btns = dlg.find_elements(By.XPATH,
                    ".//button[normalize-space(.)='All']")
                if all_btns:
                    self.driver.execute_script("arguments[0].click()", all_btns[0])
                    time.sleep(0.5)
                    log.info("Devices Selection：%s → All 已点击，设备='%s'",
                             opt_text, device_name)

            # ── 点击 Confirm ──────────────────────────────────────────────────
            confirm_btns = dlg.find_elements(By.XPATH,
                ".//button[normalize-space(.)='Confirm']")
            if not confirm_btns:
                confirm_btns = dlg.find_elements(By.XPATH,
                    ".//div[contains(@class,'el-dialog__footer')]//button")
            if confirm_btns:
                self.driver.execute_script("arguments[0].click()", confirm_btns[0])
                time.sleep(1.0)
                log.info("Devices Selection：已 Confirm，设备='%s'", device_name)
            else:
                log.warning("Devices Selection：未找到 Confirm 按钮，设备='%s'", device_name)

    # ─────────────────────────────────────────────────────────────────────────
    # 保存与测试连接
    # ─────────────────────────────────────────────────────────────────────────

    def _read_form_error(self) -> str:
        """
        读取当前页面上所有可见的表单/Toast 错误提示文字。

        同时检查：
          - el-form-item__error  ：字段级校验错误（红色提示，紧跟字段下方）
          - el-message--error    ：全局错误 Toast（右上角弹出）

        返回
        ----
        有错误时返回拼接后的文字（多条用换行分隔）；无错误时返回空串。
        """
        msgs: list[str] = []

        # ── 字段级错误 ────────────────────────────────────────────────────────
        # 同时查有无文本的和暂无文本的（Vue 3 有时 .text 为空但 textContent 有值）
        field_errs = self.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item__error')]")
        for el in field_errs:
            try:
                # 优先用 JS textContent，兼容 Vue 3 虚拟 DOM 渲染时机
                txt = (self.driver.execute_script(
                    "return arguments[0].textContent", el) or "").strip()
                if not txt:
                    txt = el.text.strip()
                if txt:
                    msgs.append(txt)
            except Exception:
                pass

        # ── 全局 Toast 错误 ───────────────────────────────────────────────────
        toast_errs = self.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-message--error')]"
            "[not(contains(@style,'display: none'))]"
            " | //div[contains(@class,'el-notification--error')]"
            "[not(contains(@style,'display: none'))]")
        for el in toast_errs:
            try:
                txt = el.text.strip()
                if txt:
                    msgs.append(txt)
            except Exception:
                pass

        return "\n".join(msgs)

    def save(self) -> bool:
        """
        点击 Save 按钮，等待 UI 反馈，返回保存是否成功。

        检测逻辑（轮询最多 3 秒）：
          1. 检测到 el-message--success / el-notification--success → True
          2. 检测到 el-message--error / el-notification--error
             或 el-form-item__error（字段校验错误）             → False
          3. 超时无明确 feedback → True（与旧行为一致，不误判为失败）

        返回
        ----
        True  — 保存成功（或无法确定）
        False — 检测到明确的错误提示
        """
        self._safe_click(self.SAVE_BTN, "Save")
        time.sleep(0.8)  # 等待 Vue 3 表单校验 / Toast 渲染

        deadline = time.time() + 4.0
        while time.time() < deadline:
            # 成功 toast
            success_els = self.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-message--success')]"
                "[not(contains(@style,'display: none'))]"
                " | //div[contains(@class,'el-notification--success')]"
                "[not(contains(@style,'display: none'))]")
            if success_els:
                log.info("MQTT 配置保存成功")
                return True

            # 错误 toast 或字段级校验错误
            err_text = self._read_form_error()
            if err_text:
                log.warning("MQTT 配置保存失败：%s", err_text)
                return False

            time.sleep(0.3)

        # 超时未见明显 feedback，默认视为成功
        log.info("MQTT 配置已保存（未检测到明确反馈）")
        return True

    def test_connection(self) -> str:
        """
        点击 Test MQTT 按钮，等待弹窗显示最终结果后返回内容文字。

        注意：
        - Test MQTT 按钮仅在 General tab 可见，本方法先切换到 General。
        - 弹窗初始显示 "Result: Connecting"，等待最多 30s 直到状态更新。
        - 只读弹窗内容区（el-message-box__content），排除按钮文字干扰。
        - 未找到按钮或超时均返回空串。
        """
        # Test MQTT 按钮只在 General tab 上，先切到 General
        self._switch_tab("General")
        time.sleep(0.5)

        btns = self.driver.find_elements(*self.TEST_BTN)
        if not btns:
            log.debug("页面无 Test MQTT 按钮，跳过")
            return ""

        self._js_click(btns[0])
        log.info("已点击 Test MQTT，等待结果弹窗…")

        # 等弹窗出现（el-overlay-message-box）
        dialog_locator = (By.XPATH,
            "//div[@role='dialog' and contains(@class,'el-overlay')]"
            " | //div[contains(@class,'el-message-box__wrapper') "
            "and not(contains(@style,'display: none'))]"
        )
        try:
            box = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(dialog_locator))
        except TimeoutException:
            log.warning("Test MQTT：10s 内未见弹窗")
            return ""

        # 等待 "Connecting" 状态消失（最多 35s）
        content_xp = (
            ".//div[contains(@class,'el-message-box__content')]"
            " | .//div[contains(@class,'el-message-box__message')]"
            " | .//p[contains(@class,'el-message-box__message')]"
        )
        deadline = time.time() + 35
        result = ""
        while time.time() < deadline:
            # 优先读内容区（不含按钮文字）
            content_els = box.find_elements(By.XPATH, content_xp)
            if content_els:
                result = (self.driver.execute_script(
                    "return arguments[0].textContent", content_els[0]) or "").strip()
            else:
                result = box.text.strip()

            if result and "connecting" not in result.lower():
                break
            time.sleep(1)

        log.info("Test MQTT 最终结果：%s", result)

        # 关闭弹窗
        for xpath in [
            ".//button[normalize-space(.)='OK']",
            ".//button[normalize-space(.)='Close']",
            ".//button[contains(@class,'el-button--primary')]",
            ".//button",
        ]:
            close_btns = box.find_elements(By.XPATH, xpath)
            if close_btns:
                self._js_click(close_btns[0])
                time.sleep(0.8)
                break

        return result

    # ─────────────────────────────────────────────────────────────────────────
    # 结果解析
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _is_test_success(result: str) -> bool:
        """
        解析 Test MQTT 返回文字，判断连接是否成功。

        WEB2 弹窗内容区格式：
          成功 → "Result: Success"
          失败 → "Result: Failed" / "Result: Timeout" / "Result: Error"

        规则（按优先级）：
          1. 空串              → False（未收到任何反馈）
          2. 含成功关键词      → True
          3. 含失败关键词      → False
          4. 有文字但无关键词  → True（弹窗出现即视为操作完成，具体内容留日志）

        注：去掉 "ok" 关键词，避免按钮文字 "OK" 干扰判断。
        """
        if not result:
            return False
        lower = result.lower()
        if any(kw in lower for kw in ("success", "connected", "pass", "test pass")):
            return True
        if any(kw in lower for kw in ("fail", "error", "timeout", "refused",
                                       "unreachable", "test fail", "failed", "connecting")):
            return False
        return True

    # ─────────────────────────────────────────────────────────────────────────
    # 主配置流程
    # ─────────────────────────────────────────────────────────────────────────

    def configure(self, cfg: MQTTConfig, test: bool = False) -> str:
        """
        完整配置流程：Enable → General → Credential → SSL-TLS → LWT → Topic → Save → [Test]

        参数
        ----
        cfg  : MQTTConfig 聚合配置，各区块为 None 时跳过对应步骤
        test : True 时在 Save 后执行 Test Connection 并返回结果文字

        返回
        ----
        Test Connection 结果文字（test=False 时返回空串）
        """
        # 先 Enable，确保所有字段可见可填
        self._set_enable(True)
        time.sleep(0.5)

        if cfg.general is not None:
            self._configure_general(cfg.general)
            self.save()
        if cfg.credential is not None:
            self._configure_credential(cfg.credential)
            self.save()
        if cfg.ssl is not None:
            self._configure_ssl(cfg.ssl)
            self.save()
        if cfg.lwt is not None:
            self._configure_lwt(cfg.lwt)
            self.save()
        if cfg.topic is not None:
            self._configure_topic(cfg.topic)
            self.save()

        if not cfg.enabled:
            self._set_enable(False)
            self.save()

        return self.test_connection() if test else ""

    def enable(self):
        """仅启用 MQTT 并保存，不修改其他配置。"""
        self._set_enable(True)
        self.save()

    def disable(self):
        """仅禁用 MQTT 并保存，不修改其他配置。"""
        self._set_enable(False)
        self.save()

    def test_both_modes(
        self,
        dual_cfg: "MQTTDualModeConfig",
        stop_on_failure: bool = True,
        assert_pass: bool = False,
    ) -> dict:
        """
        依次测试不加密（plain）和加密（SSL/TLS）两种连接模式，返回各自测试结果。

        流程：
          1. 应用不加密配置（端口 plain_port，SSL 禁用）→ Save → Test Connection
          2. 若第 1 步失败且 stop_on_failure=True，跳过第 2 步直接返回
          3. 应用加密配置（端口 ssl_port，SSL 启用，上传证书）→ Save → Test Connection
          4. 按 final_mode 保留最终状态（默认 "ssl"）

        参数
        ----
        dual_cfg         : MQTTDualModeConfig，统一描述两种模式的差异
        stop_on_failure  : True（默认）时不加密模式失败后跳过加密模式测试；
                           False 时无论第 1 步结果如何均继续执行第 2 步
        assert_pass      : True 时任一模式失败（或被跳过）立即抛出 AssertionError

        返回
        ----
        {
          "plain": {"pass": bool,       "result": str,  "skipped": False},
          "ssl":   {"pass": bool|None,  "result": str,  "skipped": bool},
        }
        ssl.skipped=True 表示因 plain 失败而未执行，此时 pass=None、result=""。

        示例
        ----
        dual = MQTTDualModeConfig(
            broker_address="mqtt.accu.com",
            client_id="gw-001",
            ssl_ca_file="C:/certs/ca.crt",
            ssl_cert_file="C:/certs/client.crt",
            ssl_key_file="C:/certs/client.key",
        )
        results = page.test_both_modes(dual)
        # plain 失败 → ssl 自动跳过，results["ssl"]["skipped"] == True
        """
        outcomes: dict = {}

        # ── 1. 不加密模式 ─────────────────────────────────────────────────────
        log.info("=" * 60)
        log.info("测试不加密模式（plain  port=%d）", dual_cfg.plain_port)
        plain_result = self.configure(dual_cfg.to_plain_config(), test=True)
        plain_pass = self._is_test_success(plain_result)
        outcomes["plain"] = {"pass": plain_pass, "result": plain_result, "skipped": False}
        log.info("Plain 结果：%s  pass=%s", plain_result or "(无弹窗)", plain_pass)

        # ── 短路判断 ──────────────────────────────────────────────────────────
        if not plain_pass and stop_on_failure:
            log.warning("不加密模式失败，跳过加密模式测试（stop_on_failure=True）")
            outcomes["ssl"] = {"pass": None, "result": "", "skipped": True}
            if assert_pass:
                raise AssertionError(
                    f"不加密模式 Test Connection 失败，结果：{plain_result!r}"
                )
            return outcomes

        if assert_pass and not plain_pass:
            raise AssertionError(
                f"不加密模式 Test Connection 失败，结果：{plain_result!r}"
            )

        # ── 2. 加密模式 ───────────────────────────────────────────────────────
        log.info("=" * 60)
        log.info("测试加密模式（SSL/TLS  port=%d）", dual_cfg.ssl_port)
        ssl_result = self.configure(dual_cfg.to_ssl_config(), test=True)
        ssl_pass = self._is_test_success(ssl_result)
        outcomes["ssl"] = {"pass": ssl_pass, "result": ssl_result, "skipped": False}
        log.info("SSL  结果：%s  pass=%s", ssl_result or "(无弹窗)", ssl_pass)

        if assert_pass and not ssl_pass:
            raise AssertionError(
                f"加密模式 Test Connection 失败，结果：{ssl_result!r}"
            )

        # ── 3. 按 final_mode 设置最终状态 ────────────────────────────────────
        if dual_cfg.final_mode == "plain":
            log.info("恢复为不加密状态（final_mode=plain）")
            self.configure(dual_cfg.to_plain_config())
        # final_mode="ssl"：最后一次 configure 已是 ssl_cfg，无需额外操作

        log.info("=" * 60)
        log.info(
            "双模式测试完毕：plain=%s  ssl=%s",
            "PASS" if plain_pass else "FAIL",
            "PASS" if ssl_pass else "FAIL",
        )
        return outcomes
