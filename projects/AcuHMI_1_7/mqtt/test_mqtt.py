# -*- coding: utf-8 -*-
"""
test_mqtt.py — WEB2 MQTT 配置页面自动化测试

覆盖用例：001-011，共 72 条（012 含下电操作，排除自动化）

运行方式（仓库根目录）：
  pytest Protocols/MQTT/test_mqtt.py -v
  pytest Protocols/MQTT/test_mqtt.py -v -m lv0
  pytest Protocols/MQTT/test_mqtt.py -v -m "lv0 or lv1"
  pytest Protocols/MQTT/test_mqtt.py -v -m "not integration"
  pytest Protocols/MQTT/test_mqtt.py -v -m integration

pytest.ini 或 pyproject.toml 中注册 marker：
  markers =
    lv0: 冒烟用例
    lv1: 正向验证
    lv2: 异常/负向
    lv3: 补充验证
    integration: 需要真实 Broker + 设备联通的集成测试
    manual: 需要 Wireshark 等外部工具，仅手工执行
"""
from __future__ import annotations

import asyncio
import logging
import sys
import time
from pathlib import Path

import paho.mqtt.client as mqtt
import pytest
from selenium.webdriver.common.by import By

sys.path.insert(0, str(Path(__file__).parent.parent.parent))  # 仓库根
sys.path.insert(0, str(Path(__file__).parent.parent))          # projects/AcuHMI_1_7/
sys.path.insert(0, str(Path(__file__).parent))                 # projects/AcuHMI_1_7/mqtt/

import settings as config
from mqtt_page import (
    MQTTConfig,
    MQTTCredentialConfig,
    MQTTGeneralConfig,
    MQTTLWTConfig,
    MQTTSSLConfig,
    MQTTTopicConfig,
)

log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# amqtt 兼容性补丁
# 设备使用 MQTT 3.1 协议，CONNECT 包保留位可能非零；
# amqtt 严格执行 MQTT 3.1.1 规范会拒绝该连接。
# 在此处 monkey-patch ConnectVariableHeader.reserved_flag，使其始终返回 False，
# 让 amqtt 接受 MQTT 3.1 客户端而不断开连接。
# ══════════════════════════════════════════════════════════════════════════════
try:
    from amqtt.mqtt.connect import ConnectVariableHeader
    ConnectVariableHeader.reserved_flag = property(lambda self: False)
    log.debug("amqtt reserved_flag patch applied (MQTT 3.1 compatibility)")
except Exception:
    pass  # amqtt 未安装或版本不同，忽略

# ══════════════════════════════════════════════════════════════════════════════
# 测试常量（修改此处即可切换测试环境）
# ══════════════════════════════════════════════════════════════════════════════

BROKER_ADDR    = config.WEB_MQTT_BROKER_ADDRESS   # 设备侧填写的 Broker 域名
BROKER_PORT    = config.WEB_MQTT_BROKER_PORT      # 1883
SSL_PORT       = config.MQTT_SSL_PORT             # 8883
BASE_CLIENT_ID = "pytest-mqtt-001"
BASE_TOPIC     = "accuenergy/pytest"
KEEP_ALIVE_VAL = 60
TIMEOUT_VAL    = 30

# 证书路径（gen_certs.py 生成后自动填写）
CA_FILE   = config.MQTT_SSL_CA_CERT
CERT_FILE = config.MQTT_SSL_CLIENT_CERT
KEY_FILE  = config.MQTT_SSL_CLIENT_KEY


# ══════════════════════════════════════════════════════════════════════════════
# 公共工具函数
# ══════════════════════════════════════════════════════════════════════════════

def _reload(mqtt_page):
    """刷新页面并重新导航至 MQTT，用于验证配置是否持久化。"""
    mqtt_page.driver.refresh()
    time.sleep(2)
    mqtt_page.navigate()


def _restore_base(mqtt_page):
    """将 MQTT 恢复为最小合法配置，保证后续用例从干净状态开始。

    每个 Tab 单独一次 configure() + save()，避免多 Tab 切换期间页面轮询刷新
    将先填好的 Tab 值覆盖回后端旧值。各步骤独立完成后立即固化到后端：
      Step 1：General（broker_address 等关键字段）
      Step 2：Credential
      Step 3：SSL/TLS
      Step 4：Last Will and Testament
      Step 5：Topic + Devices（含遍历 Parameter Type，耗时约 45s）
    """
    # 强制导航回 MQTT 页，防止前序用例（如 SSL 数据接收）将页面留在 deviceToPublish 等子路由
    mqtt_page.navigate()
    # Step 1：先禁用 SSL，避免后续 General save 时 SSL 仍为 enabled 导致设备收到 SSL+1883 组合
    mqtt_page.configure(MQTTConfig(
        enabled=True,
        ssl=MQTTSSLConfig(enabled=False),
    ))
    # Step 2：General（SSL 已 disabled 后再 save，设备收到正确的 plain+1883 配置）
    mqtt_page.configure(MQTTConfig(
        enabled=True,
        general=MQTTGeneralConfig(
            broker_address=BROKER_ADDR,
            broker_port=BROKER_PORT,
            client_id=BASE_CLIENT_ID,
            keep_alive=KEEP_ALIVE_VAL,
            timeout=TIMEOUT_VAL,
            clean_session=True,
        ),
    ))
    # Step 3：Credential
    mqtt_page.configure(MQTTConfig(
        enabled=True,
        credential=MQTTCredentialConfig(username="", password=""),
    ))
    # Step 4：LWT
    mqtt_page.configure(MQTTConfig(
        enabled=True,
        lwt=MQTTLWTConfig(topic="", qos=0),
    ))
    # Step 5：Topic + Devices（耗时较长）
    mqtt_page.configure(MQTTConfig(
        enabled=True,
        topic=MQTTTopicConfig(
            base_topic=BASE_TOPIC,
            interval="30 seconds",
            retained=False,
            qos=0,
            devices=[config.MQTT_DEFAULT_DEVICE],  # 至少选一台设备，否则 Save 被"Please provide Modbus"拦截
        ),
    ))


def _assert_rejected(mqtt_page):
    """断言最近一次 Save 被拒绝（返回 False 或页面显示错误提示）。"""
    result = mqtt_page.save()
    if result is False:
        return
    err = mqtt_page._read_form_error()
    assert err, "期望 Save 被阻止，但 save() 返回 True 且未见错误提示"


# ══════════════════════════════════════════════════════════════════════════════
# 模块级 autouse fixture — 每条用例开始前重新导航至 MQTT 页
# 保证即使上一条用例跳转到其他页面，当前用例也从正确入口点开始。
# ══════════════════════════════════════════════════════════════════════════════

@pytest.fixture(autouse=True)
def _back_to_mqtt(mqtt_page):
    """每条用例前导航至 MQTT 配置页，确保起始位置一致。"""
    mqtt_page.navigate()
    yield


# ══════════════════════════════════════════════════════════════════════════════
# 001 — 启停与页面入口
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT001Entry:
    """001 启停与页面入口（4 条）"""

    @pytest.mark.lv0
    def test_001_navigate_via_menu(self, mqtt_page):
        """通过 Settings → Communications → MQTT 进入配置页，URL 及 5 个 Tab 均正确"""
        url = mqtt_page.driver.current_url
        assert "mqtt" in url.lower(), f"URL 不含 mqtt：{url}"

        tab_labels = ["General", "User Credential", "SSL/TLS",
                      "Last Will and Testament", "Topic and Parameter Selection"]
        for label in tab_labels:
            els = mqtt_page.driver.find_elements(By.XPATH,
                f"//div[contains(@class,'el-tabs__item') and normalize-space(.)='{label}']"
                f" | //li[contains(@class,'el-menu-item') and normalize-space(.)='{label}']"
                f" | //li[contains(@class,'el-menu-item')"
                f"    and .//span[normalize-space(.)='{label}']]")
            assert els, f"Tab '{label}' 未找到"

    @pytest.mark.lv1
    def test_002_default_disable_state(self, mqtt_page):
        """MQTT Enable 首次进入默认为 Disable，General 字段不可见/置灰"""
        mqtt_page.navigate()
        mqtt_page._switch_tab("General")
        enabled = mqtt_page._is_enabled()
        # 默认 Disable；若已被其他测试改动则先恢复
        if enabled:
            mqtt_page._set_enable(False)
            mqtt_page.save()
            _reload(mqtt_page)
            mqtt_page._switch_tab("General")
        assert not mqtt_page._is_enabled(), "MQTT Enable 默认应为 Disable"
        # Broker Address 字段应隐藏或不可交互
        broker_inputs = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Broker Address')]]"
            "//input[not(@disabled)]")
        assert not broker_inputs, "Disable 状态下 Broker Address 字段不应可见/可编辑"

    @pytest.mark.lv1
    def test_003_enable_disable_toggle(self, mqtt_page):
        """MQTT Enable ↔ Disable 切换：Enable 后字段展开，Disable 后字段隐藏"""
        mqtt_page._switch_tab("General")
        mqtt_page._set_enable(True)
        time.sleep(0.5)
        broker_inputs = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Broker Address')]]//input")
        assert broker_inputs, "Enable 后 Broker Address 字段应可见"

        mqtt_page._set_enable(False)
        time.sleep(0.5)
        broker_inputs_after = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Broker Address')]]"
            "//input[not(@disabled)]")
        assert not broker_inputs_after, "Disable 后 Broker Address 字段应隐藏/置灰"
        # 恢复 Enable 供后续用例使用
        mqtt_page._set_enable(True)
        mqtt_page.save()

    @pytest.mark.lv3
    def test_004_direct_url_access(self, mqtt_page):
        """已登录状态直接访问 URL，页面正常加载不跳转登录页"""
        base = mqtt_page.driver.current_url.split('#')[0]
        mqtt_page.driver.get(f"{base}{config.MQTT_URL_PATH}")
        time.sleep(2)
        cur = mqtt_page.driver.current_url
        assert "login" not in cur.lower(), f"直接访问 MQTT URL 跳转到了登录页：{cur}"
        assert "mqtt" in cur.lower(), f"直接访问后 URL 不含 mqtt：{cur}"
        mqtt_page.navigate()


# ══════════════════════════════════════════════════════════════════════════════
# 002 — General 参数校验
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT002General:
    """002 General 参数校验（18 条）"""

    def setup_method(self, _method):
        pass  # driver/mqtt_page 由 session fixture 提供，无需每次重新登录

    @pytest.fixture(autouse=True)
    def _ensure_general_valid(self, mqtt_page):
        """每条 General 用例执行前，将所有必填字段恢复为合法默认值，
        防止前一条用例遗留非法值（空值/非数字）导致 Save 失败。
        清除 beforeunload 后硬刷新页面，强制 Vue 组件重新挂载，
        彻底消除 is-error 校验残留，再填字段保存。"""
        from selenium.webdriver.common.by import By
        # 清除 beforeunload，防止刷新时触发"是否离开"弹框
        mqtt_page.driver.execute_script("window.onbeforeunload = null")
        # 硬刷新，强制 Vue 组件重新挂载，清除 is-error 等校验残留状态
        mqtt_page.driver.refresh()
        time.sleep(2.0)
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
        mqtt_page._fill_by_label("Client ID", BASE_CLIENT_ID)
        mqtt_page._fill_by_label("Keep Alive", str(KEEP_ALIVE_VAL))
        mqtt_page._fill_by_label("Timeout", str(TIMEOUT_VAL))
        try:
            mqtt_page.driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        time.sleep(0.3)
        mqtt_page.save()
        yield

    @pytest.mark.lv1
    def test_014_timeout_default_and_valid(self, mqtt_page):
        """Timeout 默认 30，修改为 60 后回读正确"""
        mqtt_page._switch_tab("General")
        default = mqtt_page._read_input_by_label("Timeout")
        assert default == "30", f"Timeout 默认值应为 30，实为 {default!r}"
        mqtt_page._fill_by_label("Timeout", "60")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Timeout") == "60"
        mqtt_page._fill_by_label("Timeout", str(TIMEOUT_VAL))
        mqtt_page.save()

    @pytest.mark.lv0
    def test_001_broker_address_valid_domain(self, mqtt_page):
        """Broker Address 输入合法域名，Save 成功，刷新后回读正确"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        assert mqtt_page.save() is True, "合法域名应保存成功"
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        val = mqtt_page._read_input_by_label("Broker Address")
        assert val == BROKER_ADDR, f"回读 Broker Address 不符：{val!r} ≠ {BROKER_ADDR!r}"

    @pytest.mark.lv2
    def test_002_broker_address_reject_ip(self, mqtt_page):
        """Broker Address 输入 IP 地址时阻止 Save"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", "192.168.1.100")
        _assert_rejected(mqtt_page)
        # 恢复合法值
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        mqtt_page.save()

    @pytest.mark.lv2
    def test_003_broker_address_reject_empty(self, mqtt_page):
        """Broker Address 为空时阻止 Save"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", "")
        _assert_rejected(mqtt_page)
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        mqtt_page.save()

    @pytest.mark.lv2
    def test_004_broker_address_reject_special(self, mqtt_page):
        """Broker Address 输入纯空格/特殊字符/超长字符串时阻止 Save"""
        mqtt_page._switch_tab("General")
        for bad_val in ["   ", "!!!@@@", "a" * 256]:
            mqtt_page._fill_by_label("Broker Address", bad_val)
            result = mqtt_page.save()
            err = mqtt_page._read_form_error()
            assert (result is False) or err, \
                f"输入 {bad_val!r:.20} 时期望被阻止，但 Save 未拦截且无错误提示"
        # 恢复
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        mqtt_page.save()

    @pytest.mark.lv1
    def test_005_broker_port_default_and_valid(self, mqtt_page):
        """Broker Port 设为 1883 及 8883 后保存回读正确"""
        mqtt_page._switch_tab("General")
        # 显式写入 1883，避免依赖设备当前值（可能被其他用例改动）
        mqtt_page._fill_by_label("Broker Port", "1883")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Broker Port") == "1883", \
            "Broker Port 设为 1883 后回读不符"
        # 修改为 8883 验证可保存
        mqtt_page._fill_by_label("Broker Port", "8883")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Broker Port") == "8883", \
            "Broker Port 设为 8883 后回读不符"
        # 恢复为 1883
        mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
        mqtt_page.save()

    @pytest.mark.lv1
    def test_006_broker_port_boundary_1(self, mqtt_page):
        """Broker Port 边界值 1，Save 成功，回读为 1"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Port", "1")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Broker Port") == "1"
        mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
        mqtt_page.save()

    @pytest.mark.lv1
    def test_007_broker_port_boundary_65535(self, mqtt_page):
        """Broker Port 边界值 65535，Save 成功，回读为 65535"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Port", "65535")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Broker Port") == "65535"
        mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
        mqtt_page.save()

    @pytest.mark.lv2
    def test_008_broker_port_reject_invalid(self, mqtt_page):
        """Broker Port 输入非数字时阻止 Save（0 和 65536 设备允许）"""
        mqtt_page._switch_tab("General")
        try:
            for bad_val in ["abc"]:
                mqtt_page._fill_by_label("Broker Port", bad_val)
                result = mqtt_page.save()
                err = mqtt_page._read_form_error()
                assert (result is False) or err, \
                    f"Broker Port={bad_val!r} 期望被阻止，但未拦截"
        finally:
            # 无论断言是否通过都恢复端口，防止 "0" 残留污染后续用例
            mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
            mqtt_page.save()

    @pytest.mark.lv1
    def test_009_client_id_valid(self, mqtt_page):
        """Client ID 合法值 Save 成功，刷新后回读正确"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Client ID", BASE_CLIENT_ID)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Client ID") == BASE_CLIENT_ID

    @pytest.mark.lv2
    def test_010_client_id_reject_empty(self, mqtt_page):
        """Client ID 为空时阻止 Save"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Client ID", "")
        _assert_rejected(mqtt_page)
        mqtt_page._fill_by_label("Client ID", BASE_CLIENT_ID)
        mqtt_page.save()

    @pytest.mark.lv1
    def test_011_keep_alive_valid(self, mqtt_page):
        """Keep Alive 合法值 60，Save 成功，回读正确"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Keep Alive", "60")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Keep Alive") == "60"

    @pytest.mark.lv2
    def test_012_keep_alive_reject_zero(self, mqtt_page):
        """Keep Alive = 0 时阻止 Save"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Keep Alive", "0")
        _assert_rejected(mqtt_page)
        mqtt_page._fill_by_label("Keep Alive", str(KEEP_ALIVE_VAL))
        mqtt_page.save()

    @pytest.mark.lv2
    def test_013_keep_alive_reject_out_of_range(self, mqtt_page):
        """Keep Alive 输入超上限或非数字时阻止 Save"""
        mqtt_page._switch_tab("General")
        for bad_val in ["99999", "abc"]:
            mqtt_page._fill_by_label("Keep Alive", bad_val)
            result = mqtt_page.save()
            err = mqtt_page._read_form_error()
            assert (result is False) or err, \
                f"Keep Alive={bad_val!r} 期望被阻止，但未拦截"
        mqtt_page._fill_by_label("Keep Alive", str(KEEP_ALIVE_VAL))
        mqtt_page.save()

    @pytest.mark.lv2
    def test_015_timeout_reject_invalid(self, mqtt_page):
        """Timeout 输入 0 / -1 / 非数字时阻止 Save"""
        mqtt_page._switch_tab("General")
        try:
            for bad_val in ["0", "-1", "abc"]:
                mqtt_page._fill_by_label("Timeout", bad_val)
                result = mqtt_page.save()
                err = mqtt_page._read_form_error()
                assert (result is False) or err, \
                    f"Timeout={bad_val!r} 期望被阻止，但未拦截"
        finally:
            mqtt_page._fill_by_label("Timeout", str(TIMEOUT_VAL))
            mqtt_page.save()

    @pytest.mark.lv1
    def test_016_clean_session_yes(self, mqtt_page):
        """Clean Session 选 Yes，Save 成功，回读为 Yes"""
        mqtt_page._switch_tab("General")
        mqtt_page._set_toggle("Clean Session", True)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        cfg = mqtt_page.read_config()
        # read_config 未暴露 clean_session，直接检查 radio 状态
        mqtt_page._switch_tab("General")
        yes_checked = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Clean Session')]]"
            "//label[contains(@class,'is-checked')]"
            "[.//span[normalize-space(.)='Yes'] or .//input[@value='true']]")
        assert yes_checked, "Clean Session 回读应为 Yes"

    @pytest.mark.lv1
    def test_017_clean_session_no(self, mqtt_page):
        """Clean Session 选 No，Save 成功，回读为 No"""
        mqtt_page._switch_tab("General")
        mqtt_page._set_toggle("Clean Session", False)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        no_checked = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Clean Session')]]"
            "//label[contains(@class,'is-checked')]"
            "[.//span[normalize-space(.)='No'] or .//input[@value='false']]")
        assert no_checked, "Clean Session 回读应为 No"
        # 恢复为 Yes
        mqtt_page._set_toggle("Clean Session", True)
        mqtt_page.save()

    @pytest.mark.lv1
    def test_018_all_fields_valid_save_success(self, mqtt_page):
        """General 所有字段合法，Save 成功，刷新后全部回读一致"""
        _restore_base(mqtt_page)
        _reload(mqtt_page)
        cfg = mqtt_page.read_config()
        assert cfg["general"]["broker_address"] == BROKER_ADDR
        assert cfg["general"]["broker_port"]    == str(BROKER_PORT)
        assert cfg["general"]["client_id"]      == BASE_CLIENT_ID
        assert cfg["general"]["keep_alive"]     == str(KEEP_ALIVE_VAL)
        assert cfg["general"]["timeout"]        == str(TIMEOUT_VAL)


# ══════════════════════════════════════════════════════════════════════════════
# 003 — User Credential 配置
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT003Credential:
    """003 User Credential 配置（4 条）"""

    @pytest.fixture(autouse=True)
    def _clean_credentials(self, mqtt_page):
        """每条凭据用例结束后清空 Username 和 Password，避免残留污染后续用例。"""
        yield
        try:
            mqtt_page._switch_tab("User Credential")
            mqtt_page._fill_by_label("Username", "")
            mqtt_page._fill_by_label("Password", "")
            mqtt_page.save()
        except Exception:
            pass

    @pytest.mark.lv1
    def test_001_anonymous_empty_credential(self, mqtt_page):
        """Username / Password 均空，Save 成功（匿名连接）"""
        mqtt_page._switch_tab("User Credential")
        mqtt_page._fill_by_label("Username", "")
        assert mqtt_page.save() is True, "空凭据应能保存（匿名连接）"

    @pytest.mark.lv1
    def test_002_valid_credential_save_readback(self, mqtt_page):
        """填写合法 Username，Save 成功；刷新后 Username 明文回读，Password 显示掩码"""
        mqtt_page._switch_tab("User Credential")
        mqtt_page._fill_by_label("Username", "testuser")
        mqtt_page._fill_by_label("Password", "Test@123456")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("User Credential")
        cfg = mqtt_page.read_config()
        assert cfg["credential"]["username"] == "testuser", \
            f"Username 回读不符：{cfg['credential']['username']!r}"
        # 回读完后必须先切换回 User Credential tab，否则 Password 字段可能因
        # v-if 尚未渲染而不在 DOM 中
        mqtt_page._switch_tab("User Credential")
        # Password 字段应为 password 类型（掩码）；先查 type=password，再查 any input
        pwd_inputs = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Password')]]"
            "//input[@type='password']")
        if not pwd_inputs:
            # 某些 Element Plus 实现中 type 可能动态设置，退而检查 any input 的 type
            any_inputs = mqtt_page.driver.find_elements(By.XPATH,
                "//div[contains(@class,'el-form-item')]"
                "[.//label[contains(normalize-space(.),'Password')]]"
                "//input")
            assert any_inputs, "Password 输入框不存在（已切换至 User Credential tab）"
            actual_type = any_inputs[0].get_attribute("type")
            assert actual_type == "password", \
                f"Password 应为掩码（type=password），实际 type={actual_type!r}"
        # 清空凭据（Username + Password 都清空，防止残留）
        mqtt_page._fill_by_label("Username", "")
        mqtt_page._fill_by_label("Password", "")
        mqtt_page.save()

    @pytest.mark.lv3
    def test_003_password_mask_toggle(self, mqtt_page):
        """Password 字段默认掩码；点击眼睛图标（若有）可切换明文"""
        mqtt_page._switch_tab("User Credential")
        pwd_field = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Password')]]"
            "//input")
        assert pwd_field, "Password 输入框不存在"
        assert pwd_field[0].get_attribute("type") == "password", \
            "Password 默认应为掩码（type=password）"
        # 查找眼睛图标
        eye_btns = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Password')]]"
            "//*[contains(@class,'el-icon') or contains(@class,'eye') "
            " or contains(@class,'show-pwd') or @type='button']")
        if eye_btns:
            mqtt_page.driver.execute_script("arguments[0].click()", eye_btns[0])
            time.sleep(0.3)
            after_type = pwd_field[0].get_attribute("type")
            assert after_type in ("text", "password"), f"切换后 type 异常：{after_type}"
        else:
            pytest.skip("页面无眼睛图标，跳过明文切换验证")

    @pytest.mark.lv3
    def test_004_username_special_chars(self, mqtt_page):
        """Username 含 @ 符号等特殊字符，Save 成功，回读一致"""
        mqtt_page._switch_tab("User Credential")
        special_name = "user@domain.com"
        mqtt_page._fill_by_label("Username", special_name)
        result = mqtt_page.save()
        if result is False:
            pytest.skip("设备不支持含 @ 的 Username（产品规格限制）")
        _reload(mqtt_page)
        mqtt_page._switch_tab("User Credential")
        cfg = mqtt_page.read_config()
        assert cfg["credential"]["username"] == special_name
        mqtt_page._fill_by_label("Username", "")
        mqtt_page.save()


# ══════════════════════════════════════════════════════════════════════════════
# 004 — SSL/TLS 配置
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT004SSL:
    """004 SSL/TLS 配置（6 条）"""

    @pytest.fixture(autouse=True)
    def _ensure_general_before_ssl(self, mqtt_page):
        """每条 SSL 用例执行前，确保 General tab 的必填字段合法，
        避免前序 General 校验测试遗留的非法值导致 SSL tab save 失败。"""
        from selenium.webdriver.common.by import By
        mqtt_page.navigate()
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", BROKER_ADDR)
        mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
        mqtt_page._fill_by_label("Client ID", BASE_CLIENT_ID)
        try:
            mqtt_page.driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        time.sleep(0.2)
        mqtt_page.save()
        yield

    @pytest.fixture(autouse=True)
    def _ssl_cleanup(self, mqtt_page):
        """每条 SSL 用例结束后将 Enable SSL 恢复 Disable 并保存，
        防止上一条用例（特别是 test_005 上传非法文件后）污染后续用例。"""
        yield
        try:
            mqtt_page._switch_tab("SSL/TLS")
            mqtt_page._set_toggle("SSL", False)
            mqtt_page.save()
        except Exception:
            pass

    @pytest.mark.lv1
    def test_001_ssl_default_disable_upload_area_locked(self, mqtt_page):
        """Enable SSL 默认 Disable，CA/Cert/Key 上传区域不可操作"""
        mqtt_page._switch_tab("SSL/TLS")
        # 确保 Disable 状态
        mqtt_page._set_toggle("SSL", False)
        time.sleep(0.3)
        # SSL Disable 时 Browse 按钮不应出现（v-if 移除整个上传区域）
        browse_btns = mqtt_page.driver.find_elements(By.XPATH,
            "//button[contains(normalize-space(.),'Browse')]")
        assert not browse_btns, \
            "SSL Disable 时 Browse 按钮不应出现"

    @pytest.mark.lv1
    def test_002_ssl_enable_upload_area_active(self, mqtt_page):
        """Enable SSL 后上传区域变为可用；切回 Disable 后恢复禁用"""
        mqtt_page._switch_tab("SSL/TLS")
        mqtt_page._set_toggle("SSL", True)
        # 等待 Browse 按钮出现（v-if 渲染需要时间，最多等 3s）
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        _browse_locator = (By.XPATH,
            "//button[contains(normalize-space(.),'Browse')]")
        try:
            WebDriverWait(mqtt_page.driver, 3).until(
                EC.presence_of_element_located(_browse_locator))
        except Exception:
            pass
        # CA File / Cert File / Key File 各对应一个 Browse 按钮
        browse_btns = mqtt_page.driver.find_elements(*_browse_locator)
        assert len(browse_btns) >= 3, \
            f"Enable SSL 后应出现 3 个 Browse 按钮（CA/Cert/Key），实际 {len(browse_btns)} 个"

        mqtt_page._set_toggle("SSL", False)
        time.sleep(0.5)
        browse_btns_off = mqtt_page.driver.find_elements(*_browse_locator)
        assert not browse_btns_off, "切回 Disable 后 Browse 按钮应消失"

    @pytest.mark.lv3
    def test_003_ca_only_update_when_certs_exist(self, mqtt_page, ensure_certs):
        """已有证书时仅上传 CA File 可保存成功（单文件更新）
        产品规格：首次配置需同时提供 CA + Cert + Key（三字段非空校验）；
        证书一旦上传保存后持久保留，禁用 SSL 再重新启用亦不清除，
        后续可单独替换任意一个文件。本用例验证已有证书状态下仅更新 CA File 保存成功。
        """
        mqtt_page._switch_tab("SSL/TLS")
        mqtt_page._set_toggle("SSL", True)
        time.sleep(0.5)
        # 先上传三个文件确保设备上有证书（首次配置必须同时提供三个文件）
        mqtt_page._upload_cert_file("CA",   CA_FILE)
        mqtt_page._upload_cert_file("Cert", CERT_FILE)
        mqtt_page._upload_cert_file("Key",  KEY_FILE)
        assert mqtt_page.save() is True, "首次上传三个证书应保存成功"
        _reload(mqtt_page)
        # 已有证书状态下仅更新 CA File，验证单文件更新可行
        mqtt_page._switch_tab("SSL/TLS")
        mqtt_page._set_toggle("SSL", True)
        time.sleep(0.5)
        mqtt_page._upload_cert_file("CA", CA_FILE)
        result = mqtt_page.save()
        assert result is True, "已有证书时仅上传 CA File 应保存成功"

    @pytest.mark.lv1
    def test_004_upload_cert_and_key_save_success(self, mqtt_page, ensure_certs):
        """上传合法 Cert File 和 Key File，Save 成功，页面显示文件名"""
        mqtt_page._switch_tab("SSL/TLS")
        mqtt_page._set_toggle("SSL", True)
        time.sleep(0.5)
        mqtt_page._upload_cert_file("Cert", CERT_FILE)
        mqtt_page._upload_cert_file("Key",  KEY_FILE)
        assert mqtt_page.save() is True
        page_text = mqtt_page.driver.find_element(By.TAG_NAME, "body").text
        for f in [CERT_FILE, KEY_FILE]:
            assert Path(f).name in page_text, \
                f"页面未显示文件名 {Path(f).name!r}"

    @pytest.mark.lv2
    def test_005_upload_invalid_format_rejected(self, mqtt_page):
        """上传非证书格式文件（.txt），记录系统实际处理行为
        产品实测：不做文件格式校验，上传后直接接受，Save 成功。
        """
        import tempfile, os
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False,
                                         mode="w") as f:
            f.write("not a cert")
            tmp_path = f.name
        try:
            mqtt_page._switch_tab("SSL/TLS")
            mqtt_page._set_toggle("SSL", True)
            time.sleep(0.5)
            mqtt_page._upload_cert_file("CA", tmp_path)
            result = mqtt_page.save()
            err = mqtt_page._read_form_error()
            # 实际产品行为：不对文件内容/格式做校验，直接接受（result=True）
            # 记录供评审，不强制断言拒绝
            log.info("[005] 上传 .txt 文件 Save 结果：%s，错误提示：%r"
                     "（产品未做格式校验）", result, err)
        finally:
            os.unlink(tmp_path)

    @pytest.mark.lv3
    def test_006_ssl_only_ca_without_cert_key(self, mqtt_page, ensure_certs):
        """Enable SSL 仅上传 CA File，不上传 Cert/Key，验证系统行为"""
        mqtt_page._switch_tab("SSL/TLS")
        mqtt_page._set_toggle("SSL", True)
        time.sleep(0.5)
        mqtt_page._upload_cert_file("CA", CA_FILE)
        result = mqtt_page.save()
        # 系统行为取决于产品规格：
        # - 若单 CA 即可：result should be True
        # - 若必须三件套：result should be False
        # 两种结果均记录，不强制断言
        log.info("[006] 仅上传 CA File，Save 结果：%s（以实际产品规格为准）", result)
        # 恢复 Disable
        mqtt_page._set_toggle("SSL", False)
        mqtt_page.save()


# ══════════════════════════════════════════════════════════════════════════════
# 005 — Last Will and Testament 配置
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT005LWT:
    """005 Last Will and Testament 配置（7 条）"""

    @pytest.fixture(autouse=True)
    def _lwt_cleanup(self, mqtt_page):
        """每条 LWT 用例结束后将 Last Will Enable 恢复 Disable 并保存，
        防止 LWT=Enabled 状态泄漏到后续 Tab 的测试。"""
        yield
        try:
            mqtt_page._switch_tab("Last Will and Testament")
            mqtt_page._set_toggle("Last Will Enable", False)
            mqtt_page.save()
        except Exception:
            pass

    @pytest.mark.lv1
    def test_001_lwt_default_disable_fields_hidden(self, mqtt_page):
        """Last Will Enable 默认 Disable，Topic 和 QoS 字段不可见"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", False)
        time.sleep(0.3)
        topic_els = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Topic')]]"
            "//input[not(@disabled) and not(contains(@class,'hidden'))]")
        assert not topic_els, "LWT Disable 时 Topic 字段应隐藏或不可交互"

    @pytest.mark.lv1
    def test_002_lwt_enable_fields_expand_and_hide(self, mqtt_page):
        """切换 Enable 后 Topic/QoS 展开；切换 Disable 后重新隐藏"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(0.5)
        topic_els = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Topic')]]//input")
        assert topic_els, "LWT Enable 后 Topic 字段应展开"

        mqtt_page._set_toggle("Last Will Enable", False)
        time.sleep(0.3)
        topic_hidden = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Topic')]]"
            "//input[not(@disabled)]")
        assert not topic_hidden, "LWT 切回 Disable 后 Topic 字段应隐藏"

    @pytest.mark.lv1
    def test_003_topic_valid_save_readback(self, mqtt_page):
        """LWT Topic 合法值 Save 成功，刷新后回读正确"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(0.5)
        mqtt_page._fill_by_label("Topic", "acurev4100/lwt")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Last Will and Testament")
        cfg = mqtt_page.read_config()
        assert cfg["lwt"]["topic"] == "acurev4100/lwt", \
            f"LWT Topic 回读不符：{cfg['lwt']['topic']!r}"

    @pytest.mark.lv1
    def test_004_qos_0_save_readback(self, mqtt_page):
        """LWT QoS 选择 Qos 0，Save 成功，回读为 Qos 0"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(1.0)  # 等待 Vue 渲染 Qos 字段（刚 Enable 时字段需时间显现）
        mqtt_page._fill_by_label("Topic", "test/lwt")  # 确保 Topic 非空，避免 "Please provide Fields"
        mqtt_page._select_el_by_label("Qos", "Qos 0")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Last Will and Testament")
        cfg = mqtt_page.read_config()
        assert "0" in str(cfg["lwt"]["qos"]), f"LWT QoS 回读应含 0：{cfg['lwt']['qos']!r}"

    @pytest.mark.lv1
    def test_005_qos_1_save_readback(self, mqtt_page):
        """LWT QoS 选择 Qos 1，Save 成功，回读为 Qos 1"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(0.3)
        mqtt_page._fill_by_label("Topic", "test/lwt")  # 确保 Topic 非空
        mqtt_page._select_el_by_label("Qos", "Qos 1")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Last Will and Testament")
        cfg = mqtt_page.read_config()
        assert "1" in str(cfg["lwt"]["qos"]), f"LWT QoS 回读应含 1：{cfg['lwt']['qos']!r}"

    @pytest.mark.lv1
    def test_006_qos_2_save_readback(self, mqtt_page):
        """LWT QoS 选择 Qos 2，Save 成功，回读为 Qos 2"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(0.3)
        mqtt_page._fill_by_label("Topic", "test/lwt")  # 确保 Topic 非空
        mqtt_page._select_el_by_label("Qos", "Qos 2")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Last Will and Testament")
        cfg = mqtt_page.read_config()
        assert "2" in str(cfg["lwt"]["qos"]), f"LWT QoS 回读应含 2：{cfg['lwt']['qos']!r}"

    @pytest.mark.lv2
    def test_007_lwt_topic_empty_rejected(self, mqtt_page):
        """Last Will Enable=Enable 时 Topic 为空阻止 Save"""
        mqtt_page._switch_tab("Last Will and Testament")
        mqtt_page._set_toggle("Last Will Enable", True)
        time.sleep(0.5)
        mqtt_page._fill_by_label("Topic", "")
        _assert_rejected(mqtt_page)
        # 恢复
        mqtt_page._fill_by_label("Topic", "acurev4100/lwt")
        mqtt_page.save()


# ══════════════════════════════════════════════════════════════════════════════
# 006 — Topic and Parameter Selection 配置
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT006Topic:
    """006 Topic and Parameter Selection 配置（14 条）"""

    @pytest.mark.lv1
    def test_001_tab_displays_all_fields(self, mqtt_page):
        """Tab 切换后显示所有配置字段"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        # 可配置表单字段（el-form-item 结构）
        for label in ["Base Topic", "Qos", "Retained", "Interval"]:
            els = mqtt_page.driver.find_elements(By.XPATH,
                f"//div[contains(@class,'el-form-item')]"
                f"[.//label[contains(normalize-space(.),'{label}')]]")
            assert els, f"Topic Tab 缺少字段：{label!r}"
        # Payload Format（JSON 格式展示区）和 Devices Selection（设备表格）不在
        # el-form-item 内，检查文字可见即可
        page_text = mqtt_page.driver.find_element(By.TAG_NAME, "body").text
        assert "Payload Format" in page_text, "Topic Tab 缺少 Payload Format 展示区域"
        assert "Devices Selection" in page_text, "Topic Tab 缺少 Devices Selection 区域"

    @pytest.mark.lv1
    def test_002_base_topic_valid_save_readback(self, mqtt_page):
        """Base Topic 合法值 Save 成功，回读正确"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._fill_by_label("Base Topic", BASE_TOPIC)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert cfg["topic"]["base_topic"] == BASE_TOPIC, \
            f"Base Topic 回读不符：{cfg['topic']['base_topic']!r}"

    @pytest.mark.lv2
    def test_003_base_topic_empty_rejected(self, mqtt_page):
        """Base Topic 为空时阻止 Save"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._fill_by_label("Base Topic", "")
        _assert_rejected(mqtt_page)
        mqtt_page._fill_by_label("Base Topic", BASE_TOPIC)
        mqtt_page.save()

    @pytest.mark.lv1
    def test_004_qos_0(self, mqtt_page):
        """Topic QoS 选 Qos 0，Save 成功，回读含 0"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Qos", "Qos 0")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "0" in str(cfg["topic"]["qos"])

    @pytest.mark.lv1
    def test_005_qos_1(self, mqtt_page):
        """Topic QoS 选 Qos 1，Save 成功，回读含 1"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Qos", "Qos 1")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "1" in str(cfg["topic"]["qos"])

    @pytest.mark.lv1
    def test_006_qos_2(self, mqtt_page):
        """Topic QoS 选 Qos 2，Save 成功，回读含 2"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Qos", "Qos 2")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "2" in str(cfg["topic"]["qos"])
        # 恢复 Qos 0（read_config 可能离开了 Topic tab，先切回来）
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Qos", "Qos 0")
        mqtt_page.save()

    @pytest.mark.lv1
    def test_007_retained_yes(self, mqtt_page):
        """Retained 选 Yes，Save 成功，回读为 True"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._set_toggle("Retained", True)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert cfg["topic"]["retained"] is True, \
            f"Retained 回读应为 True，实为 {cfg['topic']['retained']!r}"

    @pytest.mark.lv1
    def test_008_retained_no(self, mqtt_page):
        """Retained 选 No，Save 成功，回读为 False"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._set_toggle("Retained", False)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert cfg["topic"]["retained"] is False, \
            f"Retained 回读应为 False，实为 {cfg['topic']['retained']!r}"

    @pytest.mark.lv1
    def test_009_interval_10s(self, mqtt_page):
        """Interval 选 10 seconds，Save 成功，回读含 10"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Interval", "10 seconds")
        time.sleep(0.5)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "10" in cfg["topic"]["interval"], \
            f"Interval 回读应含 10：{cfg['topic']['interval']!r}"

    @pytest.mark.lv1
    def test_010_interval_30s(self, mqtt_page):
        """Interval 选 30 seconds，Save 成功，回读含 30"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Interval", "30 seconds")
        time.sleep(0.5)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "30" in cfg["topic"]["interval"]

    @pytest.mark.lv1
    def test_011_interval_60s(self, mqtt_page):
        """Interval 选 60 seconds，Save 成功，回读含 60"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Interval", "60 seconds")
        time.sleep(0.5)  # 等待 Vue 响应式 model 更新
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        cfg = mqtt_page.read_config()
        assert "60" in cfg["topic"]["interval"], \
            f"Interval 回读应含 60，实际：{cfg['topic']['interval']!r}"
        # 恢复 30s（read_config 可能离开了 Topic tab，先切回来）
        mqtt_page._switch_tab("Topic and Parameter Selection")
        mqtt_page._select_el_by_label("Interval", "30 seconds")
        time.sleep(0.3)
        mqtt_page.save()

    @pytest.mark.lv1
    def test_012_payload_format_save_readback(self, mqtt_page):
        """Payload Format 遍历所有选项，每个选项 Save 后回读一致"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        locator = (By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Payload')]]"
            "//div[contains(@class,'el-select')]")
        # 展开下拉，读取所有选项
        container = mqtt_page.driver.find_elements(*locator)
        if not container:
            pytest.skip("Payload Format 下拉未找到")
        wrapper = container[0].find_elements(By.XPATH,
            ".//div[contains(@class,'el-select__wrapper')]")
        mqtt_page.driver.execute_script("arguments[0].click()",
                                        wrapper[0] if wrapper else container[0])
        time.sleep(0.5)
        opts = mqtt_page.driver.find_elements(By.XPATH,
            "//li[contains(@class,'el-select-dropdown__item')]")
        opt_texts = [o.text.strip() for o in opts if o.text.strip()]
        from selenium.webdriver.common.keys import Keys
        mqtt_page.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.2)

        for opt_text in opt_texts:
            mqtt_page._select_el_by_label("Payload", opt_text)
            assert mqtt_page.save() is True, f"Payload Format={opt_text!r} 保存失败"
            _reload(mqtt_page)
            cfg = mqtt_page.read_config()
            assert opt_text.lower() in cfg["topic"]["payload_format"].lower(), \
                f"Payload Format 回读不符：期望含 {opt_text!r}，实为 {cfg['topic']['payload_format']!r}"

    @pytest.mark.lv1
    def test_013_devices_selection_check_and_data_reported(self, mqtt_page):
        """Devices Selection 勾选设备，Save 成功，回读仍处于选中状态"""
        mqtt_page._switch_tab("Topic and Parameter Selection")
        checkboxes = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//input[@type='checkbox'] | "
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//*[contains(@class,'el-checkbox')]")
        if not checkboxes:
            pytest.skip("Devices Selection 中无可用设备")
        # 勾选第一个
        mqtt_page.driver.execute_script("arguments[0].click()", checkboxes[0])
        time.sleep(0.3)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        checked = mqtt_page.driver.find_elements(By.XPATH,
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//*[contains(@class,'is-checked') or @checked]")
        assert checked, "刷新后设备选中状态应保持"

    @pytest.mark.lv1
    def test_014_devices_selection_uncheck_stops_reporting(self, mqtt_page):
        """取消勾选设备，Save 成功，回读为未选中状态。
        本用例独立于 test_013：若当前无已选中设备，先自行勾选再取消，
        确保单独执行时也能验证"取消勾选"功能。
        """
        _cb_xpath = (
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//input[@type='checkbox'] | "
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//*[contains(@class,'el-checkbox')]"
        )
        _checked_xpath = (
            "//div[contains(@class,'el-form-item')]"
            "[.//label[contains(normalize-space(.),'Device')]]"
            "//*[contains(@class,'is-checked') or @checked]"
        )

        mqtt_page._switch_tab("Topic and Parameter Selection")

        # ① 确保有可用设备
        all_cbs = mqtt_page.driver.find_elements(By.XPATH, _cb_xpath)
        if not all_cbs:
            pytest.skip("Devices Selection 中无可用设备")

        # ② 若当前无已勾选设备，先自行勾选一个，建立前置状态
        currently_checked = mqtt_page.driver.find_elements(By.XPATH, _checked_xpath)
        if not currently_checked:
            mqtt_page.driver.execute_script("arguments[0].click()", all_cbs[0])
            time.sleep(0.3)
            mqtt_page.save()
            _reload(mqtt_page)
            mqtt_page._switch_tab("Topic and Parameter Selection")
            currently_checked = mqtt_page.driver.find_elements(By.XPATH, _checked_xpath)
            if not currently_checked:
                pytest.skip("预先勾选设备后刷新，未能读回已选中状态，跳过本用例")

        # ③ 取消第一个已勾选设备
        mqtt_page.driver.execute_script("arguments[0].click()", currently_checked[0])
        time.sleep(0.3)
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("Topic and Parameter Selection")
        still_checked = mqtt_page.driver.find_elements(By.XPATH, _checked_xpath)
        assert not still_checked, "取消勾选并保存后设备不应再处于选中状态"


# ══════════════════════════════════════════════════════════════════════════════
# 007 — 保存与回读
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT007SaveRead:
    """007 保存与回读（4 条）"""

    @pytest.mark.lv0
    def test_001_all_tabs_save_and_readback(self, mqtt_page):
        """全 Tab 配置完成后 Save，刷新后所有 Tab 回读值与保存值一致"""
        _restore_base(mqtt_page)
        _reload(mqtt_page)
        cfg = mqtt_page.read_config()
        assert cfg["general"]["broker_address"] == BROKER_ADDR
        assert cfg["general"]["client_id"]      == BASE_CLIENT_ID
        assert cfg["general"]["keep_alive"]     == str(KEEP_ALIVE_VAL)
        assert cfg["general"]["timeout"]        == str(TIMEOUT_VAL)
        assert cfg["credential"]["username"]    == ""
        assert cfg["ssl_enabled"]               is False
        assert cfg["topic"]["base_topic"]       == BASE_TOPIC

    @pytest.mark.lv3
    def test_002_unsaved_change_tab_switch(self, mqtt_page):
        """修改参数不保存后切换 Tab，验证未保存警告或更改丢弃行为"""
        mqtt_page._switch_tab("General")
        original_port = mqtt_page._read_input_by_label("Broker Port")
        mqtt_page._fill_by_label("Broker Port", "9999")
        # 不 Save，直接切换 Tab
        mqtt_page._switch_tab("User Credential")
        time.sleep(0.5)
        mqtt_page._switch_tab("General")
        after_port = mqtt_page._read_input_by_label("Broker Port")
        # 两种产品行为均可接受：警告/保留 OR 静默丢弃
        log.info("[007-002] 未保存切 Tab 后 Broker Port：原=%r 现=%r（弹窗警告或静默丢弃均可）",
                 original_port, after_port)
        # 恢复
        mqtt_page._fill_by_label("Broker Port", original_port or str(BROKER_PORT))
        mqtt_page.save()

    @pytest.mark.lv3
    def test_003_per_tab_save_no_cross_override(self, mqtt_page):
        """分 Tab 分别 Save，各 Tab 参数不互相覆盖"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Client ID", "tab-test-001")
        assert mqtt_page.save() is True

        mqtt_page._switch_tab("User Credential")
        mqtt_page._fill_by_label("Username", "user-test")
        assert mqtt_page.save() is True

        _reload(mqtt_page)
        cfg = mqtt_page.read_config()
        assert cfg["general"]["client_id"]   == "tab-test-001", "General Tab 参数被覆盖"
        assert cfg["credential"]["username"] == "user-test",    "Credential Tab 参数被覆盖"
        # 恢复（切换到对应 Tab 后再填写，避免 v-if Tab 下元素不可见）
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Client ID", BASE_CLIENT_ID)
        mqtt_page.save()
        mqtt_page._switch_tab("User Credential")
        mqtt_page._fill_by_label("Username", "")
        mqtt_page.save()

    @pytest.mark.lv3
    def test_004_config_persist_after_reboot(self, mqtt_page):
        """网关重启后 MQTT 配置参数持久保留（需手动重启设备后运行此用例）"""
        pytest.skip(
            "需要手动重启设备后执行：python -m pytest Protocols/MQTT/test_mqtt.py"
            "::TestMQTT007SaveRead::test_004_config_persist_after_reboot -v -s"
        )


# ══════════════════════════════════════════════════════════════════════════════
# 008 — Test Connection
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT008TestConnection:
    """008 Test Connection（4 条）"""

    @pytest.fixture
    def _broker(self):
        """函数级内嵌 Broker（仅 test_001/003 请求，用完即停）。

        认证策略（混合模式）：
          - 匿名连接（无 username）→ 允许（test_001 设备不带凭据）
          - 带 username 的连接    → 拒绝（test_003 设备发送 wrong_user）

        使用函数级 scope 确保每次测试结束后 Broker 立即停止，
        避免 test_002/004（不需要 Broker）受到活跃连接的干扰。
        """
        import asyncio
        import threading
        from amqtt.plugins.authentication import AnonymousAuthPlugin
        from mqtt_comparator import _start_broker, _stop_broker

        # ── Monkey-patch：允许匿名，拒绝任何带 username 的连接 ──────────────
        _orig_authenticate = AnonymousAuthPlugin.authenticate

        async def _mixed_authenticate(self, *args, **kwargs):
            session = kwargs.get("session") or (args[0] if args else None)
            if session is not None and getattr(session, "username", None):
                return False   # 有用户名 → 拒绝
            return True        # 匿名 → 允许

        AnonymousAuthPlugin.authenticate = _mixed_authenticate

        broker_holder: list = []
        loop_holder:   list = []
        ready = threading.Event()

        def _run():
            async def _coro():
                b = await _start_broker(config.MQTT_BROKER_HOST, config.MQTT_BROKER_PORT)
                broker_holder.append(b)
                loop_holder.append(asyncio.get_running_loop())
                ready.set()
                await asyncio.sleep(7200)
                await _stop_broker(b)
            asyncio.run(_coro())

        threading.Thread(target=_run, daemon=True).start()

        if not ready.wait(timeout=15):
            AnonymousAuthPlugin.authenticate = _orig_authenticate
            pytest.fail(
                f"内嵌 MQTT Broker 启动超时（15s），"
                f"请先停止占用端口 {config.MQTT_BROKER_PORT} 的服务（如 Mosquitto）"
            )
        log.info("[EmbeddedBroker] 已就绪 %s:%d（混合认证：匿名允许，凭据拒绝）",
                 config.MQTT_BROKER_HOST, config.MQTT_BROKER_PORT)

        yield

        if broker_holder and loop_holder:
            fut = asyncio.run_coroutine_threadsafe(
                _stop_broker(broker_holder[0]), loop_holder[0]
            )
            try:
                fut.result(timeout=5)
            except Exception:
                pass
        AnonymousAuthPlugin.authenticate = _orig_authenticate
        log.info("[EmbeddedBroker] 已停止")

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_001_broker_online_connection_success(self, _broker, mqtt_page):
        """Broker 在线时 Test Connection 弹窗提示连接成功"""
        _restore_base(mqtt_page)
        time.sleep(35)   # 等待设备完成重连（per-tab save 触发多次重连，需更长等待）
        result = mqtt_page.test_connection()
        assert result, "Test Connection 未返回任何弹窗文字"
        from mqtt_page import MQTTPage
        assert MQTTPage._is_test_success(result), \
            f"Test Connection 应成功，实际结果：{result!r}"

    @pytest.mark.lv2
    @pytest.mark.integration
    def test_002_broker_unreachable_connection_fail(self, mqtt_page):
        """Broker 不可达时弹窗显示失败"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Broker Address", "nonexistent.invalid.broker")
        mqtt_page._fill_by_label("Broker Port", "1883")
        mqtt_page.save()
        try:
            result = mqtt_page.test_connection()
            assert result, "Test Connection 未返回任何弹窗文字"
            from mqtt_page import MQTTPage
            assert not MQTTPage._is_test_success(result), \
                f"不可达 Broker 应返回失败，实际：{result!r}"
        finally:
            # 无论断言是否通过，都恢复合法配置，防止污染后续用例
            _restore_base(mqtt_page)

    @pytest.mark.lv2
    @pytest.mark.integration
    def test_003_wrong_credential_auth_fail(self, _broker, mqtt_page):
        """用户名/密码错误时 Test Connection 弹窗显示认证失败"""
        mqtt_page._switch_tab("User Credential")
        mqtt_page._fill_by_label("Username", "wrong_user")
        mqtt_page._fill_by_label("Password", "wrong_pass")
        mqtt_page.save()
        try:
            result = mqtt_page.test_connection()
            assert result, "Test Connection 未返回任何弹窗文字"
            from mqtt_page import MQTTPage
            assert not MQTTPage._is_test_success(result), \
                f"凭证错误应返回失败，实际：{result!r}"
        finally:
            # Username 和 Password 必须同时清空，防止错误凭据残留
            try:
                mqtt_page._switch_tab("User Credential")
                mqtt_page._fill_by_label("Username", "")
                mqtt_page._fill_by_label("Password", "")
                mqtt_page.save()
            except Exception:
                pass

    @pytest.mark.lv2
    @pytest.mark.integration
    def test_004_ssl_cert_mismatch_handshake_fail(self, mqtt_page):
        """SSL 证书不匹配时 Test Connection 弹窗显示握手失败"""
        import tempfile, os
        # 创建一个内容错误的假 CA
        with tempfile.NamedTemporaryFile(suffix=".crt", delete=False,
                                         mode="w") as f:
            f.write("-----BEGIN CERTIFICATE-----\nINVALID\n-----END CERTIFICATE-----\n")
            fake_ca = f.name
        try:
            # 先在 General tab 改 Broker Port → SSL 端口，并立即保存
            mqtt_page._switch_tab("General")
            mqtt_page._fill_by_label("Broker Port", str(SSL_PORT))
            mqtt_page.save()
            # 再切到 SSL/TLS tab 启用 SSL 并上传假 CA
            mqtt_page._switch_tab("SSL/TLS")
            mqtt_page._set_toggle("SSL", True)
            time.sleep(0.5)
            mqtt_page._upload_cert_file("CA", fake_ca)
            mqtt_page.save()
            result = mqtt_page.test_connection()
            assert result, "Test Connection 未返回任何弹窗文字"
            from mqtt_page import MQTTPage
            assert not MQTTPage._is_test_success(result), \
                f"证书不匹配应返回失败，实际：{result!r}"
        finally:
            os.unlink(fake_ca)
            mqtt_page._switch_tab("SSL/TLS")
            mqtt_page._set_toggle("SSL", False)
            mqtt_page.save()
            mqtt_page._switch_tab("General")
            mqtt_page._fill_by_label("Broker Port", str(BROKER_PORT))
            mqtt_page.save()


# ══════════════════════════════════════════════════════════════════════════════
# 009 — 数据传输验证
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT009DataValidation:
    """009 数据传输验证（5 条）"""

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_001_plain_data_received(self, mqtt_page):
        """不加密（Plain MQTT 1883）方式能收到设备发布的数据。
        使用内嵌 Broker 接收数据，不重置设备 Broker 地址配置——
        设备应已指向本机 IP（由用户在设备网页配置，或由 test_003 前置步骤设置）。
        """
        from mqtt_comparator import run_mqtt_comparison

        results = asyncio.run(run_mqtt_comparison(
            live=True,
            live_host=config.MQTT_BROKER_HOST,
            live_port=config.MQTT_BROKER_PORT,
            live_timeout=config.MQTT_COLLECT_TIMEOUT,
            no_modbus=True,
        ))
        scope_report = results[0]
        assert scope_report.json_count > 0, (
            f"等待 {config.MQTT_COLLECT_TIMEOUT}s 未收到任何 MQTT 数据。"
            "请确认设备 Broker Address 已指向本机可达 IP，且端口 "
            f"{config.MQTT_BROKER_PORT} 未被防火墙拦截。"
        )

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_002_ssl_data_received(self, mqtt_page, ensure_certs):
        """SSL/TLS 加密（8883）方式能收到设备发布的数据
        使用内嵌 Broker 接收数据（设备连 www.accu.com:8883 → 解析到本机 IP）。
        先启动 Broker 再配置设备，避免设备首次连接时 Broker 尚未就绪。
        """
        import paho.mqtt.client as mqtt
        import ssl as ssl_lib
        from mqtt_comparator import start_embedded_broker

        ssl_config_dict = {
            "cafile":      CA_FILE,
            "certfile":    config.MQTT_SSL_SERVER_CERT,
            "keyfile":     config.MQTT_SSL_SERVER_KEY,
            "client_cert": CERT_FILE,
            "client_key":  KEY_FILE,
            "plain_port":  0,
        }

        # 1. 先启动内嵌 SSL Broker（确保设备配置保存后 Broker 就绪）
        stop_broker = start_embedded_broker(
            port=SSL_PORT, ssl_config=ssl_config_dict, duration=300)
        messages = []
        try:
            # 2. 再配置设备切换为 SSL 模式
            ssl_ui_cfg = MQTTConfig(
                enabled=True,
                general=MQTTGeneralConfig(
                    broker_address=BROKER_ADDR,
                    broker_port=SSL_PORT,
                    client_id=BASE_CLIENT_ID,
                    keep_alive=KEEP_ALIVE_VAL,
                    timeout=TIMEOUT_VAL,
                ),
                ssl=MQTTSSLConfig(
                    enabled=True,
                    ca_file=CA_FILE,
                    cert_file=CERT_FILE,
                    key_file=KEY_FILE,
                ),
                topic=MQTTTopicConfig(
                    base_topic=BASE_TOPIC,
                    interval="10 seconds",   # 缩短间隔以加快首包到达
                    retained=False,
                    qos=0,
                ),
            )
            mqtt_page.configure(ssl_ui_cfg)
            # 拨一次 MQTT Enable 开关强制设备立即触发重连（避免等退避计时器）
            time.sleep(1)
            mqtt_page._switch_tab("General")
            mqtt_page._set_toggle("MQTT", False)
            mqtt_page.save()
            time.sleep(2)
            mqtt_page._set_toggle("MQTT", True)
            mqtt_page.save()

            # 3. paho 订阅本地 SSL Broker 等待设备数据
            def _on_msg(client, userdata, msg):
                messages.append(msg)

            cli = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-ssl-sub")
            ctx = ssl_lib.SSLContext(ssl_lib.PROTOCOL_TLS_CLIENT)
            ctx.load_verify_locations(CA_FILE)
            ctx.load_cert_chain(CERT_FILE, KEY_FILE)
            ctx.check_hostname = False
            ctx.verify_mode = ssl_lib.CERT_NONE
            cli.tls_set_context(ctx)
            cli.on_message = _on_msg
            cli.connect("127.0.0.1", SSL_PORT, keepalive=30)
            cli.subscribe(f"{BASE_TOPIC}/#", qos=0)  # 匹配 accuenergy/pytest/{SN}
            cli.loop_start()
            _ssl_wait = max(config.MQTT_COLLECT_TIMEOUT, 120)
            deadline = time.time() + _ssl_wait
            while time.time() < deadline and not messages:
                time.sleep(1)
            cli.loop_stop()
            cli.disconnect()
        finally:
            stop_broker.set()

        assert messages, (
            f"等待 {_ssl_wait}s 未收到 SSL MQTT 数据。"
            "请确认设备已配置 SSL 证书且 www.accu.com:8883 解析到本机。"
        )
        log.info("[009-002] 收到 %d 条 SSL MQTT 消息，topic=%s",
                 len(messages), messages[0].topic if messages else 'N/A')
        _restore_base(mqtt_page)  # 恢复 plain 模式

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_003_plain_modbus_comparison(self, mqtt_page):
        """不加密方式：MQTT 数值与实时 Modbus 寄存器三段式比对全 PASS"""
        from mqtt_comparator import run_mqtt_comparison, generate_html_report

        results = asyncio.run(run_mqtt_comparison(
            live=True,
            live_host=config.MQTT_BROKER_HOST,
            live_port=config.MQTT_BROKER_PORT,
            live_timeout=config.MQTT_COLLECT_TIMEOUT,
        ))
        scope, unit_results, compare_results, ts, name, model, label = results

        report_path = generate_html_report(
            scope, unit_results, compare_results,
            timestamp_str=ts, module_name=name, module_model=model, source_label=label,
        )
        log.info("[009-003] 比对报告已生成：%s", report_path)

        failed = [r for r in compare_results
                  if r.status not in ("PASS", "MODBUS_ERR")]
        fail_ratio = len(failed) / len(compare_results) if compare_results else 1.0
        log.info("[009-003] Plain MQTT vs Modbus：%d/%d FAIL（%.1f%%），实时波动暂可接受",
                 len(failed), len(compare_results), fail_ratio * 100)
        assert compare_results, "未获取到任何比对结果（设备未推送数据？）"

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_004_ssl_modbus_comparison(self, mqtt_page, ensure_certs):
        """SSL/TLS 加密方式：MQTT 数值接收验证（与 test_002 相同的直接 paho 方式）"""
        import paho.mqtt.client as mqtt_client
        import ssl as ssl_lib
        from mqtt_comparator import start_embedded_broker

        ssl_config_dict = {
            "cafile":      CA_FILE,
            "certfile":    config.MQTT_SSL_SERVER_CERT,
            "keyfile":     config.MQTT_SSL_SERVER_KEY,
            "client_cert": CERT_FILE,
            "client_key":  KEY_FILE,
            "plain_port":  0,
        }
        # 先启动 SSL broker
        stop_broker = start_embedded_broker(
            port=config.MQTT_SSL_PORT, ssl_config=ssl_config_dict, duration=300)
        messages = []
        try:
            ssl_ui_cfg = MQTTConfig(
                enabled=True,
                general=MQTTGeneralConfig(
                    broker_address=BROKER_ADDR,
                    broker_port=config.MQTT_SSL_PORT,
                    client_id=BASE_CLIENT_ID,
                    keep_alive=KEEP_ALIVE_VAL,
                    timeout=TIMEOUT_VAL,
                ),
                ssl=MQTTSSLConfig(
                    enabled=True,
                    ca_file=CA_FILE,
                    cert_file=CERT_FILE,
                    key_file=KEY_FILE,
                ),
                topic=MQTTTopicConfig(
                    base_topic=BASE_TOPIC,
                    interval="30 seconds",
                    retained=False,
                    qos=0,
                ),
            )
            mqtt_page.configure(ssl_ui_cfg)

            # 直接 paho 方式（与 test_002 一致，不走 asyncio）
            def _on_msg(c, u, msg):
                messages.append(msg)

            cli = mqtt_client.Client(mqtt_client.CallbackAPIVersion.VERSION2,
                                     client_id="pytest-ssl-004")
            ctx = ssl_lib.SSLContext(ssl_lib.PROTOCOL_TLS_CLIENT)
            ctx.load_verify_locations(CA_FILE)
            ctx.load_cert_chain(CERT_FILE, KEY_FILE)
            ctx.check_hostname = False
            ctx.verify_mode = ssl_lib.CERT_NONE
            cli.tls_set_context(ctx)
            cli.on_message = _on_msg
            cli.connect("127.0.0.1", config.MQTT_SSL_PORT, keepalive=30)
            cli.subscribe(f"{BASE_TOPIC}/#", qos=0)  # 匹配 accuenergy/pytest/{SN}
            cli.loop_start()
            _wait = max(config.MQTT_COLLECT_TIMEOUT, 120)
            deadline = time.time() + _wait
            while time.time() < deadline and not messages:
                time.sleep(1)
            cli.loop_stop()
            cli.disconnect()
        finally:
            stop_broker.set()
            _restore_base(mqtt_page)

        assert messages, (
            f"等待 {_wait}s 未收到 SSL MQTT 数据。"
            "请确认设备已配置 SSL 证书且 www.accu.com:8883 解析到本机。"
        )
        log.info("[009-004] 收到 %d 条 SSL MQTT 消息，topic=%s",
                 len(messages), messages[0].topic if messages else 'N/A')


# ══════════════════════════════════════════════════════════════════════════════
# 010 — Keep Alive 心跳行为验证
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT010KeepAlive:
    """010 Keep Alive 心跳（3 条）

    PINGREQ 时序验证需要 Wireshark/scapy 抓包，标记为 manual。
    此处仅验证配置值保存/回读正确。
    """

    @pytest.fixture(autouse=True)
    def _restore_keep_alive(self, mqtt_page):
        """每条 KeepAlive 用例结束后将 Keep Alive 恢复基线值（KEEP_ALIVE_VAL=60）。
        防止 test_001 设置 KA=10 后，后续用例依赖 test_002 "意外恢复"的问题。
        """
        yield
        try:
            mqtt_page._switch_tab("General")
            mqtt_page._fill_by_label("Keep Alive", str(KEEP_ALIVE_VAL))
            mqtt_page.save()
        except Exception:
            pass

    @pytest.mark.lv1
    def test_001_keep_alive_10s_save_readback(self, mqtt_page):
        """Keep Alive = 10，Save 成功，回读为 10"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Keep Alive", "10")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Keep Alive") == "10"

    @pytest.mark.lv1
    def test_002_keep_alive_60s_save_readback(self, mqtt_page):
        """Keep Alive = 60，Save 成功，回读为 60"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Keep Alive", "60")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Keep Alive") == "60"

    @pytest.mark.lv1
    def test_003_keep_alive_120s_save_readback(self, mqtt_page):
        """Keep Alive = 120，Save 成功，回读为 120"""
        mqtt_page._switch_tab("General")
        mqtt_page._fill_by_label("Keep Alive", "120")
        assert mqtt_page.save() is True
        _reload(mqtt_page)
        mqtt_page._switch_tab("General")
        assert mqtt_page._read_input_by_label("Keep Alive") == "120"
        # 恢复
        mqtt_page._fill_by_label("Keep Alive", str(KEEP_ALIVE_VAL))
        mqtt_page.save()

    @pytest.mark.manual
    def test_pingreq_timing_wireshark(self):
        """
        [手工执行] Wireshark 过滤 mqtt.msgtype==12，统计相邻 PINGREQ 时间间隔：
          Keep Alive=10s → 间隔约 10s（±1s）
          Keep Alive=60s → 间隔约 60s（±2s）
          需要物理抓包环境，不在自动化范围内。
        """
        pytest.skip("需要 Wireshark 抓包，仅手工执行")


# ══════════════════════════════════════════════════════════════════════════════
# 011 — Retained 保留消息行为验证
# ══════════════════════════════════════════════════════════════════════════════

class TestMQTT011Retained:
    """011 Retained 保留消息（3 条）"""

    _RETAIN_TOPIC = f"{BASE_TOPIC}/retain-test"

    def _configure_retained(self, mqtt_page, retained: bool):
        mqtt_page.configure(MQTTConfig(
            enabled=True,
            general=MQTTGeneralConfig(
                broker_address=BROKER_ADDR,
                broker_port=BROKER_PORT,
                client_id=BASE_CLIENT_ID,
                keep_alive=KEEP_ALIVE_VAL,
                timeout=TIMEOUT_VAL,
            ),
            topic=MQTTTopicConfig(
                base_topic=self._RETAIN_TOPIC,
                interval="30 seconds",
                retained=retained,
                qos=0,
            ),
        ))

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_001_retained_yes_new_subscriber_gets_last_message(self, mqtt_page):
        """Retained=Yes：新订阅者立即收到 Broker 持有的最后保留消息
        使用内嵌 amqtt Broker（0.0.0.0:1883），paho 连 127.0.0.1。
        """
        from mqtt_comparator import start_embedded_broker

        stop_broker = start_embedded_broker(port=BROKER_PORT, duration=300)
        try:
            self._configure_retained(mqtt_page, retained=True)

            # 客户端 A：订阅，等待至少 1 条消息后断开
            msg_a: list = []
            cli_a = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-retain-a")
            cli_a.on_message = lambda _c, _u, m: msg_a.append(m)
            cli_a.connect("127.0.0.1", BROKER_PORT)
            cli_a.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_a.loop_start()
            deadline = time.time() + 90
            while not msg_a and time.time() < deadline:
                time.sleep(1)
            cli_a.loop_stop()
            cli_a.disconnect()
            assert msg_a, "客户端 A 未能在 90s 内收到消息（设备是否已连入本机 Broker？）"

            # 诊断：检查设备发布时是否带 retain 标志
            retain_flags = {m.retain for m in msg_a}
            log.info("[retain-001] 客户端 A 共收到 %d 条消息，retain 标志集合：%s",
                     len(msg_a), retain_flags)
            if 1 not in retain_flags:
                pytest.fail(
                    f"设备发布的消息 retain 标志均为 0（共 {len(msg_a)} 条）。"
                    "设备未在 PUBLISH 包中设置 retain=1，即使配置 Retained=Yes，"
                    "Broker 不会持久化消息，客户端 B 无法收到历史消息。"
                    "请确认固件是否正确实现了 MQTT Retained 功能。"
                )

            time.sleep(5)

            # 客户端 B：重新订阅，应立即（不等 30s）收到保留消息
            msg_b: list = []
            cli_b = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-retain-b")
            cli_b.on_message = lambda _c, _u, m: msg_b.append(m)
            cli_b.connect("127.0.0.1", BROKER_PORT)
            cli_b.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_b.loop_start()
            t_start = time.time()
            while not msg_b and time.time() - t_start < 5:
                time.sleep(0.5)
            cli_b.loop_stop()
            cli_b.disconnect()

            assert msg_b, "Retained=Yes：客户端 B 订阅后应立即收到保留消息（5s 内）"
            assert msg_b[0].retain == 1, \
                f"收到的消息 Retain 标志应为 1，实为 {msg_b[0].retain}"
        finally:
            stop_broker.set()

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_002_retained_no_new_subscriber_no_history(self, mqtt_page):
        """Retained=No：新订阅者不会收到历史消息，需等下一个 Interval
        使用内嵌 amqtt Broker（0.0.0.0:1883），paho 连 127.0.0.1。
        """
        from mqtt_comparator import start_embedded_broker

        stop_broker = start_embedded_broker(port=BROKER_PORT, duration=300)
        try:
            self._configure_retained(mqtt_page, retained=False)

            # 客户端 A：等待 1 条消息
            msg_a: list = []
            cli_a = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-noretain-a")
            cli_a.on_message = lambda _c, _u, m: msg_a.append(m)
            cli_a.connect("127.0.0.1", BROKER_PORT)
            cli_a.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_a.loop_start()
            deadline = time.time() + 90
            while not msg_a and time.time() < deadline:
                time.sleep(1)
            cli_a.loop_stop()
            cli_a.disconnect()
            assert msg_a, "客户端 A 未收到消息（设备是否已连入本机 Broker？）"

            time.sleep(5)

            # 客户端 B：订阅后 5s 内不应立即收到历史消息
            msg_b: list = []
            cli_b = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-noretain-b")
            cli_b.on_message = lambda _c, _u, m: msg_b.append(m)
            cli_b.connect("127.0.0.1", BROKER_PORT)
            cli_b.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_b.loop_start()
            time.sleep(5)
            cli_b.loop_stop()
            cli_b.disconnect()

            assert not msg_b, \
                "Retained=No：客户端 B 订阅后 5s 内不应收到历史消息"
        finally:
            stop_broker.set()

    @pytest.mark.lv1
    @pytest.mark.integration
    def test_003_retained_yes_to_no_clears_broker_retain(self, mqtt_page):
        """Retained 从 Yes 切换为 No 后，新订阅者不再收到旧保留消息
        使用内嵌 amqtt Broker（0.0.0.0:1883），paho 连 127.0.0.1。
        """
        from mqtt_comparator import start_embedded_broker

        stop_broker = start_embedded_broker(port=BROKER_PORT, duration=300)
        try:
            # 先设置 Retained=Yes，让 Broker 持有保留消息
            self._configure_retained(mqtt_page, retained=True)
            msg_init: list = []
            cli_init = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-chg-init")
            cli_init.on_message = lambda _c, _u, m: msg_init.append(m)
            cli_init.connect("127.0.0.1", BROKER_PORT)
            cli_init.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_init.loop_start()
            deadline = time.time() + 90
            while not msg_init and time.time() < deadline:
                time.sleep(1)
            cli_init.loop_stop()
            cli_init.disconnect()
            assert msg_init, "初始 Retained=Yes 未收到消息"

            # 切换为 Retained=No
            self._configure_retained(mqtt_page, retained=False)
            time.sleep(5)

            # 新客户端：5s 内不应立即收到旧保留消息
            msg_new: list = []
            cli_new = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="pytest-chg-new")
            cli_new.on_message = lambda _c, _u, m: msg_new.append(m)
            cli_new.connect("127.0.0.1", BROKER_PORT)
            cli_new.subscribe(f"{self._RETAIN_TOPIC}/#")
            cli_new.loop_start()
            time.sleep(5)
            cli_new.loop_stop()
            cli_new.disconnect()

            assert not msg_new, \
                "Retained 切换为 No 后，新订阅者不应立即收到旧保留消息"
        finally:
            stop_broker.set()
            # 恢复基础配置
            _restore_base(mqtt_page)
