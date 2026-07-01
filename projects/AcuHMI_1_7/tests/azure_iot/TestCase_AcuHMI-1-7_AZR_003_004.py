# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_003_004
用例标题: Certificate 与 Key 不匹配时连接失败
用例级别: LV1

预置条件:
  1. 设备已正常上电，已登录网关 Web UI
  2. Azure IoT 处于 Enable 状态，Enable SSL 已开启
  3. tests/protocols/azure_iot/certs/ 下存在两套不同密钥对的证书文件

测试步骤:
  1. 上传不匹配的 Certificate 和 Key 文件（来自不同密钥对）
  2. 保存并执行 Test Connection

预期结果:
  1. 连接失败，页面显示证书错误提示
  2. Azure IoT Hub 不接收到任何数据
"""

from pathlib import Path

import pytest

_THIS_DIR = Path(__file__).resolve().parent
_CERTS_DIR = _THIS_DIR / "certs"

# 两套不同密钥对的证书/私钥
_CERT_A = str(_CERTS_DIR / "29e287f92d07dc357838878febdac0cfc1d046b2abf81c715d2f42be4a2477c6-certificate.pem.crt")
_KEY_B  = str(_CERTS_DIR / "a01fc88bfb456f3b9e5e8edeed55882ed5fffaa196e25199b0bd36fe39137530-private.pem.key")

_SSL_TOGGLE_SELECTORS = [
    "xpath=//div[contains(@class,'el-form-item')][.//*[contains(normalize-space(.),'Enable SSL')]]"
    "//*[contains(@class,'el-switch')]",
    "xpath=//*[contains(normalize-space(.),'Enable SSL')]/following::*[contains(@class,'el-switch')][1]",
]


def _enable_ssl(page) -> None:
    for sel in _SSL_TOGGLE_SELECTORS:
        els = page.locator(sel).all()
        if not els:
            continue
        toggle = els[0]
        cls = toggle.get_attribute("class") or ""
        if "is-checked" not in cls:
            toggle.evaluate("el => el.click()")
            page.wait_for_timeout(800)
        return
    raise RuntimeError("未找到 Enable SSL 开关")


def _connection_failed(result: str) -> bool:
    return any(kw in result.lower() for kw in (
        "fail", "error", "invalid", "连接失败", "未检测到", "certificate", "cert"
    ))


class TestCase_AcuHMI_1_7_AZR_003_004:

    @pytest.mark.azure_iot
    def test_mismatched_cert_and_key_fails(self, azure_page, azure_cfg):
        """Certificate 与 Key 来自不同密钥对，Test Connection 应失败"""
        if not Path(_CERT_A).exists() or not Path(_KEY_B).exists():
            pytest.skip("所需的不匹配证书文件不存在，跳过本用例")

        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")

        _enable_ssl(azure_page.page)

        cert_inputs = azure_page.page.locator(
            "xpath=//div[contains(normalize-space(.),'Certificate')]//input[@type='file']"
        ).all()
        key_inputs = azure_page.page.locator(
            "xpath=//div[contains(normalize-space(.),'Key')]//input[@type='file']"
        ).all()

        if not cert_inputs or not key_inputs:
            pytest.skip("未找到 Certificate/Key 上传控件，请确认选择器与实际 UI 一致")

        # 上传不匹配的证书 A + 私钥 B
        cert_inputs[0].set_input_files(_CERT_A)
        azure_page.page.wait_for_timeout(500)
        key_inputs[0].set_input_files(_KEY_B)
        azure_page.page.wait_for_timeout(500)

        azure_page.save()
        result = azure_page.test_connection()
        assert _connection_failed(result), (
            f"Certificate 与 Key 不匹配时 Test Connection 应失败，实际：{result!r}"
        )
