# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_003_005
用例标题: SSL 连接时数据仍能正常上报
用例级别: LV1

预置条件:
  1. 设备已正常上电，已登录网关 Web UI
  2. Azure IoT Hub 已为设备注册 X509 证书认证方式
  3. tests/protocols/azure_iot/certs/client.pem 和 key.pem 为合法且匹配的证书/私钥
  4. config.yaml 中 primary_conn_str 和 eventhub_conn_str 均已正确填写

测试步骤:
  1. Enable Azure IoT，填写合法 Primary Connection String 和 Interval
  2. 开启 Enable SSL，上传合法匹配的 Certificate 和 Key 文件
  3. 选择设备并配置参数，保存
  4. 执行 Test Connection，验证连接成功
  5. 通过 azure_iot_verifier 验证 Azure IoT Hub 收到设备上报数据

预期结果:
  1. Certificate 和 Key 文件上传成功
  2. SSL 模式下 Test Connection 返回连接成功
  3. Azure IoT Hub Event Hub 端能正常收到设备数据，上报值正确
"""

from pathlib import Path

import pytest

from utils.azure_iot_verifier import run_verifier

_THIS_DIR = Path(__file__).resolve().parent
_CERTS_DIR = _THIS_DIR / "certs"
_CERT_FILE = str(_CERTS_DIR / "client.pem")
_KEY_FILE  = str(_CERTS_DIR / "key.pem")

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


class TestCase_AcuHMI_1_7_AZR_003_005:

    @pytest.mark.azure_iot
    @pytest.mark.slow
    def test_ssl_connection_and_data_upload(self, azure_page, azure_cfg):
        """开启 SSL + 合法证书，Test Connection 成功且 Azure IoT Hub 收到数据"""
        if not Path(_CERT_FILE).exists() or not Path(_KEY_FILE).exists():
            pytest.skip(
                f"SSL 证书文件不存在（{_CERT_FILE}），请先放置合法的 client.pem + key.pem"
            )

        azure = azure_cfg["azure_iot"]
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure["primary_conn_str"])
        azure_page.set_interval(azure.get("interval", "30 seconds"))
        azure_page.select_only_device("")
        azure_page.configure_all_devices_parameters(checked_only=True)

        _enable_ssl(azure_page.page)

        cert_inputs = azure_page.page.locator(
            "xpath=//div[contains(normalize-space(.),'Certificate')]//input[@type='file']"
        ).all()
        key_inputs = azure_page.page.locator(
            "xpath=//div[contains(normalize-space(.),'Key')]//input[@type='file']"
        ).all()

        if not cert_inputs or not key_inputs:
            pytest.skip("未找到 Certificate/Key 上传控件，请确认选择器与实际 UI 一致")

        cert_inputs[0].set_input_files(_CERT_FILE)
        azure_page.page.wait_for_timeout(500)
        key_inputs[0].set_input_files(_KEY_FILE)
        azure_page.page.wait_for_timeout(500)

        azure_page.save()

        result = azure_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"SSL 模式下 Test Connection 应成功，实际：{result!r}"

        ok = run_verifier(azure_cfg, timeout=120, skip_web=True)
        assert ok, "SSL 连接成功后 Azure IoT Hub 应能正常收到设备上报数据"
