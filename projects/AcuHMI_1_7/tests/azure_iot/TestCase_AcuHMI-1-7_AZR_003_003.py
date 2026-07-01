# -*- coding: utf-8 -*-
"""
用例编号: TestCase_AcuHMI-1-7_AZR_003_003
用例标题: 格式非法的证书/私钥文件禁止上传
用例级别: LV1

预置条件:
  1. 设备已正常上电，已登录网关 Web UI
  2. Azure IoT 处于 Enable 状态，Enable SSL 已开启

测试步骤:
  1. 上传 .txt 格式的 Certificate 文件，验证被拒绝并提示格式错误
  2. 上传损坏（非 PEM）的 Key 文件，验证被拒绝并提示格式错误

预期结果:
  1. 禁止上传，显示证书文件格式错误提示
  2. 禁止上传，显示私钥文件格式错误提示
"""

import os
import tempfile
from pathlib import Path

import pytest

_THIS_DIR = Path(__file__).resolve().parent

_SSL_TOGGLE_SELECTORS = [
    "xpath=//label[contains(@class,'el-switch') and ./following-sibling::*[contains(.,'Enable SSL')]]",
    "xpath=//*[contains(normalize-space(.),'Enable SSL')]/following::*[contains(@class,'el-switch')][1]",
    "xpath=//div[contains(@class,'el-form-item')][.//*[contains(normalize-space(.),'Enable SSL')]]"
    "//*[contains(@class,'el-switch')]",
]


def _enable_ssl(page) -> None:
    """打开 Enable SSL 开关（若尚未打开）。"""
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
    raise RuntimeError("未找到 Enable SSL 开关，请确认选择器与实际页面匹配")


def _get_upload_error(page) -> str:
    """读取上传或保存后的错误提示文本，无则返回空串。"""
    page.wait_for_timeout(1500)
    for xpath in [
        "xpath=//div[contains(@class,'el-message--error')]",
        "xpath=//div[contains(@class,'el-notification--error')]",
        "xpath=//*[contains(@class,'is-error')]//div[contains(@class,'el-form-item__error')]",
        "xpath=//*[contains(normalize-space(.),'格式') or contains(normalize-space(.),'format')"
        " or contains(normalize-space(.),'invalid') or contains(normalize-space(.),'非法')]"
        "[contains(@class,'error') or contains(@class,'warning') or contains(@class,'message')]",
    ]:
        for el in page.locator(xpath).all():
            txt = (el.text_content() or "").strip()
            if txt:
                return txt
    return ""


class TestCase_AcuHMI_1_7_AZR_003_003:

    @pytest.mark.azure_iot
    def test_invalid_cert_format_blocked(self, azure_page):
        """上传 .txt 格式的 Certificate 文件，应被拒绝"""
        azure_page.ensure_enabled()
        _enable_ssl(azure_page.page)

        with tempfile.NamedTemporaryFile(
            suffix=".txt", delete=False, mode="w", encoding="utf-8"
        ) as f:
            f.write("This is not a valid PEM certificate\n")
            tmp_cert = f.name

        try:
            cert_inputs = azure_page.page.locator(
                "xpath=//div[contains(normalize-space(.),'Certificate')]//input[@type='file']"
            ).all()
            if not cert_inputs:
                pytest.skip("未找到 Certificate 文件上传控件，请确认选择器与实际 UI 一致")

            cert_inputs[0].set_input_files(tmp_cert)
            azure_page.page.wait_for_timeout(1500)

            err_inline = _get_upload_error(azure_page.page)
            azure_page.save()
            err_save = _get_upload_error(azure_page.page)

            assert err_inline or err_save, (
                "上传 .txt 格式 Certificate 文件后，页面应显示格式错误提示或阻止保存"
            )
        finally:
            os.unlink(tmp_cert)

    @pytest.mark.azure_iot
    def test_corrupted_key_format_blocked(self, azure_page):
        """上传损坏/非 PEM 格式的 Key 文件，应被拒绝"""
        azure_page.ensure_enabled()
        _enable_ssl(azure_page.page)

        with tempfile.NamedTemporaryFile(
            suffix=".key", delete=False, mode="w", encoding="utf-8"
        ) as f:
            f.write("CORRUPTED KEY DATA !!! NOT A VALID PEM\n")
            tmp_key = f.name

        try:
            key_inputs = azure_page.page.locator(
                "xpath=//div[contains(normalize-space(.),'Key')]//input[@type='file']"
            ).all()
            if not key_inputs:
                pytest.skip("未找到 Key 文件上传控件，请确认选择器与实际 UI 一致")

            key_inputs[0].set_input_files(tmp_key)
            azure_page.page.wait_for_timeout(1500)

            err_inline = _get_upload_error(azure_page.page)
            azure_page.save()
            err_save = _get_upload_error(azure_page.page)

            assert err_inline or err_save, (
                "上传损坏的 Key 文件后，页面应显示格式错误提示或阻止保存"
            )
        finally:
            os.unlink(tmp_key)
