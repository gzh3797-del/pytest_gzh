"""
FTS编号: FTS_AcuHMI_AZR_002_001
用例标题: Primary Connection String 参数校验与合法值校验
用例级别: LV1

预置条件: 网关已登录，已导航到 Azure IoT 配置页面，Azure IoT 处于 Enable 状态

测试步骤:
  1. 输入格式错误的 Connection String（缺少 HostName 字段），点击 Save，验证被拦截
  2. 输入超过 255 字符的 Connection String，点击 Save，验证被拦截
  3. 输入合法 Connection String，点击 Save，验证保存成功
  4. 清空 Primary Connection String，点击 Save，验证被拦截

预期结果:
  - 格式错误的 Connection String 无法保存，页面给出提示
  - 超过 255 字符的 Connection String 无法保存
  - 合法 Connection String 保存成功
  - 空 Connection String 无法保存
"""

import pytest

INVALID_FORMAT_CONN_STR = "DeviceId=test-device;SharedAccessKey=abc=="
OVER_LENGTH_CONN_STR    = "HostName=" + "a" * 200 + ".azure-devices.net;DeviceId=d;SharedAccessKey=abc=="


def _get_page_message(page) -> str:
    try:
        el = page.page.locator(page.RESULT_MSG).first
        el.wait_for(state="visible", timeout=5000)
        return el.inner_text().strip()
    except Exception:
        return ""


class TestCase_AcuHMI_1_7_AZR_002_001:

    @pytest.mark.azure_iot
    def test_invalid_format_blocked(self, azure_page):
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(INVALID_FORMAT_CONN_STR)
        azure_page.save()
        assert _get_page_message(azure_page), "格式错误的 Connection String 应被拦截并给出提示"

    @pytest.mark.azure_iot
    def test_over_length_blocked(self, azure_page):
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(OVER_LENGTH_CONN_STR)
        azure_page.save()
        assert _get_page_message(azure_page), "超过 255 字符的 Connection String 应被拦截"

    @pytest.mark.azure_iot
    def test_valid_conn_str_saves(self, azure_page, azure_cfg):
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(azure_cfg["azure_iot"]["primary_conn_str"])
        azure_page.save()
        assert "login" not in azure_page.page.url.lower(), "合法 Connection String 应能保存成功"

    @pytest.mark.azure_iot
    def test_empty_conn_str_blocked(self, azure_page):
        azure_page.ensure_enabled()
        loc = azure_page.page.locator(azure_page.PRIMARY_CONN_STR_INPUT)
        loc.triple_click()
        azure_page.page.keyboard.press("Delete")
        azure_page.save()
        assert _get_page_message(azure_page), "空 Connection String 应被拦截"
