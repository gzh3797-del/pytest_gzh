"""
FTS编号: FTS_AcuHMI_AZR_003_002
用例标题: 错误 Connection String 连接失败
用例级别: LV2

预置条件:
  - 网关已登录，已导航到 Azure IoT 配置页面

测试步骤:
  Case A: 合法格式但错误 SharedAccessKey（Base64 乱码）
  Case B: HostName 格式错误（非 .azure-devices.net）
  Case C: DeviceId 为空（Connection String 中缺少 DeviceId）
  Case D: Primary Connection String 内容为全随机乱码

预期结果:
  - 以上 4 种情况 Test Connection 均失败
  - 页面无崩溃，可重新操作
"""

import pytest

_WRONG_KEY_CONN_STR        = "HostName=test.azure-devices.net;DeviceId=dev01;SharedAccessKey=WRONGKEY_ABCDEFGHIJK=="
_BAD_HOSTNAME_CONN_STR     = "HostName=invalid-host.example.com;DeviceId=dev01;SharedAccessKey=abc123=="
_MISSING_DEVICEID_CONN_STR = "HostName=test.azure-devices.net;SharedAccessKey=abc123=="
_GIBBERISH_CONN_STR        = "this_is_not_a_connection_string_@#$%"


def _connection_failed(result: str) -> bool:
    return any(kw in result.lower() for kw in ("fail", "error", "invalid", "连接失败", "未检测到"))


class TestCase_AcuHMI_1_7_AZR_003_002:

    @pytest.mark.azure_iot
    def test_wrong_shared_access_key(self, azure_page):
        """合法格式但 SharedAccessKey 错误 → 连接应失败"""
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(_WRONG_KEY_CONN_STR)
        azure_page.select_only_device("")
        azure_page.save()
        assert _connection_failed(azure_page.test_connection()), \
            "错误 SharedAccessKey 应连接失败"

    @pytest.mark.azure_iot
    def test_bad_hostname(self, azure_page):
        """非 azure-devices.net 的 HostName → 连接应失败"""
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(_BAD_HOSTNAME_CONN_STR)
        azure_page.select_only_device("")
        azure_page.save()
        assert _connection_failed(azure_page.test_connection()), \
            "错误 HostName 应连接失败"

    @pytest.mark.azure_iot
    def test_missing_device_id(self, azure_page):
        """缺少 DeviceId 字段 → 连接应失败"""
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(_MISSING_DEVICEID_CONN_STR)
        azure_page.select_only_device("")
        azure_page.save()
        assert _connection_failed(azure_page.test_connection()), \
            "缺少 DeviceId 的 Connection String 应连接失败"

    @pytest.mark.azure_iot
    def test_gibberish_conn_str(self, azure_page):
        """完全乱码的 Connection String → 连接应失败"""
        azure_page.ensure_enabled()
        azure_page.set_primary_conn_str(_GIBBERISH_CONN_STR)
        azure_page.select_only_device("")
        azure_page.save()
        assert _connection_failed(azure_page.test_connection()), \
            "乱码 Connection String 应连接失败"
