"""
Testcase_AcuHMI_SNMP_v3_Auth_006: FTS_case16: 设备 v3 SHA/DES 正确配置，NMS 使用错误凭据 → GET 超时
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import (SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3,  # noqa: F401
                        V3_USERNAME, snmp_walk_device_v3)


class Test_Testcase_AcuHMI_SNMP_v3_Auth_006(SNMPBase):
    """Testcase_AcuHMI_SNMP_v3_Auth_006"""

    def test_Testcase_AcuHMI_SNMP_v3_Auth_006(self, snmp_page):
        """FTS_case16: 设备 v3 SHA/DES 正确配置，NMS 使用错误凭据 → GET 超时"""
        page = snmp_page
        self._reload(page)

        step("case16: 设备配置 v3 SHA/DES port=16100（正确凭据）")
        ok = apply_snmp_v3(page, self._base_v3(port="16100", auth="SHA", priv="DES"),
                            selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case16: 设备配置保存失败"

        step("case16: 等待 SNMP agent 重启 10s...")
        time.sleep(10)
        step("case16: SnmpWalk 使用错误密码，期望无数据（30s）")
        data = snmp_walk_device_v3(
            port=16100,
            security_name=V3_USERNAME,
            auth_protocol="SHA",
            auth_password="wrongpassword1",
            priv_protocol="DES",
            priv_password="wrongprivpass1",
            security_level="authPriv",
            total_timeout=30,
        )
        assert len(data) == 0, \
            f"case16: 凭据不匹配时不应有数据，实际返回 {len(data)} OID"
        step("  ✓ 凭据不匹配，无数据（符合预期）")

