"""
Testcase_AcuHMI_SNMP_v3_Auth_002: FTS_case12: v3 MD5/DES port=16100 AcuRev1300
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import (SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3,  # noqa: F401
                        V3_PASSWORD, V3_PRIV_PASSWORD)


class Test_Testcase_AcuHMI_SNMP_v3_Auth_002(SNMPBase):
    """Testcase_AcuHMI_SNMP_v3_Auth_002"""

    def test_Testcase_AcuHMI_SNMP_v3_Auth_002(self, snmp_page):
        """FTS_case12: v3 MD5/DES port=16100 AcuRev1300"""
        page = snmp_page
        self._reload(page)

        step("case12: 配置 v3 MD5/DES port=16100")
        ok = apply_snmp_v3(page, self._base_v3(port="16100", auth="MD5", priv="DES"),
                            selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case12: v3 配置保存失败"
        self._download_mib_once(page)

        step("case12: SnmpWalk v3 MD5/DES port=16100")
        data = self._walk_v3(port=16100, auth_protocol="MD5", password=V3_PASSWORD,
                              priv_protocol="DES", priv_password=V3_PRIV_PASSWORD)
        assert len(data) > 0, "case12: v3 MD5/DES SnmpWalk 无数据"
        step(f"  ✓ v3 MD5/DES 返回 {len(data)} OID")

