"""
Testcase_AcuHMI_SNMP_v3_Auth_001: FTS_case11: v3 MD5/NONE PRIV port=161 AcuRev1300
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import (SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3,  # noqa: F401
                        V3_PASSWORD)


class Test_Testcase_AcuHMI_SNMP_v3_Auth_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_v3_Auth_001"""

    def test_Testcase_AcuHMI_SNMP_v3_Auth_001(self, snmp_page):
        """FTS_case11: v3 MD5/NONE PRIV port=161 AcuRev1300"""
        page = snmp_page
        self._reload(page)

        step("case11: 配置 v3 MD5/NONE PRIV port=161")
        ok = apply_snmp_v3(page, self._base_v3(port="161", auth="MD5", priv="NONE PRIV"),
                            selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case11: v3 配置保存失败"
        self._download_mib_once(page)

        step("case11: SnmpWalk v3 MD5/authNoPriv port=161")
        data = self._walk_v3(port=161, auth_protocol="MD5", password=V3_PASSWORD,
                              priv_protocol="NONE PRIV")
        assert len(data) > 0, "case11: v3 MD5/NONE PRIV SnmpWalk 无数据"
        step(f"  ✓ v3 MD5/NONE PRIV 返回 {len(data)} OID")

