"""
Testcase_AcuHMI_SNMP_v2c_Port_003: FTS_case03: v2c port=16159 AcuRev1300 配置保存 + SnmpWalk 验证有数据
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_v2c_Port_003(SNMPBase):
    """Testcase_AcuHMI_SNMP_v2c_Port_003"""

    def test_Testcase_AcuHMI_SNMP_v2c_Port_003(self, snmp_page):
        """FTS_case03: v2c port=16159 AcuRev1300 配置保存 + SnmpWalk 验证有数据"""
        page = snmp_page
        self._reload(page)

        step("case03: 配置 v2c port=16159 AcuRev1300")
        ok = apply_snmp_v2c(page, self._base_v2c(port="16159"),
                             selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case03: 配置保存失败"
        self._download_mib_once(page)

        step("case03: SnmpWalk v2c port=16159")
        data = self._walk_v2c(port=16159)
        assert len(data) > 0, "case03: SnmpWalk 无数据"
        step(f"  ✓ 返回 {len(data)} OID")
