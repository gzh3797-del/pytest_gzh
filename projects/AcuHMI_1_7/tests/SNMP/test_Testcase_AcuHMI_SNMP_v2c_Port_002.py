"""
Testcase_AcuHMI_SNMP_v2c_Port_002: FTS_case02: v2c port=16100 仅 AcuRev1300 验证数据
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_v2c_Port_002(SNMPBase):
    """Testcase_AcuHMI_SNMP_v2c_Port_002"""

    def test_Testcase_AcuHMI_SNMP_v2c_Port_002(self, snmp_page):
        """FTS_case02: v2c port=16100 仅 AcuRev1300 验证数据"""
        page = snmp_page
        self._reload(page)

        step("case02: 配置 v2c port=16100 AcuRev1300")
        ok = apply_snmp_v2c(page, self._base_v2c(port="16100"),
                             selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case02: 参数保存失败"

        self._download_mib_once(page)

        step("case02: SnmpWalk v2c port=16100")
        data = self._walk_v2c(port=16100)
        assert len(data) > 0, "case02: SnmpWalk 无数据"
        step(f"  ✓ 返回 {len(data)} OID")
