"""
Testcase_AcuHMI_SNMP_v2c_Port_001: FTS_case01: v2c port=161 仅 AcuRev1300 下载 MIB 验证数据
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_v2c_Port_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_v2c_Port_001"""

    def test_Testcase_AcuHMI_SNMP_v2c_Port_001(self, snmp_page):
        """FTS_case01: v2c port=161 仅 AcuRev1300 下载 MIB 验证数据"""
        page = snmp_page
        self._reload(page)

        step("case01: 配置 v2c port=161 AcuRev1300")
        ok = apply_snmp_v2c(page, self._base_v2c(port="161"),
                             selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        assert ok, "case01: 参数保存失败"

        self._download_mib_once(page)

        step("case01: SnmpWalk v2c port=161")
        data = self._walk_v2c(port=161)
        assert len(data) > 0, f"case01: SnmpWalk 无数据（port=161 community={VALID_COMMUNITY}）"
        step(f"  ✓ 返回 {len(data)} OID")

