"""
Testcase_AcuHMI_SNMP_PortMismatch_001: FTS_case07: NMS 端口与设备不一致，GET 超时
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_PortMismatch_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_PortMismatch_001"""

    def test_Testcase_AcuHMI_SNMP_PortMismatch_001(self, snmp_page):
        """FTS_case07: NMS 端口与设备不一致，GET 超时"""
        page = snmp_page
        self._reload(page)

        step("case07: 设备 port=161，NMS 使用 port=163")
        ok = apply_snmp_v2c(page, self._base_v2c(port="161"),
                             selected_devices=["AcuRev4100"])
        assert ok, "case07: 配置保存失败"

        step("case07: SnmpWalk port=163（错误端口），期望超时/无数据（30s）")
        data = self._walk_v2c(port=163, timeout=30)
        assert len(data) == 0, \
            f"case07: 端口不匹配时不应有数据，实际 {len(data)} OID"
        step("  ✓ SnmpWalk 无数据（符合预期）")

