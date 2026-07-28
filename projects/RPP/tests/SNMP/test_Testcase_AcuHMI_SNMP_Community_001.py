"""
Testcase_AcuHMI_SNMP_Community_001: FTS_case05: 空 Community 配置后 NMS GET 超时（或 UI 拒绝保存）
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Community_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_Community_001"""

    def test_Testcase_AcuHMI_SNMP_Community_001(self, snmp_page):
        """FTS_case05: 空 Community 配置后 NMS GET 超时（或 UI 拒绝保存）"""
        page = snmp_page
        self._reload(page)

        step("case05: 配置空 Community")
        ok = apply_snmp_v2c(page, self._base_v2c(port="16199", community=""),
                             selected_devices=["AcuvimIIR"])

        if not ok:
            step("  ✓ UI 拒绝空 Community（表单校验错误，NMS 无法认证）")
            return

        step("case05: 空 Community 已保存，SnmpWalk 期望超时/无数据（30s）")
        data = self._walk_v2c(port=16199, community=VALID_COMMUNITY, timeout=30)
        assert len(data) == 0, \
            f"case05: 空 Community 配置下不应有 SNMP 数据，实际返回 {len(data)} OID"
        step("  ✓ SnmpWalk 无数据（符合预期）")

