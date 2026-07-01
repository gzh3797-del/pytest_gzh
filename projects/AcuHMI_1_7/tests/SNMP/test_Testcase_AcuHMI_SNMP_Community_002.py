"""
Testcase_AcuHMI_SNMP_Community_002: FTS_case06: 设备 Community 与 NMS 不一致，GET 超时
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Community_002(SNMPBase):
    """Testcase_AcuHMI_SNMP_Community_002"""

    def test_Testcase_AcuHMI_SNMP_Community_002(self, snmp_page):
        """FTS_case06: 设备 Community 与 NMS 不一致，GET 超时"""
        page = snmp_page
        self._reload(page)

        DEVICE_COMM = "devicecomm123456"
        NMS_COMM    = "nmscommunnity1234"

        step(f"case06: 设备 Community={DEVICE_COMM}，NMS 使用不同 Community")
        ok = apply_snmp_v2c(page, self._base_v2c(port="161", community=DEVICE_COMM),
                             selected_devices=None)
        assert ok, "case06: 配置保存失败"

        step("case06: SnmpWalk 使用不匹配 Community，期望超时/无数据（30s）")
        data = self._walk_v2c(port=161, community=NMS_COMM, timeout=30)
        assert len(data) == 0, \
            f"case06: Community 不匹配时不应有数据，实际 {len(data)} OID"
        step("  ✓ SnmpWalk 无数据（符合预期）")

        # 恢复正确 Community，供后续用例使用
        self._reload(page)
        apply_snmp_v2c(page, self._base_v2c(port="161"), selected_devices=None)
        step("case06: 已恢复正确 Community")

