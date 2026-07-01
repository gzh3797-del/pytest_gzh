"""
Testcase_AcuHMI_SNMP_Enable_001: FTS_case10: 关闭 SNMP → GET 超时；重新开启 → GET 成功
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import (SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3,  # noqa: F401
                        VALID_COMMUNITY)


class Test_Testcase_AcuHMI_SNMP_Enable_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_Enable_001"""

    def test_Testcase_AcuHMI_SNMP_Enable_001(self, snmp_page):
        """FTS_case10: 关闭 SNMP → GET 超时；重新开启 → GET 成功"""
        page = snmp_page
        self._reload(page)

        step("case10: 配置 v2c port=161 并确认 SnmpWalk 正常")
        ok = apply_snmp_v2c(page, self._base_v2c(port="161"),
                             selected_devices=["AcuRev4100"])
        assert ok, "case10: 初始配置保存失败"
        data1 = self._walk_v2c(port=161)
        assert len(data1) > 0, "case10: 初始 SnmpWalk 无数据"
        step(f"  ✓ 初始 SnmpWalk 正常: {len(data1)} OID")

        step("case10: 关闭 SNMP")
        self._reload(page)
        ok2 = apply_snmp_v2c(page, {
            "enable": False,
            "port": "161",
            "ro_community": VALID_COMMUNITY,
            "trap_enable": False,
            "buffer_size": "30",
            "hold_time": "0",
        }, selected_devices=["AcuRev4100"])
        assert ok2, "case10: 关闭 SNMP 保存失败"

        step("case10: SNMP 已关闭，SnmpWalk 期望无数据（30s）")
        data2 = self._walk_v2c(port=161, timeout=30)
        assert len(data2) == 0, \
            f"case10: SNMP 关闭后不应有数据，实际 {len(data2)} OID"
        step("  ✓ SNMP 关闭后无数据")

        step("case10: 重新开启 SNMP")
        self._reload(page)
        ok3 = apply_snmp_v2c(page, self._base_v2c(port="161"),
                              selected_devices=["AcuRev4100"])
        assert ok3, "case10: 重新开启 SNMP 保存失败"
        data3 = self._walk_v2c(port=161)
        assert len(data3) > 0, "case10: SNMP 重开后 SnmpWalk 无数据"
        step(f"  ✓ SNMP 重开后正常: {len(data3)} OID")

