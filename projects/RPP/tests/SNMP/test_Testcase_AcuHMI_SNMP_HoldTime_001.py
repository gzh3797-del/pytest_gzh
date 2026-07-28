"""
Testcase_AcuHMI_SNMP_HoldTime_001: FTS_case23: Report Hold Time 边界：0/300 保存成功，301 保存失败
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_HoldTime_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_HoldTime_001"""

    def test_Testcase_AcuHMI_SNMP_HoldTime_001(self, snmp_page):
        """FTS_case23: Report Hold Time 边界：0/300 保存成功，301 保存失败"""
        page = snmp_page
        base = self._base_v2c(port="161", trap_enable=True,
                               trap_target="192.168.2.9", buf="30")

        for hold, expect_ok in [("0", True), ("300", True), ("301", False)]:
            self._reload(page)
            step(f"case23: Report Hold Time={hold} → 期望{'成功' if expect_ok else '失败'}")
            cfg = {**base, "hold_time": hold}
            ok = apply_snmp_v2c(page, cfg, selected_devices=["AcuRev4100"])
            if expect_ok:
                assert ok, f"case23: Hold Time={hold} 应保存成功"
                step(f"  ✓ Hold Time={hold} 保存成功")
            else:
                assert not ok, f"case23: Hold Time={hold} 应保存失败但成功了"
                step(f"  ✓ Hold Time={hold} 保存失败（符合预期）")

