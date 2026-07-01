"""
Testcase_AcuHMI_SNMP_BufferSize_001: FTS_case22: Report Buffer Size 边界：0/30 保存成功，31 保存失败
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_BufferSize_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_BufferSize_001"""

    def test_Testcase_AcuHMI_SNMP_BufferSize_001(self, snmp_page):
        """FTS_case22: Report Buffer Size 边界：0/30 保存成功，31 保存失败"""
        page = snmp_page
        base = self._base_v2c(port="161", trap_enable=True,
                               trap_target="192.168.2.9")

        for buf, expect_ok in [("0", True), ("30", True), ("31", False)]:
            self._reload(page)
            step(f"case22: Report Buffer Size={buf} → 期望{'成功' if expect_ok else '失败'}")
            cfg = {**base, "buffer_size": buf}
            ok = apply_snmp_v2c(page, cfg, selected_devices=["AcuRev4100"])
            if expect_ok:
                assert ok, f"case22: Buffer Size={buf} 应保存成功"
                step(f"  ✓ Buffer Size={buf} 保存成功")
            else:
                assert not ok, f"case22: Buffer Size={buf} 应保存失败但成功了"
                step(f"  ✓ Buffer Size={buf} 保存失败（符合预期）")

