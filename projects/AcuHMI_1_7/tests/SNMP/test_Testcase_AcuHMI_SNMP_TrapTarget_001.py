"""
Testcase_AcuHMI_SNMP_TrapTarget_001: FTS_case25: 非法 Trap Target（192.168.2.a / 255.255.255.256 / 192.168.2.）保存失败
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3, check_form_error  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_TrapTarget_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_TrapTarget_001"""

    def test_Testcase_AcuHMI_SNMP_TrapTarget_001(self, snmp_page):
        """FTS_case25: 非法 Trap Target（192.168.2.a / 255.255.255.256 / 192.168.2.）保存失败"""
        page = snmp_page

        invalid_ips = ["192.168.2.a", "255.255.255.256", "192.168.2."]
        for ip in invalid_ips:
            self._reload(page)
            step(f"case25: Trap Target 1={ip!r} → 期望保存失败")

            # 开启 Trap（第二个 radio 组的 Enable）
            page.locator(".el-radio").filter(has_text="Enable").nth(1).click()
            page.wait_for_timeout(500)

            t1 = page.locator('input[placeholder="Enter Trap Target 1"]')
            t1.click()
            t1.fill(ip)

            page.click('button:has-text("Save")')
            page.wait_for_timeout(2000)

            has_error = check_form_error(page)
            assert has_error, f"case25: Trap Target={ip!r} 应保存失败但未出现校验错误"
            step(f"  ✓ Trap Target={ip!r} 保存失败（符合预期）")

        # 恢复正常配置
        self._reload(page)
        apply_snmp_v2c(page, self._base_v2c(port="161"),
                       selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        step("case25: 已恢复正常配置")

