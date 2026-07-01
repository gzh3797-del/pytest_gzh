"""
Testcase_AcuHMI_SNMP_PortBoundary_001: FTS_case24: 非法端口（160/16200/-161/A/@）保存失败
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page, step, apply_snmp_v2c, apply_snmp_v3, check_form_error  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_PortBoundary_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_PortBoundary_001"""

    def test_Testcase_AcuHMI_SNMP_PortBoundary_001(self, snmp_page):
        """FTS_case24: 非法端口（160/16200/-161/A/@）保存失败"""
        page = snmp_page

        for invalid_port in ["160", "16200", "-161", "A", "@"]:
            self._reload(page)
            step(f"case24: Port={invalid_port!r} → 期望保存失败")

            port_f = page.locator('input[placeholder="Enter Port"]')
            port_f.click()
            port_f.fill(invalid_port)

            page.click('button:has-text("Save")')
            page.wait_for_timeout(2000)

            has_error = check_form_error(page)
            assert has_error, f"case24: Port={invalid_port!r} 应保存失败但未出现校验错误"
            step(f"  ✓ Port={invalid_port!r} 保存失败（符合预期）")

        # 恢复合法端口
        self._reload(page)
        apply_snmp_v2c(page, self._base_v2c(port="161"),
                       selected_devices=["AcuRev1300"], fallback_devices=["AcuvimIIW"])
        step("case24: 已恢复合法端口 161")

