"""
Testcase_AcuHMI_SNMP_Trap_004: FTS_case20: 关闭 Trap 后不上报告警消息
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Trap_004(SNMPBase):
    """Testcase_AcuHMI_SNMP_Trap_004"""

    @pytest.mark.skip(reason="需要 Trap 监听服务")
    def test_Testcase_AcuHMI_SNMP_Trap_004(self, snmp_page):
        """FTS_case20: 关闭 Trap 后不上报告警消息"""
        pass

