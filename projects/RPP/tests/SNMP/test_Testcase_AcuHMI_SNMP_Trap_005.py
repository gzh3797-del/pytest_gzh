"""
Testcase_AcuHMI_SNMP_Trap_005: FTS_case21: Report Buffer Size=30 时 Trap 消息缓存行为验证
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Trap_005(SNMPBase):
    """Testcase_AcuHMI_SNMP_Trap_005"""

    @pytest.mark.skip(reason="需要触发多条告警且需要 Trap 监听服务，含时序验证")
    def test_Testcase_AcuHMI_SNMP_Trap_005(self, snmp_page):
        """FTS_case21: Report Buffer Size=30 时 Trap 消息缓存行为验证"""
        pass

