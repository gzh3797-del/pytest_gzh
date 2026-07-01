"""
Testcase_AcuHMI_SNMP_Trap_003: FTS_case19: AcuRev4100 设备告警 Trap 上报
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Trap_003(SNMPBase):
    """Testcase_AcuHMI_SNMP_Trap_003"""

    @pytest.mark.skip(reason="需要触发 AcuRev4100 设备告警，以及 Trap 监听服务")
    def test_Testcase_AcuHMI_SNMP_Trap_003(self, snmp_page):
        """FTS_case19: AcuRev4100 设备告警 Trap 上报"""
        pass

