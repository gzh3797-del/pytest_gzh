"""
Testcase_AcuHMI_SNMP_Trap_002: FTS_case18: AcuIOM 通信异常 Trap 上报
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Trap_002(SNMPBase):
    """Testcase_AcuHMI_SNMP_Trap_002"""

    @pytest.mark.skip(reason="需要 AcuIOM 设备通信异常，以及 Trap 监听服务")
    def test_Testcase_AcuHMI_SNMP_Trap_002(self, snmp_page):
        """FTS_case18: AcuIOM 通信异常 Trap 上报"""
        pass

