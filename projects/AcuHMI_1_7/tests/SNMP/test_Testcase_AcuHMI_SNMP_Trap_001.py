"""
Testcase_AcuHMI_SNMP_Trap_001: FTS_case17: AcuRev4100/AcuIOM 掉电告警上报 Trap
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Trap_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_Trap_001"""

    @pytest.mark.skip(reason="需要 AcuIOM 设备掉电触发告警，以及 Trap 监听服务")
    def test_Testcase_AcuHMI_SNMP_Trap_001(self, snmp_page):
        """FTS_case17: AcuRev4100/AcuIOM 掉电告警上报 Trap"""
        pass

