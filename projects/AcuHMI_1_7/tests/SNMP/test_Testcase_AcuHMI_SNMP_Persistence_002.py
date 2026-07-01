"""
Testcase_AcuHMI_SNMP_Persistence_002: FTS_case09: 网络异常恢复后 NMS 重新请求成功
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Persistence_002(SNMPBase):
    """Testcase_AcuHMI_SNMP_Persistence_002"""

    @pytest.mark.skip(reason="需要手动操作：断开/恢复网络连接")
    def test_Testcase_AcuHMI_SNMP_Persistence_002(self, snmp_page):
        """FTS_case09: 网络异常恢复后 NMS 重新请求成功"""
        pass

