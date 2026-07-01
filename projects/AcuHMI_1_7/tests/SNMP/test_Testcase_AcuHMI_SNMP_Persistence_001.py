"""
Testcase_AcuHMI_SNMP_Persistence_001: FTS_case08: 重启设备后配置不丢失，数据可正常读取
"""
import sys, os
import pytest
sys.path.insert(0, os.path.dirname(__file__))
from helpers_ui import SNMPBase, snmp_page  # noqa: F401


class Test_Testcase_AcuHMI_SNMP_Persistence_001(SNMPBase):
    """Testcase_AcuHMI_SNMP_Persistence_001"""

    @pytest.mark.skip(reason="需要手动操作：重启 AcuHMI 设备")
    def test_Testcase_AcuHMI_SNMP_Persistence_001(self, snmp_page):
        """FTS_case08: 重启设备后配置不丢失，数据可正常读取"""
        pass

