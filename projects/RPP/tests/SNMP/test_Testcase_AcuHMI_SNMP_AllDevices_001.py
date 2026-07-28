"""
Testcase_AcuHMI_SNMP_AllDevices_001: SC-06: 全选所有设备，walk -os:.1，验证 AcuRev1300 全量参数
"""
import sys
import os
import pytest
import logging
sys.path.insert(0, os.path.dirname(__file__))
from helpers_data import (  # noqa: F401
    SNMPDataBase,
    _preread_modbus_snapshot, _select_and_walk, _compare_with_modbus,
)
log = logging.getLogger(__name__)


class Test_Testcase_AcuHMI_SNMP_AllDevices_001(SNMPDataBase):
    """Testcase_AcuHMI_SNMP_AllDevices_001"""

    def test_Testcase_AcuHMI_SNMP_AllDevices_001(self):
        """SC-06: 全选所有设备，walk -os:.1，验证 AcuRev1300 全量参数"""
        log.info("=" * 70)
        log.info("SC-06: 全选所有设备，验证 AcuRev1300")
        dev = self._get_dev("AcuRev1300")
        snmp_data = _select_and_walk(None, "SC-06 全选")
        assert len(snmp_data) > 0, "SNMP walk 返回空数据"
        modbus_snapshot = _preread_modbus_snapshot(dev)
        pass_c, fail_c, skip_c, failures = _compare_with_modbus(dev, snmp_data, modbus_snapshot)
        if pass_c == 0 and fail_c == 0:
            pytest.fail(f"SC-06: Modbus 全部读取失败（skip={skip_c}），请检查设备连接")
        if fail_c > 0:
            pytest.fail(f"SC-06: {fail_c} 个参数超出容差:\n" + "\n".join(failures))
        log.info("SC-06 PASS")
