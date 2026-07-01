"""
Testcase_AcuHMI_SNMP_AcuvimIIR_001: SC-06: 勾选所有 AcuvimIIR 实例，walk -os:.1，逐实例全量参数与 Modbus 对比
"""
import sys
import os
import pytest
import logging
sys.path.insert(0, os.path.dirname(__file__))
from helpers_data import (  # noqa: F401
    SNMPDataBase,
    _preread_modbus_snapshot, _select_and_walk, _compare_with_modbus,
    TOLERANCE,
)
log = logging.getLogger(__name__)


class Test_Testcase_AcuHMI_SNMP_AcuvimIIR_001(SNMPDataBase):
    """Testcase_AcuHMI_SNMP_AcuvimIIR_001"""

    def test_Testcase_AcuHMI_SNMP_AcuvimIIR_001(self):
        """SC-06: 勾选所有 AcuvimIIR 实例，walk -os:.1，逐实例全量参数与 Modbus 对比"""
        log.info("=" * 70)
        log.info("SC-06: AcuvimIIR 所有实例")

        snmp_names, devs = self._get_devs_by_model("AcuvimIIR")
        log.info("SC-06: SNMP 勾选 %d 台, Modbus 对比 %d 台", len(snmp_names), len(devs))
        snmp_data = _select_and_walk(snmp_names, f"SC-06 AcuvimIIR {snmp_names}")
        assert len(snmp_data) > 0, "SNMP walk 返回空数据"
        snapshots = {d["name"]: _preread_modbus_snapshot(d) for d in devs}

        total_fail = 0
        all_failures = []
        for dev in devs:
            pass_c, fail_c, skip_c, failures = _compare_with_modbus(
                dev, snmp_data, snapshots[dev["name"]]
            )
            if pass_c == 0 and fail_c == 0:
                reason = failures[0] if failures else f"Modbus 全部读取失败（skip={skip_c}），请检查设备连接"
                all_failures.append(f"[{dev['name']}] {reason}")
            elif fail_c > 0:
                total_fail += fail_c
                all_failures.extend(f"[{dev['name']}] {f}" for f in failures)
            log.info("SC-06 %s: PASS=%d FAIL=%d SKIP=%d", dev["name"], pass_c, fail_c, skip_c)

        if all_failures:
            pytest.fail(f"SC-06: {total_fail} 个参数超出容差 {TOLERANCE}:\n" + "\n".join(all_failures))
        log.info("SC-06 PASS（SNMP %d 台 / Modbus对比 %d 台）", len(snmp_names), len(devs))
