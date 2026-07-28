"""
Testcase_AcuHMI_SNMP_AcuRev4100_001: SC-01: 勾选所有 AcuRev4100 实例，walk -os:.1，逐实例全量参数与 Modbus 对比
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


class Test_Testcase_AcuHMI_SNMP_AcuRev4100_001(SNMPDataBase):
    """Testcase_AcuHMI_SNMP_AcuRev4100_001"""

    def test_Testcase_AcuHMI_SNMP_AcuRev4100_001(self):
        """SC-01: 勾选所有 AcuRev4100 实例，walk -os:.1，逐实例全量参数与 Modbus 对比"""
        log.info("=" * 70)
        log.info("SC-01: AcuRev4100 所有实例")

        snmp_names, devs = self._get_devs_by_model("AcuRev4100")
        log.info("SC-01: SNMP 勾选 %d 台, Modbus 对比 %d 台", len(snmp_names), len(devs))

        snmp_data = _select_and_walk(snmp_names, f"SC-01 AcuRev4100 {snmp_names}")
        assert len(snmp_data) > 0, "SNMP walk 返回空数据"
        snapshots = {d["name"]: _preread_modbus_snapshot(d) for d in devs}

        total_fail = 0
        data_failures = []    # 数据对比失败（SNMP ≠ Modbus）
        conn_warnings = []    # Modbus 连接失败（设备离线/IP 错误）
        for dev in devs:
            pass_c, fail_c, skip_c, failures = _compare_with_modbus(
                dev, snmp_data, snapshots[dev["name"]]
            )
            log.info("SC-01 %s: PASS=%d FAIL=%d SKIP=%d", dev["name"], pass_c, fail_c, skip_c)
            if pass_c == 0 and fail_c == 0:
                # 所有参数均跳过 → Modbus 连接失败，属于环境问题，不计入数据失败
                reason = failures[0] if failures else (
                    f"Modbus 全部读取失败（skip={skip_c}），"
                    f"请确认 {dev['modbus_host']}:{dev['modbus_port']} unit={dev['modbus_unit']} 可达"
                )
                conn_warnings.append(f"[{dev['name']}] {reason}")
                log.warning("SC-01 %s: Modbus 连接失败，跳过对比", dev["name"])
            elif fail_c > 0:
                total_fail += fail_c
                data_failures.extend(f"[{dev['name']}] {f}" for f in failures)

        # 连接失败以 warning 形式打印，不导致测试 FAIL
        if conn_warnings:
            log.warning("SC-01: 以下设备 Modbus 不可达（已跳过对比，请检查环境）:\n%s",
                        "\n".join(conn_warnings))
            print("\n[WARNING] 以下设备 Modbus 不可达（跳过对比）：")
            for w in conn_warnings:
                print(f"  {w}")

        if data_failures:
            pytest.fail(
                f"SC-01: {total_fail} 个参数超出容差 {TOLERANCE}:\n"
                + "\n".join(data_failures)
            )
        log.info("SC-01 PASS（SNMP %d 台 / Modbus对比 %d 台，连接警告 %d 台）",
                 len(snmp_names), len(devs), len(conn_warnings))
