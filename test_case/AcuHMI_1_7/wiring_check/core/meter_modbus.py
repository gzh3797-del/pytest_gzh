import time
import logging
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '../../../..'))
from pymodbus.client import ModbusTcpClient
from test_case.AcuHMI_1_7.wiring_check.core import config as cfg


class WiringCheckModbus:
    def __init__(self):
        self._client = ModbusTcpClient(host=cfg.METER_TCP_IP, port=cfg.METER_TCP_PORT)
        self._client.connect()

    def close(self):
        self._client.close()

    def _write(self, address: int, values: list):
        """FC16 写多寄存器；若解码异常则降级用 FC6 逐个写"""
        try:
            resp = self._client.write_registers(
                address=address, values=values, device_id=cfg.MODBUS_SLAVE)
            if resp.isError():
                raise ValueError(str(resp))
        except Exception:
            for offset, val in enumerate(values):
                self._client.write_register(
                    address=address + offset, value=val, device_id=cfg.MODBUS_SLAVE)

    def _read(self, address: int, count: int):
        resp = self._client.read_holding_registers(
            address=address, count=count, device_id=cfg.MODBUS_SLAVE)
        return resp.registers if not resp.isError() else None

    # ── 配置写入 ─────────────────────────────────────────────────────────────

    def write_service_config(self, mode: int):
        """写接线方式，见 config.SERVICE_* 常量"""
        self._write(cfg.REG_SERVICE_CONFIG, [mode])
        time.sleep(1.0)   # 等设备处理接线方式切换（可能触发内部重置）

    def write_phase_order(self, order: int):
        """写相序：0=ABC，1=ACB"""
        self._write(cfg.REG_PHASE_ORDER, [order])
        time.sleep(0.5)   # 等设备处理写入

    def write_nominal_voltage(self, voltage: int = cfg.NOMINAL_VOLTAGE):
        """写额定电压（uint32，拆成两个不连续寄存器）"""
        low  = voltage & 0xFFFF
        high = (voltage >> 16) & 0xFFFF
        self._write(cfg.REG_NOMINAL_VOLTAGE_L, [low])
        self._write(cfg.REG_NOMINAL_VOLTAGE_H, [high])

    # ── 检查触发 / 轮询 ──────────────────────────────────────────────────────

    def trigger_check(self):
        """写 0x1300=1 触发接线检查"""
        self._write(cfg.REG_WIRE_CHECK_START, [1])

    def wait_for_completion(self, timeout: int = cfg.CHECK_TIMEOUT) -> bool:
        """轮询 0x1301，直到状态=2（Completed）或超时"""
        deadline = time.time() + timeout
        while time.time() < deadline:
            regs = self._read(cfg.REG_WIRE_CHECK_STATUS, 1)
            if regs and regs[0] == 2:
                return True
            time.sleep(1)
        logging.error('Wiring check did not complete within %ds', timeout)
        return False

    # ── 结果读取 ─────────────────────────────────────────────────────────────

    def read_voltage_error(self) -> int | None:
        regs = self._read(cfg.REG_VOLTAGE_ERROR, 1)
        return regs[0] if regs else None

    def read_current_errors(self) -> list[int | None]:
        regs = self._read(cfg.REG_CURRENT_ERROR_BASE, cfg.NUM_USER_CHANNELS)
        return list(regs) if regs else [None] * cfg.NUM_USER_CHANNELS

    # ── 一次性配置 ───────────────────────────────────────────────────────────

    def setup_3e4wy(self, phase_order: int = cfg.PHASE_ABC,
                    nominal_voltage: int = cfg.NOMINAL_VOLTAGE):
        """3E4WY 测试前写入全部固定配置"""
        self.write_service_config(cfg.SERVICE_3E4WY)
        self.write_phase_order(phase_order)
        self.write_nominal_voltage(nominal_voltage)
        logging.info('Meter configured: 3E4WY, phase_order=%d, VRATE=%dV',
                     phase_order, nominal_voltage)
