"""
电表 Modbus RTU 通信模块（独立，不依赖 autotest 工程）
支持 pymodbus 标准 RTU 串口通信
所有寄存器地址从外部传入，不硬编码
"""
import struct
import time
import logging
import statistics
from pymodbus.client import ModbusSerialClient


class MeterComm:
    def __init__(self, port: str, baudrate: int = 19200, parity: str = 'N', slave_id: int = 1):
        self.port = port
        self.baudrate = baudrate
        self.parity = parity
        self.slave_id = slave_id
        self.client = None
        self._connected = False

    def connect(self) -> bool:
        try:
            self.client = ModbusSerialClient(
                port=self.port,
                baudrate=self.baudrate,
                parity=self.parity,
                stopbits=1,
                bytesize=8,
                timeout=2,
            )
            self.client.inter_byte_timeout = 0.1
            self._connected = self.client.connect()
            return self._connected
        except Exception as e:
            logging.error(f"Meter connect error: {e}")
            self._connected = False
            return False

    def disconnect(self):
        if self.client:
            self.client.close()
            self._connected = False

    @property
    def connected(self) -> bool:
        return self._connected

    # ── 底层读写 ──────────────────────────────────────────────

    def read_float(self, address: int) -> float | None:
        """读 2 个寄存器，解析为 big-endian float32"""
        for attempt in range(5):
            try:
                resp = self.client.read_holding_registers(address=address, count=2, device_id=self.slave_id)
                if resp and not resp.isError():
                    raw = resp.registers
                    b = bytes([
                        (raw[0] >> 8) & 0xff, raw[0] & 0xff,
                        (raw[1] >> 8) & 0xff, raw[1] & 0xff,
                    ])
                    return struct.unpack('!f', b)[0]
            except Exception as e:
                logging.warning(f"read_float addr={hex(address)} attempt {attempt+1}: {e}")
                time.sleep(0.1)
        return None

    def read_uint16(self, address: int) -> int | None:
        """读 1 个寄存器，解析为 uint16"""
        for attempt in range(5):
            try:
                resp = self.client.read_holding_registers(address=address, count=1, device_id=self.slave_id)
                if resp and not resp.isError():
                    return resp.registers[0]
            except Exception as e:
                logging.warning(f"read_uint16 addr={hex(address)} attempt {attempt+1}: {e}")
                time.sleep(0.1)
        return None

    def write_uint16(self, address: int, value: int) -> bool:
        """写单个寄存器"""
        for attempt in range(3):
            try:
                resp = self.client.write_registers(address=address, values=[value], device_id=self.slave_id)
                if resp and not resp.isError():
                    return True
            except Exception as e:
                logging.warning(f"write_uint16 addr={hex(address)} attempt {attempt+1}: {e}")
                time.sleep(0.1)
        return False

    # ── 业务级读取（地址由外部地址字典传入）────────────────

    def read_measure_batch(self, addr_map: dict) -> dict:
        """
        批量读取测量量：对连续地址段发一次 read_holding_registers，
        避免逐个请求带来的多次 RTU 往返延迟。
        addr_map 示例: {"ua": 0x9309, "ub": 0x930B, ...}
        """
        if not addr_map:
            return {}

        sorted_items = sorted(addr_map.items(), key=lambda x: x[1])
        min_addr = sorted_items[0][1]
        max_addr = sorted_items[-1][1]
        count = (max_addr + 2) - min_addr   # 每个 float32 占 2 个寄存器

        if count <= 125:   # Modbus 单次最多读 125 个寄存器
            for attempt in range(3):
                try:
                    resp = self.client.read_holding_registers(
                        address=min_addr, count=count, device_id=self.slave_id)
                    if resp and not resp.isError():
                        regs = resp.registers
                        result = {}
                        for name, addr in addr_map.items():
                            off = addr - min_addr
                            if off + 1 < len(regs):
                                b = bytes([
                                    (regs[off]   >> 8) & 0xff, regs[off]   & 0xff,
                                    (regs[off+1] >> 8) & 0xff, regs[off+1] & 0xff,
                                ])
                                result[name] = struct.unpack('!f', b)[0]
                            else:
                                result[name] = 0.0
                        return result
                except Exception as e:
                    logging.warning(f"read_measure_batch block attempt {attempt+1}: {e}")
                    time.sleep(0.1)

        # 地址不连续或超出 125 寄存器限制，退回逐个读
        result = {}
        for name, addr in addr_map.items():
            val = self.read_float(addr)
            result[name] = val if val is not None else 0.0
        return result

    # ── 设备配置 ──────────────────────────────────────────────

    def set_wire_mode(self, addr_wire_mode: int, value: int) -> bool:
        return self.write_uint16(addr_wire_mode, value)

    def set_ct_type(self, addr_ct_ch1: int, addr_ct_ch2: int, addr_ct_ch3: int, ct_value: int) -> bool:
        """统一设置三通道 CT 类型（0=100mA, 2=333mV, 3=RCT）"""
        ok = True
        for addr in [addr_ct_ch1, addr_ct_ch2, addr_ct_ch3]:
            ok = ok and self.write_uint16(addr, ct_value)
        return ok

    def set_freq_mode(self, addr_freq_set: int, freq_hz: int) -> bool:
        """设置频率：50Hz→0，60Hz→1"""
        freq_map = {50: 0, 60: 1}
        value = freq_map.get(freq_hz, 0)
        return self.write_uint16(addr_freq_set, value)

    def get_freq_mode(self, addr_freq_set: int) -> int | None:
        return self.read_uint16(addr_freq_set)

    def soft_reboot(self, addr_reboot: int) -> bool:
        """软重启：写 0x0001 到重启寄存器"""
        return self.write_uint16(addr_reboot, 0x0001)

    # ── 多次采样 ──────────────────────────────────────────────

    def sample_n_times(self, addr_map: dict, n: int, interval_ms: int) -> dict[str, list]:
        """
        采样 n 次，每次间隔 interval_ms 毫秒
        返回 {name: [v1, v2, ..., vn]}
        """
        samples: dict[str, list] = {name: [] for name in addr_map}
        for _ in range(n):
            time.sleep(interval_ms / 1000.0)
            batch = self.read_measure_batch(addr_map)
            for name, val in batch.items():
                samples[name].append(val)
        return samples

    # ── 精度计算 ──────────────────────────────────────────────

    @staticmethod
    def calc_accuracy(expected: float, measured_list: list[float], threshold: float) -> dict:
        """
        有预期精度的量（电压、电流、有功）
        返回 {min, min_err, max, max_err, avg, avg_err, pass}
        误差 = (measured - expected) / expected
        """
        if not measured_list or expected == 0:
            return {"min": 0, "min_err": 0, "max": 0, "max_err": 0,
                    "avg": 0, "avg_err": 0, "pass": False}
        errors = [(m - expected) / expected for m in measured_list]
        mn, mx, av = min(measured_list), max(measured_list), statistics.mean(measured_list)
        mn_e, mx_e, av_e = min(errors), max(errors), statistics.mean(errors)
        # 与参考脚本一致：min/max/avg 的绝对误差全部 ≤ 阈值才算 Passed
        passed = (abs(mn_e) <= threshold) and (abs(mx_e) <= threshold) and (abs(av_e) <= threshold)
        return {
            "min": round(mn, 5), "min_err": round(mn_e, 6),
            "max": round(mx, 5), "max_err": round(mx_e, 6),
            "avg": round(av, 5), "avg_err": round(av_e, 6),
            "pass": passed,
        }

    @staticmethod
    def calc_accuracy_no_threshold(expected: float, measured_list: list[float]) -> dict:
        """
        无预期精度的量（无功、视在）——只统计 min/max/avg 偏差，不判合格
        """
        if not measured_list:
            return {"min": 0, "min_err": 0, "max": 0, "max_err": 0, "avg": 0, "avg_err": 0, "pass": None}
        mn, mx, av = min(measured_list), max(measured_list), statistics.mean(measured_list)
        if expected != 0:
            mn_e = (mn - expected) / expected
            mx_e = (mx - expected) / expected
            av_e = (av - expected) / expected
        else:
            mn_e, mx_e, av_e = mn - expected, mx - expected, av - expected
        return {
            "min": round(mn, 5), "min_err": round(mn_e, 6),
            "max": round(mx, 5), "max_err": round(mx_e, 6),
            "avg": round(av, 5), "avg_err": round(av_e, 6),
            "pass": None,
        }

    @staticmethod
    def calc_phase_angle_accuracy(expected_deg: float, measured_list: list[float], threshold: float) -> dict:
        """
        相角精度：处理 360° 跨越（如期望 0°，测量到 359.x°）
        """
        if not measured_list:
            return {"min": 0, "min_err": 0, "max": 0, "max_err": 0, "avg": 0, "avg_err": 0, "pass": False}
        corrected = []
        for v in measured_list:
            if expected_deg == 0 and 350 <= v <= 360:
                corrected.append(v - 360)
            else:
                corrected.append(v)
        errors = [v - expected_deg for v in corrected]
        mn_e, mx_e, av_e = min(errors), max(errors), statistics.mean(errors)
        mn = (min(corrected) + 360) if min(corrected) < 0 else min(corrected)
        mx = (max(corrected) + 360) if max(corrected) < 0 else max(corrected)
        av = (statistics.mean(corrected) + 360) if statistics.mean(corrected) < 0 else statistics.mean(corrected)
        return {
            "min": round(mn, 4), "min_err": round(mn_e, 4),
            "max": round(mx, 4), "max_err": round(mx_e, 4),
            "avg": round(av, 4), "avg_err": round(av_e, 4),
            "pass": abs(av_e) <= threshold,
        }
