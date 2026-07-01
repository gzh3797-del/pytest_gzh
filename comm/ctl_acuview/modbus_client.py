"""Modbus TCP/RTU 直读直写。

既是"读取/下发"的底层能力，也是 GUI 操作后的 *校验真值源*。
按寄存器名或地址读写；按 Excel 数据类型自动编解码；带字序自动校准与范围校验。

pymodbus 3.7 API：
    client.read_holding_registers(address, count=N, slave=id) -> rr; rr.registers
    client.write_registers(address, [vals], slave=id)
"""
from __future__ import annotations

import struct
import sys
from dataclasses import dataclass

from pymodbus.client import ModbusSerialClient, ModbusTcpClient

from .config import get_config
from .spec_loader import load_spec, reg_by_addr, reg_by_name

# ---- 数据类型规范化 ----
_DTYPE_ALIASES = {
    "uint16_t": "u16", "uint16": "u16", "word": "u16",
    "int16_t": "i16", "int16": "i16",
    "uint32_t": "u32", "uint32": "u32", "dword": "u32",
    "int32_t": "i32", "int32": "i32",
    "float": "f32", "float32": "f32", "real": "f32",
    "uint8_t": "u16",   # 8 位通常占低字节，按 u16 原值读出
    "int8_t": "i16",
}
_REGS_PER = {"u16": 1, "i16": 1, "u32": 2, "i32": 2, "f32": 2}


def norm_dtype(dtype: str) -> str:
    d = str(dtype).strip().lower()
    if d in _REGS_PER:          # 已是规范码(u16/i16/u32/i32/f32)
        return d
    return _DTYPE_ALIASES.get(d, "u16")


# ---- 字序：AcuRev 多为大端字序(高字在前)。可由 selftest 自动校准。----
@dataclass
class WordOrder:
    word_big: bool = True   # True: 高 16 位寄存器在前
    byte_big: bool = True   # True: 寄存器内大端

    def regs_to_bytes(self, regs: list[int]) -> bytes:
        order = regs if self.word_big else list(reversed(regs))
        fmt = ">" if self.byte_big else "<"
        return b"".join(struct.pack(fmt + "H", r & 0xFFFF) for r in order)

    def bytes_to_regs(self, raw: bytes) -> list[int]:
        fmt = ">" if self.byte_big else "<"
        regs = [struct.unpack(fmt + "H", raw[i:i + 2])[0] for i in range(0, len(raw), 2)]
        return regs if self.word_big else list(reversed(regs))


def decode(regs: list[int], dtype: str, wo: WordOrder):
    t = norm_dtype(dtype)
    if t == "u16":
        return regs[0] & 0xFFFF
    if t == "i16":
        return struct.unpack(">h", struct.pack(">H", regs[0] & 0xFFFF))[0]
    raw = wo.regs_to_bytes(regs[:_REGS_PER[t]])
    if t == "u32":
        return struct.unpack(">I", raw)[0]
    if t == "i32":
        return struct.unpack(">i", raw)[0]
    if t == "f32":
        return struct.unpack(">f", raw)[0]
    return regs[0]


def encode(value, dtype: str, wo: WordOrder) -> list[int]:
    t = norm_dtype(dtype)
    if t == "u16":
        return [int(round(value)) & 0xFFFF]
    if t == "i16":
        return [struct.unpack(">H", struct.pack(">h", int(round(value))))[0]]
    if t == "u32":
        raw = struct.pack(">I", int(round(value)) & 0xFFFFFFFF)
    elif t == "i32":
        raw = struct.pack(">i", int(round(value)))
    elif t == "f32":
        raw = struct.pack(">f", float(value))
    else:
        return [int(round(value)) & 0xFFFF]
    return wo.bytes_to_regs(raw)


class ModbusError(RuntimeError):
    pass


class MeterClient:
    """对一路 Modbus 连接(TCP 或 RTU)的统一封装。"""

    def __init__(self, transport: str | None = None, word_order: WordOrder | None = None):
        self.cfg = get_config()
        self.transport = transport or self.cfg.transport.verify
        self.wo = word_order or WordOrder()
        self.registers, self.pages = load_spec()
        self._client = None
        self._slave = None

    # ---- 连接管理 ----
    def connect(self):
        if self.transport == "tcp":
            t = self.cfg.transport.tcp
            self._client = ModbusTcpClient(host=t["host"], port=t["port"], timeout=t["timeout_s"])
            self._slave = t["slave_id"]
        elif self.transport == "rtu":
            t = self.cfg.transport.rtu
            self._client = ModbusSerialClient(
                port=t["port"], baudrate=t["baudrate"], parity=t["parity"],
                bytesize=t["bytesize"], stopbits=t["stopbits"], timeout=t["timeout_s"],
            )
            self._slave = t["slave_id"]
        else:
            raise ModbusError(f"未知传输方式: {self.transport}")
        if not self._client.connect():
            raise ModbusError(f"无法连接 {self.transport} ({t})")
        return self

    def close(self):
        if self._client:
            self._client.close()
            self._client = None

    def __enter__(self):
        return self.connect()

    def __exit__(self, *exc):
        self.close()

    # ---- 兼容不同 pymodbus 版本的 slave 关键字 ----
    def _read_regs(self, address: int, count: int) -> list[int]:
        try:
            rr = self._client.read_holding_registers(address, count=count, slave=self._slave)
        except TypeError:
            rr = self._client.read_holding_registers(address, count, slave=self._slave)
        if rr is None or rr.isError():
            raise ModbusError(f"读寄存器失败 addr={address} count={count}: {rr}")
        return list(rr.registers)

    def _write_regs(self, address: int, values: list[int]):
        # 统一用 FC16(write multiple)，即使单寄存器：AcuRev 对 FC6 的响应帧会让
        # pymodbus 解析报错(struct unpack 4 bytes)，FC16 响应规整、更通用。
        try:
            rr = self._client.write_registers(address, values, slave=self._slave)
        except TypeError:
            rr = self._client.write_registers(address, values, self._slave)
        if rr is None or rr.isError():
            raise ModbusError(f"写寄存器失败 addr={address} vals={values}: {rr}")
        return rr

    # ---- 按地址读 ----
    def read_addr(self, addr: int, dtype: str = "u16", count: int | None = None):
        n = count or _REGS_PER.get(norm_dtype(dtype), 1)
        regs = self._read_regs(addr, n)
        return decode(regs, dtype, self.wo)

    def read_block(self, start_addr: int, reg_num: int) -> list[int]:
        """读原始寄存器块(给读序列用)。"""
        return self._read_regs(start_addr, reg_num)

    # ---- 按名读/写(用 spec 解析 dtype 与地址) ----
    def _resolve(self, name_or_addr) -> dict:
        if isinstance(name_or_addr, int):
            e = reg_by_addr(self.registers, name_or_addr)
        else:
            e = reg_by_name(self.registers, name_or_addr)
        if not e:
            raise ModbusError(f"spec 中找不到寄存器: {name_or_addr}")
        return e

    def read(self, name_or_addr):
        e = self._resolve(name_or_addr)
        return self.read_addr(e["addr"], e["dtype"], e["reg_num"])

    def write(self, name_or_addr, value, check_range: bool = True, check_rw: bool = True):
        e = self._resolve(name_or_addr)
        if check_rw and "W" not in e["rw"].upper():
            raise ModbusError(f"寄存器只读，禁止写: {e['description']} (RW={e['rw']})")
        if check_range:
            self._validate_range(e, value)
        vals = encode(value, e["dtype"], self.wo)
        self._write_regs(e["addr"], vals)
        return vals

    @staticmethod
    def _validate_range(e: dict, value):
        import re
        m = re.search(r"(-?\d+)\s*[-~]\s*(-?\d+)", e.get("range", ""))
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            if not (lo <= value <= hi):
                raise ModbusError(f"值 {value} 超出范围 [{lo},{hi}] ({e['description']})")

    # ---- 字序自动校准 ----
    def calibrate_word_order(self, probe_addr: int = 12288, dtype: str = "f32",
                             expected_range=(40.0, 70.0)) -> WordOrder:
        """用 System Frequency(0x3000, ~50/60Hz) 试探，挑选给出合理值的字序。"""
        best = None
        for wb in (True, False):
            for bb in (True, False):
                wo = WordOrder(word_big=wb, byte_big=bb)
                try:
                    regs = self._read_regs(probe_addr, _REGS_PER[norm_dtype(dtype)])
                    val = decode(regs, dtype, wo)
                except ModbusError:
                    continue
                if expected_range[0] <= val <= expected_range[1]:
                    best = wo
                    break
            if best:
                break
        if best:
            self.wo = best
        return self.wo


def _selftest():
    """连校验通道(默认 tcp)，校准字序，抽读若干寄存器。"""
    cfg = get_config()
    client = MeterClient(transport=cfg.transport.verify)
    print(f"[selftest] 连接 {client.transport} ...")
    with client:
        wo = client.calibrate_word_order()
        print(f"[selftest] 字序校准: word_big={wo.word_big} byte_big={wo.byte_big}")
        samples = [
            ("System Frequency", 12288, "f32"),
            ("Phase A Line-to-Neutral Voltage", 12290, "f32"),
            ("Modbus Slave ID", 4111, "u16"),
            ("Baud Rate", 4097, "u16"),
        ]
        for label, addr, dt in samples:
            try:
                v = client.read_addr(addr, dt)
                print(f"  {label:<36} @{addr}(0x{addr:04X}) = {v}")
            except ModbusError as exc:
                print(f"  {label:<36} @{addr} 读取失败: {exc}")
    print("[selftest] 完成")


if __name__ == "__main__":
    if "--selftest" in sys.argv or len(sys.argv) == 1:
        _selftest()
