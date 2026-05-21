# -*- coding: utf-8 -*-
"""
modbus_reader.py — 通用 Modbus TCP 读取模块（pymodbus 3.x async）

设备无关层：参数地址映射由 config.DEVICE_MODULE 指定的设备模块提供。
  IP:   config.MODBUS_HOST
  Port: config.MODBUS_PORT
  Unit: config.MODBUS_UNIT

支持的功能码：
  FC01 — Read Coils（线圈，bit）
  FC02 — Read Discrete Inputs（离散输入，bit）
  FC03 — Read Holding Registers（保持寄存器）

数据类型：
  float32 — 2 个寄存器，IEEE-754 32-bit big-endian（高字优先）
  float64 — 4 个寄存器，IEEE-754 64-bit big-endian
  uint32  — 2 个寄存器，无符号整数
  bit     — 1 个线圈/离散输入，解码为 0.0 / 1.0

用法：
  import asyncio
  from modbus_reader import ModbusReader

  async def main():
      async with ModbusReader() as reader:
          results = await reader.read_params(['FREQ_Hz', 'VLN_a_V', 'I_a_A'])
          for r in results:
              print(r)

  asyncio.run(main())
"""

from __future__ import annotations

import asyncio
import importlib
import logging
import struct
from dataclasses import dataclass
from typing import Optional

from pymodbus.client import AsyncModbusTcpClient

import config

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ModbusRegister:
    """单个 Modbus 参数描述。"""
    param_key: str
    address:   int    # 寄存器/线圈地址（十进制）
    dtype:     str    # 'float32' | 'float64' | 'uint32' | 'uint16' | 'bit'
    fc:        int   = 3    # Modbus 功能码：1=线圈, 2=离散输入, 3=保持寄存器
    scale:     float = 1.0  # 解码后乘以此系数（uint16 等需缩放时使用）

    @property
    def count(self) -> int:
        if self.dtype == 'float64':
            return 4
        elif self.dtype in ('bit', 'uint16', 'int16'):
            return 1
        else:
            return 2  # float32, uint32, int32


@dataclass
class ModbusResult:
    """单次读取结果，与 bacnet_reader.ReadResult 结构对齐。"""
    param_key: str
    value:     Optional[float] = None
    error:     str = ""

    @property
    def ok(self) -> bool:
        return not self.error and self.value is not None

    def __str__(self) -> str:
        if self.ok:
            return f"{self.param_key} = {self.value}"
        return f"{self.param_key} ERROR: {self.error}"


# ─────────────────────────────────────────────────────────────────────────────
# 设备参数映射 — 动态加载
# ─────────────────────────────────────────────────────────────────────────────

# 模块级缓存（每次进程只加载一次）
_PARAM_MAP: Optional[dict[str, ModbusRegister]] = None


def get_param_map() -> dict[str, ModbusRegister]:
    """
    从 config.DEVICE_MODULE 指定的设备模块加载 param_key → ModbusRegister 映射。

    设备模块须实现：
        build_param_map() -> dict[str, ModbusRegister]
    """
    global _PARAM_MAP
    if _PARAM_MAP is None:
        module = importlib.import_module(config.DEVICE_MODULE)
        _PARAM_MAP = module.build_param_map()
        log.info("已加载设备映射：%s（%d 个参数）",
                 config.DEVICE_MODULE, len(_PARAM_MAP))
    return _PARAM_MAP


# ─────────────────────────────────────────────────────────────────────────────
# 寄存器解码工具
# ─────────────────────────────────────────────────────────────────────────────

def _decode_registers(registers: list[int], dtype: str) -> float:
    """将 Modbus 寄存器列表解码为浮点值（big-endian 字节序，高字优先）。"""
    if dtype == 'float32':
        raw = struct.pack('>HH', registers[0], registers[1])
        return struct.unpack('>f', raw)[0]
    elif dtype == 'float64':
        raw = struct.pack('>HHHH', *registers[:4])
        return struct.unpack('>d', raw)[0]
    elif dtype == 'uint32':
        raw = struct.pack('>HH', registers[0], registers[1])
        return float(struct.unpack('>I', raw)[0])
    elif dtype == 'uint16':
        return float(registers[0])
    elif dtype == 'int16':
        raw = struct.pack('>H', registers[0])
        return float(struct.unpack('>h', raw)[0])
    elif dtype == 'int32':
        raw = struct.pack('>HH', registers[0], registers[1])
        return float(struct.unpack('>i', raw)[0])
    else:
        raise ValueError(f"未知数据类型: {dtype}")


# ─────────────────────────────────────────────────────────────────────────────
# Modbus 读取器
# ─────────────────────────────────────────────────────────────────────────────

# 单次 FC03 请求最多读取的寄存器数量
_MAX_REGS_PER_REQUEST = 125


class ModbusReader:
    """
    Modbus TCP 异步客户端，批量读取 AcuRev4100 下挂设备参数。

    以 async context manager 方式使用：
        async with ModbusReader() as reader:
            results = await reader.read_params(['FREQ_Hz', 'VLN_a_V'])
    """

    def __init__(self) -> None:
        self._client: Optional[AsyncModbusTcpClient] = None
        self._param_map = get_param_map()

    # ── async context manager ─────────────────────────────────────────────────

    async def __aenter__(self) -> "ModbusReader":
        log.info("连接 Modbus TCP  %s:%d  unit=%d  (设备=%s)",
                 config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT,
                 config.DEVICE_NAME)
        self._client = AsyncModbusTcpClient(
            host=config.MODBUS_HOST,
            port=config.MODBUS_PORT,
        )
        await self._client.connect()
        if not self._client.connected:
            raise ConnectionError(
                f"Modbus TCP 连接失败: {config.MODBUS_HOST}:{config.MODBUS_PORT}"
            )
        log.info("Modbus TCP 已连接")
        return self

    async def __aexit__(self, *args) -> None:
        if self._client is not None:
            self._client.close()
            self._client = None
            log.info("Modbus TCP 已断开")

    # ── 参数读取 ──────────────────────────────────────────────────────────────

    async def read_param(self, param_key: str) -> ModbusResult:
        """读取单个参数。"""
        reg = self._param_map.get(param_key)
        if reg is None:
            return ModbusResult(param_key=param_key, error=f"参数未在地址表中：{param_key}")
        return await self._read_reg(reg)

    async def read_params(
        self,
        param_keys: list[str],
        progress_cb=None,
    ) -> list[ModbusResult]:
        """
        批量读取参数列表，按功能码分组批量请求。
        FC03 每批最多 125 个寄存器；FC01/FC02 每批最多 2000 个线圈。

        Args:
            param_keys:  参数名称列表（param_key）
            progress_cb: 可选回调 f(done, total)

        Returns:
            ModbusResult 列表，顺序与 param_keys 一致
        """
        unknown_keys: set[str] = set()
        fc3_regs: list[ModbusRegister] = []
        fc1_regs: list[ModbusRegister] = []
        fc2_regs: list[ModbusRegister] = []

        for k in param_keys:
            r = self._param_map.get(k)
            if r is None:
                unknown_keys.add(k)
            elif r.fc == 1:
                fc1_regs.append(r)
            elif r.fc == 2:
                fc2_regs.append(r)
            else:
                fc3_regs.append(r)

        # raw_cache: param_key → list[int] (FC03) or bool (FC01/FC02)
        raw_cache: dict[str, object] = {}

        await self._batch_fc3(fc3_regs, raw_cache)
        await self._batch_bits(fc1_regs, raw_cache, fc=1)
        await self._batch_bits(fc2_regs, raw_cache, fc=2)

        results: list[ModbusResult] = []
        total = len(param_keys)
        for i, key in enumerate(param_keys):
            if key in unknown_keys:
                results.append(ModbusResult(param_key=key, error="未在地址表中"))
            else:
                reg = self._param_map[key]
                raw = raw_cache.get(key)
                if raw is None:
                    results.append(ModbusResult(param_key=key, error="读取失败"))
                else:
                    try:
                        if reg.dtype == 'bit':
                            value = 1.0 if raw else 0.0
                        else:
                            value = _decode_registers(raw, reg.dtype) * reg.scale
                        results.append(ModbusResult(param_key=key, value=value))
                    except Exception as exc:
                        results.append(ModbusResult(param_key=key, error=str(exc)))

            if progress_cb:
                ret = progress_cb(i + 1, total)
                if asyncio.iscoroutine(ret):
                    await ret

        return results

    async def _batch_fc3(
        self,
        regs: list[ModbusRegister],
        cache: dict[str, object],
    ) -> None:
        """批量执行 FC03，结果写入 cache[param_key] = list[int]。"""
        if not regs:
            return
        regs_sorted = sorted(regs, key=lambda r: r.address)
        batch: list[ModbusRegister] = []

        async def _flush() -> None:
            if not batch:
                return
            start = batch[0].address
            end   = batch[-1].address + batch[-1].count
            try:
                raw = await self._fc03(start, end - start)
                for r in batch:
                    off = r.address - start
                    cache[r.param_key] = raw[off: off + r.count]
            except IOError as exc:
                log.warning("FC03 批量读取失败（addr=0x%04X，count=%d）: %s",
                            start, end - start, exc)
                for r in batch:
                    cache[r.param_key] = None

        for reg in regs_sorted:
            prev_end = batch[-1].address + batch[-1].count if batch else None
            if not batch:
                batch.append(reg)
            elif (reg.address == prev_end                          # 严格连续
                  and (reg.address + reg.count - batch[0].address) <= _MAX_REGS_PER_REQUEST):
                batch.append(reg)
            else:
                await _flush()
                batch = [reg]
        await _flush()

    async def _batch_bits(
        self,
        regs: list[ModbusRegister],
        cache: dict[str, object],
        fc: int,
    ) -> None:
        """批量执行 FC01 或 FC02，结果写入 cache[param_key] = bool。"""
        if not regs:
            return
        regs_sorted = sorted(regs, key=lambda r: r.address)
        batch: list[ModbusRegister] = []
        _MAX_BITS = 2000

        async def _flush() -> None:
            if not batch:
                return
            start = batch[0].address
            end   = batch[-1].address + 1
            try:
                bits = await (self._fc01(start, end - start) if fc == 1
                              else self._fc02(start, end - start))
                for r in batch:
                    cache[r.param_key] = bits[r.address - start]
            except IOError as exc:
                log.warning("FC%02d 批量读取失败（addr=0x%04X，count=%d）: %s",
                            fc, start, end - start, exc)
                for r in batch:
                    cache[r.param_key] = None

        for reg in regs_sorted:
            if not batch:
                batch.append(reg)
            elif reg.address == batch[-1].address + 1:  # 严格连续才合批
                batch.append(reg)
            else:
                await _flush()
                batch = [reg]
        await _flush()

    async def read_all_mapped(self, progress_cb=None) -> list[ModbusResult]:
        """读取地址表中所有已映射的参数（约 700 个）。"""
        keys = list(self._param_map.keys())
        return await self.read_params(keys, progress_cb=progress_cb)

    # ── 底层 FC01 / FC02 / FC03 ───────────────────────────────────────────────

    async def _fc03(self, address: int, count: int) -> list[int]:
        """执行 FC03 读取，返回寄存器列表。失败时抛出异常。"""
        assert self._client and self._client.connected, "未连接"
        last_err = ""
        for attempt in range(config.MAX_RETRIES + 1):
            try:
                resp = await asyncio.wait_for(
                    self._client.read_holding_registers(
                        address=address,
                        count=count,
                        device_id=config.MODBUS_UNIT,
                    ),
                    timeout=config.READ_TIMEOUT,
                )
                if resp.isError():
                    raise IOError(f"FC03 错误响应: {resp}")
                return resp.registers
            except asyncio.TimeoutError:
                last_err = f"超时 (地址=0x{address:04X}, count={count})"
            except Exception as exc:
                last_err = str(exc)
            if attempt < config.MAX_RETRIES:
                log.debug("FC03 第%d次重试 addr=0x%04X: %s", attempt + 1, address, last_err)
                await asyncio.sleep(0.2)
        raise IOError(last_err)

    async def _fc01(self, address: int, count: int) -> list[bool]:
        """执行 FC01 Read Coils，返回 bool 列表。失败时抛出异常。"""
        assert self._client and self._client.connected, "未连接"
        last_err = ""
        for attempt in range(config.MAX_RETRIES + 1):
            try:
                resp = await asyncio.wait_for(
                    self._client.read_coils(
                        address=address,
                        count=count,
                        device_id=config.MODBUS_UNIT,
                    ),
                    timeout=config.READ_TIMEOUT,
                )
                if resp.isError():
                    raise IOError(f"FC01 错误响应: {resp}")
                return resp.bits[:count]
            except asyncio.TimeoutError:
                last_err = f"超时 (线圈=0x{address:04X}, count={count})"
            except Exception as exc:
                last_err = str(exc)
            if attempt < config.MAX_RETRIES:
                log.debug("FC01 第%d次重试 addr=0x%04X: %s", attempt + 1, address, last_err)
                await asyncio.sleep(0.2)
        raise IOError(last_err)

    async def _fc02(self, address: int, count: int) -> list[bool]:
        """执行 FC02 Read Discrete Inputs，返回 bool 列表。失败时抛出异常。"""
        assert self._client and self._client.connected, "未连接"
        last_err = ""
        for attempt in range(config.MAX_RETRIES + 1):
            try:
                resp = await asyncio.wait_for(
                    self._client.read_discrete_inputs(
                        address=address,
                        count=count,
                        device_id=config.MODBUS_UNIT,
                    ),
                    timeout=config.READ_TIMEOUT,
                )
                if resp.isError():
                    raise IOError(f"FC02 错误响应: {resp}")
                return resp.bits[:count]
            except asyncio.TimeoutError:
                last_err = f"超时 (离散输入=0x{address:04X}, count={count})"
            except Exception as exc:
                last_err = str(exc)
            if attempt < config.MAX_RETRIES:
                log.debug("FC02 第%d次重试 addr=0x%04X: %s", attempt + 1, address, last_err)
                await asyncio.sleep(0.2)
        raise IOError(last_err)

    async def _read_reg(self, reg: ModbusRegister) -> ModbusResult:
        """读取单个参数，返回 ModbusResult。"""
        try:
            if reg.fc == 1:
                bits = await self._fc01(reg.address, 1)
                return ModbusResult(param_key=reg.param_key, value=1.0 if bits[0] else 0.0)
            elif reg.fc == 2:
                bits = await self._fc02(reg.address, 1)
                return ModbusResult(param_key=reg.param_key, value=1.0 if bits[0] else 0.0)
            else:
                raw = await self._fc03(reg.address, reg.count)
                value = _decode_registers(raw, reg.dtype)
                return ModbusResult(param_key=reg.param_key, value=value)
        except Exception as exc:
            return ModbusResult(param_key=reg.param_key, error=str(exc))

    # ── 工具 ──────────────────────────────────────────────────────────────────

    def known_params(self) -> list[str]:
        """返回所有已映射的参数名称列表。"""
        return list(self._param_map.keys())

    def lookup(self, param_key: str) -> Optional[ModbusRegister]:
        """查询参数地址信息，未知参数返回 None。"""
        return self._param_map.get(param_key)

    @staticmethod
    def summary(results: list[ModbusResult]) -> dict:
        """返回读取结果统计。"""
        ok  = [r for r in results if r.ok]
        err = [r for r in results if not r.ok]
        return {
            "total":        len(results),
            "success":      len(ok),
            "failed":       len(err),
            "success_rate": f"{len(ok)/len(results)*100:.1f}%" if results else "N/A",
            "errors":       {r.param_key: r.error for r in err},
        }


# ─────────────────────────────────────────────────────────────────────────────
# Modbus RTU 读取器（串口，设备与电脑不同网段时使用）
# ─────────────────────────────────────────────────────────────────────────────

class ModbusRtuReader(ModbusReader):
    """
    Modbus RTU 串口读取器。

    与 ModbusReader 接口完全一致，仅连接方式不同：
    使用 AsyncModbusSerialClient 通过 RS485 串口通信，
    适用于下挂设备与电脑不同网段、无法走 Modbus TCP 的场景。

    串口参数来自 config.MODBUS_RTU_* 配置项。
    slave 地址仍使用 config.MODBUS_UNIT（与 TCP 模式保持一致）。
    """

    async def __aenter__(self) -> "ModbusRtuReader":
        from pymodbus.client import AsyncModbusSerialClient
        log.info("连接 Modbus RTU  port=%s  baud=%d  unit=%d  (设备=%s)",
                 config.MODBUS_RTU_PORT, config.MODBUS_RTU_BAUDRATE,
                 config.MODBUS_UNIT, config.DEVICE_NAME)
        self._client = AsyncModbusSerialClient(
            port     = config.MODBUS_RTU_PORT,
            baudrate = config.MODBUS_RTU_BAUDRATE,
            parity   = config.MODBUS_RTU_PARITY,
            stopbits = config.MODBUS_RTU_STOPBITS,
            bytesize = config.MODBUS_RTU_BYTESIZE,
        )
        await self._client.connect()
        if not self._client.connected:
            raise ConnectionError(
                f"Modbus RTU 连接失败: {config.MODBUS_RTU_PORT} "
                f"(baud={config.MODBUS_RTU_BAUDRATE})"
            )
        log.info("Modbus RTU 已连接（%s）", config.MODBUS_RTU_PORT)
        return self

    async def __aexit__(self, *args) -> None:
        if self._client is not None:
            self._client.close()
            self._client = None
            log.info("Modbus RTU 已断开（%s）", config.MODBUS_RTU_PORT)


# ─────────────────────────────────────────────────────────────────────────────
# 工厂函数：根据 config.MODBUS_MODE 返回合适的读取器
# ─────────────────────────────────────────────────────────────────────────────

def get_reader() -> ModbusReader:
    """
    根据 config.MODBUS_MODE 返回 TCP 或 RTU 读取器实例。

    返回类型均为 ModbusReader（或其子类），可作为 async context manager 使用：
        async with get_reader() as modbus:
            results = await modbus.read_params([...])
    """
    mode = getattr(config, 'MODBUS_MODE', 'tcp').lower()
    if mode == 'rtu':
        return ModbusRtuReader()
    return ModbusReader()
