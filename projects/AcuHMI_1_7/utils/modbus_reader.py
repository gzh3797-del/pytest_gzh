"""Async Modbus TCP/RTU reader using pymodbus + excel_param_map for register info."""
from __future__ import annotations
import asyncio
import logging
import struct
from dataclasses import dataclass, field
from typing import Optional

log = logging.getLogger(__name__)

@dataclass
class ModbusResult:
    param_key: str
    ok: bool = False
    value: Optional[float] = None
    error: str = ""

_DTYPE_REGS: dict[str, int] = {
    "float": 2, "double": 4, "uint32": 2, "int32": 2,
    "word": 1, "word_signed": 1,
}

def _decode(regs: list[int], dtype: str, scale: float) -> float:
    if dtype == "float":
        b = struct.pack(">HH", regs[0], regs[1])
        v = struct.unpack(">f", b)[0]
    elif dtype == "double":
        b = struct.pack(">HHHH", *regs[:4])
        v = struct.unpack(">d", b)[0]
    elif dtype == "uint32":
        v = float((regs[0] << 16) | regs[1])
    elif dtype == "int32":
        raw = (regs[0] << 16) | regs[1]
        v = float(raw - 0x100000000 if raw >= 0x80000000 else raw)
    elif dtype == "word":
        v = float(regs[0])
    elif dtype == "word_signed":
        r = regs[0]
        v = float(r - 0x10000 if r >= 0x8000 else r)
    else:
        raise ValueError(f"Unknown dtype: {dtype}")
    return v * scale


async def read_device_params(
    device_name: str,
    param_keys: list[str],
    host: str,
    port: int,
    unit: int,
    timeout: int = 30,
) -> dict[str, ModbusResult]:
    """Read holding registers for param_keys on device_name.

    Register info (addr, type, scale) is obtained from excel_param_map.
    Returns {param_key: ModbusResult}.
    """
    from utils.excel_param_map import load_device_params
    param_map = load_device_params(device_name, snmp_only=False)

    try:
        from pymodbus.client import AsyncModbusTcpClient
    except ImportError:
        try:
            from pymodbus.client.tcp import AsyncModbusTcpClient
        except ImportError:
            return {k: ModbusResult(param_key=k, ok=False, error="pymodbus 未安装")
                    for k in param_keys}

    results: dict[str, ModbusResult] = {}

    client = AsyncModbusTcpClient(host, port=port, timeout=timeout)
    try:
        await asyncio.wait_for(client.connect(), timeout=30)
    except Exception as e:
        return {k: ModbusResult(param_key=k, ok=False, error=f"连接 {host}:{port} 失败: {e}")
                for k in param_keys}

    try:
        for key in param_keys:
            info = param_map.get(key)
            if info is None:
                results[key] = ModbusResult(param_key=key, ok=False, error="无寄存器映射")
                continue

            dtype  = info["type"]
            addr   = info["addr"]
            scale  = info["scale"]
            count  = _DTYPE_REGS.get(dtype, 2)

            try:
                rr = await client.read_holding_registers(addr, count=count, device_id=unit)
                if rr.isError():
                    results[key] = ModbusResult(param_key=key, ok=False, error=str(rr))
                else:
                    value = _decode(rr.registers, dtype, scale)
                    results[key] = ModbusResult(param_key=key, ok=True, value=value)
            except Exception as e:
                results[key] = ModbusResult(param_key=key, ok=False, error=str(e))
    finally:
        try:
            client.close()
        except Exception:
            pass

    return results


async def read_device_params_rtu(
    device_name: str,
    param_keys: list[str],
    serial_port: str,
    unit: int,
    baudrate: int = 9600,
    parity: str = "N",
    stopbits: int = 1,
    bytesize: int = 8,
    timeout: int = 30,
) -> dict[str, ModbusResult]:
    """Read holding registers for param_keys on device_name via Modbus RTU serial.

    Returns {param_key: ModbusResult}.
    """
    from utils.excel_param_map import load_device_params
    param_map = load_device_params(device_name, snmp_only=False)

    try:
        from pymodbus.client import AsyncModbusSerialClient
    except ImportError:
        try:
            from pymodbus.client.serial import AsyncModbusSerialClient
        except ImportError:
            return {k: ModbusResult(param_key=k, ok=False, error="pymodbus serial 未安装")
                    for k in param_keys}

    results: dict[str, ModbusResult] = {}

    client = AsyncModbusSerialClient(
        serial_port,
        baudrate=baudrate,
        parity=parity,
        stopbits=stopbits,
        bytesize=bytesize,
        timeout=timeout,
    )
    try:
        await asyncio.wait_for(client.connect(), timeout=30)
    except Exception as e:
        return {k: ModbusResult(param_key=k, ok=False, error=f"连接 {serial_port} 失败: {e}")
                for k in param_keys}

    try:
        for key in param_keys:
            info = param_map.get(key)
            if info is None:
                results[key] = ModbusResult(param_key=key, ok=False, error="无寄存器映射")
                continue

            dtype  = info["type"]
            addr   = info["addr"]
            scale  = info["scale"]
            count  = _DTYPE_REGS.get(dtype, 2)

            try:
                rr = await client.read_holding_registers(addr, count=count, device_id=unit)
                if rr.isError():
                    results[key] = ModbusResult(param_key=key, ok=False, error=str(rr))
                else:
                    value = _decode(rr.registers, dtype, scale)
                    results[key] = ModbusResult(param_key=key, ok=True, value=value)
            except Exception as e:
                results[key] = ModbusResult(param_key=key, ok=False, error=str(e))
    finally:
        try:
            client.close()
        except Exception:
            pass

    return results
