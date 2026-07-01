# -*- coding: utf-8 -*-
"""
hmi_modbus_client.py — HMI 网关 Modbus TCP 实时值读取（BACnet 上传值比对用）

HMI 网关对外提供 Modbus TCP 服务器，下挂设备按 Modbus Unit ID 区分，**无需各下挂
设备自身 IP**：读「网关 Modbus 地址:端口 + 该设备 Unit ID」即可拿到该设备实时值。

设备 param_key → 寄存器 的映射复用 Protocols/devices/<设备>.py 的 build_param_map()
（与 BACnet objectName 解析出的 param_key 对齐）。解码逻辑直接复用
tools/Protocols/modbus_reader._decode_registers（单一事实源），不维护并行实现。

典型用法：
    from projects.AcuHMI_1_7.helpers.hmi_modbus_client import read_modbus_values
    values = read_modbus_values("devices.acuvimiir", "192.168.3.51", 502, 1,
                                ["FREQ_Hz", "VLN_a_V"])
    # -> {"FREQ_Hz": (59.98, ""), "VLN_a_V": (120.3, ""), ...}
"""
from __future__ import annotations

import asyncio
import importlib
import logging
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Coroutine, Optional, Sequence, TypeVar

from pymodbus.client import AsyncModbusTcpClient

# tools/Protocols 目录入 path（仅需插入一次）：
#   1. 设备模块（devices/acurev2100.py 等）内部以 `from modbus_reader import ...` 自引用。
#   2. 本模块直接 import modbus_reader._decode_registers（单一解码事实源）。
# 路径：hmi_modbus_client.py -> helpers -> AcuHMI_1_7 -> projects -> autotest(仓库根)
_PROTOCOLS_DIR = str(Path(__file__).parents[3] / "tools" / "Protocols")
if _PROTOCOLS_DIR not in sys.path:
    sys.path.insert(0, _PROTOCOLS_DIR)

# 直接复用协议侧解码函数（单一事实源），避免与 modbus_reader 漂移。
from modbus_reader import _decode_registers as _decode  # noqa: E402

log = logging.getLogger(__name__)

_T = TypeVar("_T")

# 单次 FC03 请求最多读取的寄存器数量（Modbus 协议上限 125）
_MAX_REGS_PER_REQUEST = 120

# 建链重试：部分下挂表 Modbus TCP 连接槽极少（≈1 个）且 TCP 释放慢，
# 槽位被占用（如网关正轮询该表）或刚断开未释放时，connect() 会瞬时拒连
# （client.connected 为 False）。不重试则一次瞬时拒连导致整批参数全判失败，
# 故对建链单独做带退避的重试，每次用全新 client，给表留出释放/腾出连接的时间。
_CONNECT_MAX_RETRIES = 4      # 建链最多额外重试次数（总尝试 = 1 + 该值）
_CONNECT_BACKOFF_STEP = 0.8   # 退避步长（秒），第 n 次重试前等待 n * step


def _run_coro(coro: Coroutine[Any, Any, _T]) -> _T:
    """同步执行协程，兼容 Playwright sync API 的运行中事件循环（独立线程跑全新循环）。"""
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(coro)
    with ThreadPoolExecutor(max_workers=1) as executor:
        return executor.submit(asyncio.run, coro).result()


def build_device_param_map(modbus_module: str) -> dict[str, Any]:
    """加载设备 Modbus 地址映射 param_key → ModbusRegister（每次直接调用，无全局缓存）。"""
    module = importlib.import_module(modbus_module)
    return module.build_param_map()


async def _connect_with_retry(
    host: str,
    port: int,
    max_retries: int = _CONNECT_MAX_RETRIES,
) -> Optional[AsyncModbusTcpClient]:
    """带退避重试地建立 Modbus TCP 连接，成功返回已连接的 client，全部失败返回 None。

    每次尝试都新建 client（避免在失败 client 上复用残留状态），失败后递增退避，
    给连接槽极少、TCP 释放慢的下挂表留出腾出连接的时间。返回的 client 由调用方负责 close。
    """
    for attempt in range(max_retries + 1):
        client = AsyncModbusTcpClient(host=host, port=port)
        try:
            await client.connect()
        except Exception as exc:  # noqa: BLE001 - 建链异常统一按未连接处理后重试
            log.debug("Modbus 建链异常 [%s:%d] 第%d次: %s", host, port, attempt + 1, exc)
        if client.connected:
            if attempt > 0:
                log.info("Modbus 建链成功 [%s:%d]（第%d次尝试）", host, port, attempt + 1)
            return client
        client.close()
        if attempt < max_retries:
            await asyncio.sleep(_CONNECT_BACKOFF_STEP * (attempt + 1))
    return None


def _reg_count(reg: Any) -> int:
    """寄存器占用的寄存器/线圈数量（duck-type，兼容 Protocols.ModbusRegister.count）。"""
    count = getattr(reg, "count", None)
    if count is not None:
        return int(count)
    if reg.dtype == "float64":
        return 4
    if reg.dtype in ("bit", "uint16", "int16"):
        return 1
    return 2


async def _fc03_read(
    client: AsyncModbusTcpClient,
    unit: int,
    address: int,
    count: int,
    timeout: float,
    max_retries: int,
) -> list[int]:
    """执行 FC03 读取，返回寄存器列表；失败抛 IOError。"""
    last_err = ""
    for attempt in range(max_retries + 1):
        try:
            resp = await asyncio.wait_for(
                client.read_holding_registers(address=address, count=count, device_id=unit),
                timeout=timeout,
            )
            if resp.isError():
                last_err = f"FC03 错误响应: {resp}"
                break
            return list(resp.registers)
        except asyncio.TimeoutError:
            last_err = f"超时 (addr=0x{address:04X}, count={count})"
        except Exception as exc:  # noqa: BLE001 - 统一转 IOError 由上层归类
            last_err = str(exc)
        if attempt < max_retries:
            await asyncio.sleep(0.2)
    raise IOError(last_err)


async def _read_bit(
    client: AsyncModbusTcpClient,
    unit: int,
    fc: int,
    address: int,
    timeout: float,
    max_retries: int,
) -> bool:
    """执行 FC01/FC02 单点读取，返回 bool；失败抛 IOError。"""
    last_err = ""
    for attempt in range(max_retries + 1):
        try:
            if fc == 1:
                resp = await asyncio.wait_for(
                    client.read_coils(address=address, count=1, device_id=unit),
                    timeout=timeout,
                )
            else:
                resp = await asyncio.wait_for(
                    client.read_discrete_inputs(address=address, count=1, device_id=unit),
                    timeout=timeout,
                )
            if resp.isError():
                last_err = f"FC{fc:02d} 错误响应: {resp}"
                break
            return bool(resp.bits[0])
        except asyncio.TimeoutError:
            last_err = f"超时 (addr=0x{address:04X})"
        except Exception as exc:  # noqa: BLE001
            last_err = str(exc)
        if attempt < max_retries:
            await asyncio.sleep(0.2)
    raise IOError(last_err)


async def _async_read_values(
    param_map: dict[str, Any],
    host: str,
    port: int,
    unit: int,
    param_keys: Sequence[str],
    timeout: float,
    max_retries: int,
) -> dict[str, tuple[Optional[float], str]]:
    """连接网关 Modbus，批量读取指定参数，返回 {param_key: (value, error)}。"""
    results: dict[str, tuple[Optional[float], str]] = {}

    # 按功能码分组；未知参数直接标错
    fc3_regs: list[Any] = []
    bit_regs: list[Any] = []
    for key in param_keys:
        reg = param_map.get(key)
        if reg is None:
            results[key] = (None, "未在 Modbus 地址表中")
        elif getattr(reg, "fc", 3) in (1, 2):
            bit_regs.append(reg)
        else:
            fc3_regs.append(reg)

    client = await _connect_with_retry(host, port)
    if client is None:
        for key in param_keys:
            results.setdefault(key, (None, f"Modbus 连接失败 {host}:{port}"))
        return results
    try:
        # FC03：按地址排序，连续段合并批量读
        fc3_sorted = sorted(fc3_regs, key=lambda r: r.address)
        batch: list[Any] = []

        async def _flush_fc3() -> None:
            if not batch:
                return
            start = batch[0].address
            end = batch[-1].address + _reg_count(batch[-1])
            try:
                raw = await _fc03_read(client, unit, start, end - start, timeout, max_retries)
                for r in batch:
                    off = r.address - start
                    try:
                        results[r.param_key] = (
                            _decode(raw[off: off + _reg_count(r)], r.dtype) * getattr(r, "scale", 1.0),
                            "",
                        )
                    except Exception as exc:  # noqa: BLE001
                        results[r.param_key] = (None, f"解码失败: {exc}")
            except IOError as exc:
                for r in batch:
                    results[r.param_key] = (None, str(exc))

        for reg in fc3_sorted:
            if not batch:
                batch.append(reg)
                continue
            prev_end = batch[-1].address + _reg_count(batch[-1])
            within = (reg.address + _reg_count(reg) - batch[0].address) <= _MAX_REGS_PER_REQUEST
            if reg.address == prev_end and within:
                batch.append(reg)
            else:
                await _flush_fc3()
                batch = [reg]
        await _flush_fc3()

        # FC01/FC02：逐点读（这些设备位类参数极少）
        for reg in bit_regs:
            try:
                bit = await _read_bit(client, unit, reg.fc, reg.address, timeout, max_retries)
                results[reg.param_key] = (1.0 if bit else 0.0, "")
            except IOError as exc:
                results[reg.param_key] = (None, str(exc))
    except Exception as exc:  # noqa: BLE001
        for key in param_keys:
            results.setdefault(key, (None, f"Modbus 读取异常: {exc}"))
    finally:
        client.close()

    return results


def read_modbus_values(
    modbus_module: str,
    host: str,
    port: int,
    unit: int,
    param_keys: Sequence[str],
    timeout: float = 10.0,
    max_retries: int = 2,
) -> dict[str, tuple[Optional[float], str]]:
    """
    读取网关 Modbus 上指定设备（按 Unit ID）的实时值。

    Args:
        modbus_module: 设备 Modbus 地址映射模块（如 "devices.acuvimiir"）。
        host/port/unit: 网关 Modbus 地址、端口、该设备 Unit ID。
        param_keys:    要读取的 param_key 列表（通常取 BACnet 已发布 ∩ 模板 的交集）。

    Returns:
        {param_key: (value, error)}；value 为 None 时 error 非空说明原因。
        连接失败时所有 key 均返回连接错误。
    """
    try:
        param_map = build_device_param_map(modbus_module)
    except Exception as exc:  # noqa: BLE001
        log.warning("加载 Modbus 地址表失败 [%s]: %s", modbus_module, exc)
        return {k: (None, f"地址表加载失败: {exc}") for k in param_keys}
    return _run_coro(
        _async_read_values(param_map, host, port, unit, param_keys, timeout, max_retries)
    )


def modbus_param_keys(modbus_module: str) -> set[str]:
    """返回设备 Modbus 地址表中的全部 param_key（用于与 BACnet/模板取交集）。"""
    try:
        return set(build_device_param_map(modbus_module))
    except Exception as exc:  # noqa: BLE001
        log.warning("加载 Modbus 地址表失败 [%s]: %s", modbus_module, exc)
        return set()
