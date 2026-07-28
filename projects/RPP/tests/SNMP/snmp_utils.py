"""
SNMP 工具函数：通过 SnmpWalk.exe 读取数据并解析
Modbus 工具函数：通过 pymodbus 读取寄存器数据
"""
import os
import subprocess
import struct
import logging
from typing import Optional
from urllib.parse import urlparse

log = logging.getLogger(__name__)

SNMP_WALK_EXE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SnmpWalk", "SnmpWalk.exe")

import sys
from pathlib import Path

# 迁移到 RPP：连接配置改用 projects/RPP/settings.py，不再依赖 projects.AcuHMI_1_7。
_RPP_ROOT = Path(__file__).resolve().parents[2]  # projects/RPP/
if str(_RPP_ROOT) not in sys.path:
    sys.path.insert(0, str(_RPP_ROOT))
from settings import BASE_URL as _BASE_URL

SNMP_HOST = urlparse(_BASE_URL).hostname or "192.168.2.9"
SNMP_COMMUNITY = "123456789012"
SNMP_VERSION = "2c"
SNMP_PORT = 161

# subprocess 总存活时长；4100 上千 OID 需要充裕时间
_WALK_PROCESS_TIMEOUT = 600


def snmp_walk(start_oid: str, stop_oid: Optional[str] = None,
              total_timeout: int = _WALK_PROCESS_TIMEOUT) -> dict:
    """执行 SnmpWalk（v2c），返回 {oid: value} 字典。"""
    cmd = [
        SNMP_WALK_EXE,
        f"-r:{SNMP_HOST}",
        f"-c:{SNMP_COMMUNITY}",
        f"-v:{SNMP_VERSION}",
        f"-p:{SNMP_PORT}",
        f"-os:{start_oid}",
    ]
    if stop_oid:
        cmd.append(f"-op:{stop_oid}")

    cmd_str = " ".join(cmd)
    log.info("[SnmpWalk] %s", cmd_str)
    print(f"\n[SnmpWalk] {cmd_str}")

    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=total_timeout,
        cwd=os.path.dirname(SNMP_WALK_EXE),
    )

    if result.returncode != 0:
        log.warning("[SnmpWalk] returncode=%d  stderr=%s", result.returncode, result.stderr[:200])

    data = _parse_snmp_output(result.stdout)
    log.info("[SnmpWalk] Returned %d OIDs", len(data))
    print(f"[SnmpWalk] Returned {len(data)} OIDs")
    if len(data) == 0:
        log.warning("[SnmpWalk] 0 OIDs，raw stdout前500字符: %s", result.stdout[:500])
        log.warning("[SnmpWalk] stderr前200字符: %s", result.stderr[:200])
    return data


def snmp_walk_device(base_oid: str = ".1.3.6.1.4.1.39604",
                     port: Optional[int] = None,
                     community: Optional[str] = None,
                     total_timeout: int = _WALK_PROCESS_TIMEOUT) -> dict:
    """获取设备全量 SNMP 数据（v2c）。port/community 为 None 时使用模块全局值。"""
    actual_port      = port      if port      is not None else SNMP_PORT
    actual_community = community if community is not None else SNMP_COMMUNITY
    cmd = [
        SNMP_WALK_EXE,
        f"-r:{SNMP_HOST}",
        f"-c:{actual_community}",
        f"-v:{SNMP_VERSION}",
        f"-p:{actual_port}",
        f"-os:{base_oid}",
    ]

    cmd_str = " ".join(cmd)
    log.info("[SnmpWalk] %s", cmd_str)
    print(f"\n[SnmpWalk] {cmd_str}")

    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=total_timeout,
        cwd=os.path.dirname(SNMP_WALK_EXE),
    )

    if result.returncode != 0:
        log.warning("[SnmpWalk] returncode=%d  stderr=%s", result.returncode, result.stderr[:200])
        print(f"[SnmpWalk] returncode={result.returncode}  stderr={result.stderr[:200]}")

    data = _parse_snmp_output(result.stdout)
    log.info("[SnmpWalk] Returned %d OIDs", len(data))
    print(f"[SnmpWalk] Returned {len(data)} OIDs")
    if len(data) == 0:
        log.warning("[SnmpWalk] 0 OIDs，raw stdout前500字符: %s", result.stdout[:500])
        log.warning("[SnmpWalk] stderr前200字符: %s", result.stderr[:200])
    return data


def snmp_walk_device_v3(
    base_oid: str = ".1.3.6.1.4.1.39604",
    port: Optional[int] = None,
    security_name: str = "",
    auth_protocol: str = "MD5",
    auth_password: str = "",
    priv_protocol: str = "DES",
    priv_password: str = "",
    security_level: str = "authPriv",
    total_timeout: int = _WALK_PROCESS_TIMEOUT,
) -> dict:
    """获取设备全量 SNMP 数据（v3 USM）。port 为 None 时使用模块全局值。"""
    actual_port = port if port is not None else SNMP_PORT
    cmd = [
        SNMP_WALK_EXE,
        f"-r:{SNMP_HOST}",
        f"-v:3",
        f"-p:{actual_port}",
        f"-os:{base_oid}",
        f"-sn:{security_name}",
        f"-sl:{security_level}",
    ]
    if security_level in ("authNoPriv", "authPriv") and auth_password:
        cmd += [f"-ap:{auth_protocol}", f"-aw:{auth_password}"]
    if security_level == "authPriv" and priv_password:
        _priv = {"AES": "AES128", "AES-128": "AES128", "AES-192": "AES192", "AES-256": "AES256"}.get(
            priv_protocol.upper(), priv_protocol
        )
        cmd += [f"-pp:{_priv}", f"-pw:{priv_password}"]

    cmd_str = " ".join(cmd)
    log.info("[SnmpWalk v3] %s", cmd_str)
    print(f"\n[SnmpWalk v3] {cmd_str}")

    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=total_timeout,
        cwd=os.path.dirname(SNMP_WALK_EXE),
    )

    if result.returncode != 0:
        log.warning("[SnmpWalk v3] returncode=%d  stderr=%s", result.returncode, result.stderr[:200])
        print(f"[SnmpWalk v3] returncode={result.returncode}  stderr={result.stderr[:200]}")

    data = _parse_snmp_output(result.stdout)
    log.info("[SnmpWalk v3] Returned %d OIDs", len(data))
    print(f"[SnmpWalk v3] Returned {len(data)} OIDs")
    if len(data) == 0:
        log.warning("[SnmpWalk v3] 0 OIDs，raw stdout前500字符: %s", result.stdout[:500])
        log.warning("[SnmpWalk v3] stderr前200字符: %s", result.stderr[:200])
    return data


def _parse_snmp_output(output: str) -> dict:
    """解析 SnmpWalk 输出，返回 {oid: value} 字典。"""
    data = {}
    for line in output.splitlines():
        line = line.strip()
        if not line.startswith("OID="):
            continue
        parts = line.split(", ")
        if len(parts) < 3:
            continue
        oid = parts[0].replace("OID=", "").strip()
        value_part = parts[2].replace("Value=", "").strip()
        data[oid] = value_part
    return data


def validate_walk_completeness(data: dict, min_oid_count: int, label: str = "") -> None:
    """断言 walk 结果 OID 数量不少于预期最小值，捕获因进程超时导致的静默截断。"""
    actual = len(data)
    tag = f"({label})" if label else ""
    assert actual >= min_oid_count, (
        f"SNMP walk{tag} 数据不完整: 返回 {actual} 个 OID，期望至少 {min_oid_count} 个。"
        f"可能是 walk 进程被超时截断，请检查 total_timeout 配置。"
    )


def get_snmp_value(oid: str, all_data: dict) -> Optional[float]:
    """从已有 SNMP 数据中提取指定 OID 的浮点值。"""
    raw = all_data.get(oid)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


# ─── Modbus 工具 ──────────────────────────────────────────────────────────────

from pymodbus.client import ModbusTcpClient


def read_modbus_float(host: str, port: int, unit_id: int, start_addr: int) -> Optional[float]:
    """读取单个 float 值 (2 个保持寄存器, ABCD 大端序)。"""
    client = ModbusTcpClient(host, port=port)
    if not client.connect():
        raise ConnectionError(f"无法连接 Modbus {host}:{port}")
    try:
        rr = client.read_holding_registers(start_addr, count=2, device_id=unit_id)
        if rr.isError():
            return None
        raw = struct.pack(">HH", rr.registers[0], rr.registers[1])
        return struct.unpack(">f", raw)[0]
    finally:
        client.close()


def read_modbus_registers(host: str, port: int, unit_id: int,
                          start_addr: int, count: int) -> Optional[list]:
    """批量读取寄存器，返回原始寄存器列表。"""
    client = ModbusTcpClient(host, port=port)
    if not client.connect():
        raise ConnectionError(f"无法连接 Modbus {host}:{port}")
    try:
        rr = client.read_holding_registers(start_addr, count=count, device_id=unit_id)
        if rr.isError():
            return None
        return rr.registers
    finally:
        client.close()


def _reconnect(host: str, port: int) -> ModbusTcpClient:
    """创建并连接一个新的 ModbusTcpClient，供批量读取自动重连使用。"""
    c = ModbusTcpClient(host, port=port, timeout=10)
    c.connect()
    return c


def batch_read_modbus_floats(host: str, port: int, unit_id: int,
                             addr_list: list) -> dict:
    """
    一次 TCP 连接批量读取多个 float 寄存器（每个 2 registers，ABCD 大端序）。
    返回 {addr: value}，读取失败的地址值为 None。
    """
    result = {}
    log.info("[Modbus] 批量读取 %d 个地址，目标 %s:%d unit=%d", len(addr_list), host, port, unit_id)
    print(f"[Modbus] 批量读取 {len(addr_list)} 个地址，目标 {host}:{port} unit={unit_id}")

    client = _reconnect(host, port)
    if not client.is_socket_open():
        log.warning("[Modbus] 连接失败 %s:%d", host, port)
        return {addr: None for addr in addr_list}
    try:
        for addr in addr_list:
            try:
                if not client.is_socket_open():
                    client = _reconnect(host, port)
                rr = client.read_holding_registers(addr, count=2, device_id=unit_id)
                if rr.isError():
                    result[addr] = None
                else:
                    result[addr] = decode_float_abcd(rr.registers, 0)
            except Exception as e:
                log.debug("[Modbus] addr=%d 读取异常: %s", addr, e)
                result[addr] = None
                try:
                    client.close()
                except Exception:
                    pass
                client = _reconnect(host, port)
    finally:
        client.close()

    valid = sum(1 for v in result.values() if v is not None)
    log.info("[Modbus] 读取完成: %d/%d 成功", valid, len(addr_list))
    print(f"[Modbus] 读取完成: {valid}/{len(addr_list)} 成功")
    return result


def decode_float_abcd(regs: list, offset: int = 0) -> float:
    """从寄存器列表按 ABCD 大端序解码 float（2 寄存器）。"""
    raw = struct.pack(">HH", regs[offset], regs[offset + 1])
    return struct.unpack(">f", raw)[0]


def decode_double_abcd(regs: list, offset: int = 0) -> float:
    """从寄存器列表按 ABCDEFGH 大端序解码 double（4 寄存器）。"""
    raw = struct.pack(">HHHH", regs[offset], regs[offset + 1],
                      regs[offset + 2], regs[offset + 3])
    return struct.unpack(">d", raw)[0]


def decode_uint32_abcd(regs: list, offset: int = 0) -> int:
    """从寄存器列表按 ABCD 大端序解码 uint32。"""
    raw = struct.pack(">HH", regs[offset], regs[offset + 1])
    return struct.unpack(">I", raw)[0]


def decode_int32_abcd(regs: list, offset: int = 0) -> int:
    """从寄存器列表按 ABCD 大端序解码有符号 int32。"""
    raw = struct.pack(">HH", regs[offset], regs[offset + 1])
    return struct.unpack(">i", raw)[0]


def batch_read_modbus_uint32s(host: str, port: int, unit_id: int,
                              addr_list: list) -> dict:
    """
    批量读取 uint32 保持寄存器（function code 03H，每个 2 registers，ABCD 大端序）。
    返回 {addr: int_value}。
    """
    result = {}
    client = _reconnect(host, port)
    if not client.is_socket_open():
        return {addr: None for addr in addr_list}
    try:
        for addr in addr_list:
            try:
                if not client.is_socket_open():
                    client = _reconnect(host, port)
                rr = client.read_holding_registers(addr, count=2, device_id=unit_id)
                if rr.isError():
                    result[addr] = None
                else:
                    result[addr] = decode_uint32_abcd(rr.registers, 0)
            except Exception as e:
                log.debug("[Modbus] addr=%d uint32 读取异常: %s", addr, e)
                result[addr] = None
                try:
                    client.close()
                except Exception:
                    pass
                client = _reconnect(host, port)
    finally:
        client.close()
    return result


def batch_read_modbus_int32s(host: str, port: int, unit_id: int,
                             addr_list: list) -> dict:
    """
    批量读取有符号 int32 保持寄存器（function code 03H，每个 2 registers，ABCD 大端序）。
    返回 {addr: int_value}。用于可能为负值的参数（如净电能 EP_NET/EQ_NET）。
    """
    result = {}
    client = _reconnect(host, port)
    if not client.is_socket_open():
        return {addr: None for addr in addr_list}
    try:
        for addr in addr_list:
            try:
                if not client.is_socket_open():
                    client = _reconnect(host, port)
                rr = client.read_holding_registers(addr, count=2, device_id=unit_id)
                if rr.isError():
                    result[addr] = None
                else:
                    result[addr] = decode_int32_abcd(rr.registers, 0)
            except Exception as e:
                log.debug("[Modbus] addr=%d int32 读取异常: %s", addr, e)
                result[addr] = None
                try:
                    client.close()
                except Exception:
                    pass
                client = _reconnect(host, port)
    finally:
        client.close()
    return result


def batch_read_modbus_discrete_inputs(host: str, port: int, unit_id: int,
                                      addr_list: list) -> dict:
    """
    批量读取离散输入（function code 02H）。
    返回 {addr: 0/1}。
    """
    result = {}
    if not addr_list:
        return result
    client = ModbusTcpClient(host, port=port)
    if not client.connect():
        return {addr: None for addr in addr_list}
    try:
        # 读取覆盖所有地址的连续范围
        min_addr, max_addr = min(addr_list), max(addr_list)
        count = max_addr - min_addr + 1
        rr = client.read_discrete_inputs(min_addr, count=count, device_id=unit_id)
        if rr.isError():
            return {addr: None for addr in addr_list}
        for addr in addr_list:
            bit = rr.bits[addr - min_addr]
            result[addr] = 1 if bit else 0
    except Exception as e:
        log.debug("[Modbus] discrete_inputs 读取异常: %s", e)
        result = {addr: None for addr in addr_list}
    finally:
        client.close()
    return result


def batch_read_modbus_coils(host: str, port: int, unit_id: int,
                             addr_list: list) -> dict:
    """
    批量读取线圈（function code 01H）。
    返回 {addr: 0/1}。地址可不连续（DO 0-7 和 RO 32-33 分两次读）。
    """
    result = {}
    if not addr_list:
        return result
    client = ModbusTcpClient(host, port=port)
    if not client.connect():
        return {addr: None for addr in addr_list}
    try:
        # 按连续区间分组，分别读取
        sorted_addrs = sorted(addr_list)
        groups = []
        cur_start = sorted_addrs[0]
        cur_end   = sorted_addrs[0]
        for addr in sorted_addrs[1:]:
            if addr == cur_end + 1:
                cur_end = addr
            else:
                groups.append((cur_start, cur_end))
                cur_start = cur_end = addr
        groups.append((cur_start, cur_end))

        for g_start, g_end in groups:
            count = g_end - g_start + 1
            rr = client.read_coils(g_start, count=count, device_id=unit_id)
            if rr.isError():
                for a in range(g_start, g_end + 1):
                    if a in addr_list:
                        result[a] = None
            else:
                for a in range(g_start, g_end + 1):
                    if a in addr_list:
                        result[a] = 1 if rr.bits[a - g_start] else 0
    except Exception as e:
        log.debug("[Modbus] coils 读取异常: %s", e)
        result = {addr: None for addr in addr_list}
    finally:
        client.close()
    return result


def batch_read_modbus_words(host: str, port: int, unit_id: int,
                             addr_list: list) -> dict:
    """
    批量读取 uint16 保持寄存器（function code 03H，每个 1 register）。
    返回 {addr: int_value}，用于 PQ 电能质量参数。
    """
    result = {}
    client = _reconnect(host, port)
    if not client.is_socket_open():
        log.warning("[Modbus] 连接失败 %s:%d", host, port)
        return {addr: None for addr in addr_list}
    try:
        for addr in addr_list:
            try:
                if not client.is_socket_open():
                    client = _reconnect(host, port)
                rr = client.read_holding_registers(addr, count=1, device_id=unit_id)
                if rr.isError():
                    result[addr] = None
                else:
                    result[addr] = rr.registers[0]
            except Exception as e:
                log.debug("[Modbus] addr=%d word 读取异常: %s", addr, e)
                result[addr] = None
                try:
                    client.close()
                except Exception:
                    pass
                client = _reconnect(host, port)
    finally:
        client.close()
    return result


def batch_read_modbus_doubles(host: str, port: int, unit_id: int,
                              addr_list: list) -> dict:
    """
    一次 TCP 连接批量读取多个 double 寄存器（每个 4 registers，ABCDEFGH 大端序）。
    返回 {addr: value}，读取失败的地址值为 None。
    """
    result = {}
    client = _reconnect(host, port)
    if not client.is_socket_open():
        log.warning("[Modbus] 连接失败 %s:%d", host, port)
        return {addr: None for addr in addr_list}
    try:
        for addr in addr_list:
            try:
                if not client.is_socket_open():
                    client = _reconnect(host, port)
                rr = client.read_holding_registers(addr, count=4, device_id=unit_id)
                if rr.isError():
                    result[addr] = None
                else:
                    result[addr] = decode_double_abcd(rr.registers, 0)
            except Exception as e:
                log.debug("[Modbus] addr=%d double 读取异常: %s", addr, e)
                result[addr] = None
                try:
                    client.close()
                except Exception:
                    pass
                client = _reconnect(host, port)
    finally:
        client.close()
    return result


def read_serial_number(host: str, port: int, unit_id: int) -> str:
    """读取设备序列号 (Modbus 0xF040, 16个寄存器, ASCII)。"""
    regs = read_modbus_registers(host, port, unit_id, 0xF040, 16)
    if regs is None:
        return "N/A"
    chars = []
    for r in regs:
        high = (r >> 8) & 0xFF
        low = r & 0xFF
        if high and high != 0xFF:
            chars.append(chr(high))
        if low and low != 0xFF:
            chars.append(chr(low))
    return "".join(chars).strip("\x00").strip()
