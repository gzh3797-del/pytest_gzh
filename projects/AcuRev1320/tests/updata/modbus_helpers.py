"""Modbus 回读辅助（updata 用例集共用）。

用于「升级前后配置/数据不变」类用例（023_02 case7_1 / case7_2）：升级前后各做一次
寄存器快照，比对是否被升级流程意外改写。

连接参数取自分层配置 `configs/global.yaml`（经 modbus_config），TCP 用 QT_tcp，
RTU 用 rtu 段。
"""
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(_HERE))))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from modbus_config import modbus_config  # noqa: E402

try:
    from pymodbus.client import ModbusTcpClient, ModbusSerialClient
except ImportError:  # pragma: no cover
    ModbusTcpClient = ModbusSerialClient = None

# 要做「升级前后不变」比对的配置寄存器块；按 (起始地址, 寄存器数) 列出。
# 默认空 —— 真机使用前请按 AcuRev-1320 Modbus 地址表填入 basic setting / 配置类寄存器，
# 块为空时相关用例会自动 skip（避免拿不到地址而误判）。
CONFIG_REGISTER_BLOCKS = [
    # (0x0100, 20),
]


def _make_client():
    if ModbusTcpClient is None:
        raise RuntimeError('未安装 pymodbus，无法做 Modbus 回读（pip install pymodbus）')
    mode = str(modbus_config.get('conn_mode', 'tcp')).lower()
    if mode == 'rtu':
        rtu = modbus_config['rtu']
        client = ModbusSerialClient(port=rtu['port'], baudrate=int(rtu['baudrate']),
                                    parity=rtu.get('parity', 'N'), timeout=5)
        return client, int(rtu.get('slaveid', 1))
    tcp = modbus_config.get('QT_tcp') or modbus_config['tcp']
    host = tcp.get('host') or tcp.get('ip')
    client = ModbusTcpClient(host=host, port=int(tcp.get('port', 502)), timeout=5)
    return client, int(tcp.get('slave_id') or tcp.get('slaveid') or 1)


def make_tcp_client():
    """强制走 TCP 客户端（用于「TCP 通信并发」类用例），不受 conn_mode 影响。"""
    if ModbusTcpClient is None:
        raise RuntimeError('未安装 pymodbus（pip install pymodbus）')
    tcp = modbus_config.get('QT_tcp') or modbus_config['tcp']
    host = tcp.get('host') or tcp.get('ip')
    client = ModbusTcpClient(host=host, port=int(tcp.get('port', 502)), timeout=5)
    return client, int(tcp.get('slave_id') or tcp.get('slaveid') or 1)


def snapshot_config_registers():
    """读取 CONFIG_REGISTER_BLOCKS 指定的全部寄存器，返回 {addr: value} 快照。

    块列表为空时返回 None（调用方据此 skip）。
    """
    if not CONFIG_REGISTER_BLOCKS:
        return None
    client, unit = _make_client()
    if not client.connect():
        raise RuntimeError('Modbus 连接失败，无法读取寄存器快照')
    snap = {}
    try:
        for addr, count in CONFIG_REGISTER_BLOCKS:
            rr = client.read_holding_registers(address=addr, count=count, slave=unit)
            if rr.isError():
                raise RuntimeError(f'读寄存器失败 addr={addr} count={count}: {rr}')
            for i, val in enumerate(rr.registers):
                snap[addr + i] = val
    finally:
        client.close()
    return snap


def diff_snapshots(before, after):
    """比对两份快照，返回发生变化的 {addr: (before, after)}（空 dict = 完全一致）。"""
    changed = {}
    for addr, b in (before or {}).items():
        a = (after or {}).get(addr)
        if a != b:
            changed[addr] = (b, a)
    return changed
