"""
验证写入 SERVICE_CONFIG（0x1042）是否会导致 CT Primary 被重置。

使用方法：
    # 默认跑 1 轮
    python test_case/AcuHMI_1_7/wiring_check/check_ct_primary.py

    # 指定循环 3 轮
    python test_case/AcuHMI_1_7/wiring_check/check_ct_primary.py 3

需要先确认：
    CT_PRIMARY_REG  — CT Primary 寄存器地址（见 4100 Modbus 地址表）
    CT_PRIMARY_LEN  — CT Primary 占用寄存器数（uint16=1，uint32=2）
"""
import socket
import struct
import time

# ── 连接参数 ──────────────────────────────────────────────────────────────────
METER_IP   = '192.168.2.242'
METER_PORT = 502
SLAVE_ID   = 245   # 0x66

# ── ！填写 CT Primary 寄存器地址 ───────────────────────────────────────────────
CT_PRIMARY_REG = 0x1048   # TODO：替换为实际地址，例如 0x1009
CT_PRIMARY_LEN = 2         # uint32 占 2 个寄存器；若 uint16 改为 1

# ── 接线方式 ──────────────────────────────────────────────────────────────────
SERVICE_CONFIGS = {
    0: '1E2W',
    1: '2E3W 1Phase',
    2: '2E3W Delta',
    3: '2E3W Network',
    4: '3E4WY',
}
SERVICE_CONFIG_REG = 0x1042

WAIT_SECONDS = 2   # 写完后等待设备处理的时间（秒）

_tid = 0


def _next_tid() -> int:
    global _tid
    _tid = (_tid + 1) & 0xFFFF
    return _tid


def fc06(sock, reg: int, value: int):
    """FC06 写单寄存器"""
    tid = _next_tid()
    pdu = struct.pack('>BBHH', SLAVE_ID, 0x06, reg, value)
    mbap = struct.pack('>HHH', tid, 0, len(pdu))
    sock.sendall(mbap + pdu)
    return sock.recv(256)


def fc03(sock, reg: int, count: int) -> list[int]:
    """FC03 读保持寄存器，返回寄存器值列表"""
    tid = _next_tid()
    pdu = struct.pack('>BBHH', SLAVE_ID, 0x03, reg, count)
    mbap = struct.pack('>HHH', tid, 0, len(pdu))
    sock.sendall(mbap + pdu)
    resp = sock.recv(256)
    # resp: MBAP(6) + UnitID(1) + FC(1) + ByteCount(1) + Data
    data_start = 6 + 1 + 1 + 1
    byte_count = resp[8]
    regs = []
    for i in range(0, byte_count, 2):
        regs.append(struct.unpack('>H', resp[data_start + i: data_start + i + 2])[0])
    return regs


def read_ct_primary(sock) -> int | None:
    regs = fc03(sock, CT_PRIMARY_REG, CT_PRIMARY_LEN)
    if not regs:
        return None
    if CT_PRIMARY_LEN == 2:
        return (regs[0] << 16) | regs[1]   # uint32 big-endian
    return regs[0]                           # uint16


def main(repeat: int = 1):
    with socket.create_connection((METER_IP, METER_PORT), timeout=5) as sock:
        print(f'连接 {METER_IP}:{METER_PORT}  Slave={SLAVE_ID}  循环轮数={repeat}\n')
        print(f'{"轮次":>4} {"接线方式":<20} {"写入前 CT Primary":>18} {"写入后 CT Primary":>18} {"是否变化":>10}')
        print('-' * 80)

        for rnd in range(1, repeat + 1):
            for value, name in SERVICE_CONFIGS.items():
                before = read_ct_primary(sock)
                fc06(sock, SERVICE_CONFIG_REG, value)
                time.sleep(WAIT_SECONDS)
                after = read_ct_primary(sock)
                changed = '*** 变化 ***' if before != after else '无变化'
                print(f'{rnd:>4} {name:<20} {str(before):>18} {str(after):>18} {changed:>10}')
            if rnd < repeat:
                print()

        print('\n测试完成。')


if __name__ == '__main__':
    import sys
    _repeat = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    main(_repeat)
