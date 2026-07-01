"""
CL3021 AC Source 串口通信验证脚本
参数: COM18, 9600, 8N1 (parity=0, dataBit=8, stopBit=0=1 stop bit)
顺序: 联机 -> 切换AC版面 -> 设定线制/范围 -> 控制量程 -> 读取测量
"""

import serial
import time
import struct
import sys


# ─────────────────────────────── 协议工具 ────────────────────────────────

def calc_cs(data: bytes | bytearray) -> int:
    """XOR checksum: bytes[1] ^ bytes[2] ^ ... (跳过 bytes[0]=0x81 帧头)"""
    cs = 0
    for b in data[1:]:
        cs ^= b
    return cs


def encode_fixed_le(value: float, is_current: bool) -> bytes:
    """固定点数编码: V*10000 或 I*1000000, 小端 int32"""
    scale = 1_000_000.0 if is_current else 10_000.0
    scaled = round(value * scale)
    return struct.pack('<i', scaled)


# ─────────────────────────────── 帧构建 ──────────────────────────────────

kHead = 0x81
kRxId = 0x01
kTxId = 0x25


def build_connect() -> bytes:
    """联机命令"""
    f = bytearray([kHead, kRxId, kTxId, 0x06, 0xc9])
    f.append(calc_cs(f))
    return bytes(f)


def build_ac_screen() -> bytes:
    """切换到 AC 输出版面"""
    f = bytearray([kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x10, 0x80, 0x01])
    f.append(calc_cs(f))
    return bytes(f)


def build_set_line() -> bytes:
    """设定线制/输出范围 (单相)"""
    f = bytearray([kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x01, 0x20, 0x08])
    f.append(calc_cs(f))
    return bytes(f)


def build_exit_control() -> bytes:
    """退出 AC 版面 (CS 0x1d 已硬编码验证正确)"""
    return bytes([kHead, kRxId, kTxId, 0x0a, 0xa3, 0x00, 0x10, 0x80, 0x00, 0x1d])


def build_read_command() -> bytes:
    """读取输出测量值"""
    return bytes([
        0x81, 0x01, 0x25, 0x0F, 0xA0, 0x02,
        0x7F, 0xFF, 0x80, 0x3F, 0xFF, 0xFF,
        0x0F, 0x80, 0x39
    ])


def build_angle_update(
    ua_angle=0.0, ub_angle=240.0, uc_angle=120.0,
    ia_angle=0.0, ib_angle=240.0, ic_angle=120.0,
    freq=50.0,
) -> bytes:
    """
    72B 角度帧 (cmd A3 05 46 3F) — 只传相角，振幅字段固定为 0
    实际输出幅值由 build_amplitude_update (41B) 单独控制
    若此帧也写入振幅，设备将两帧叠加导致输出翻倍
    """
    out = bytearray([kHead, kRxId, kTxId, 0x48, 0xa3, 0x05, 0x46, 0x3f])

    # 电压角度: C/B/A 顺序
    out += encode_fixed_le(uc_angle, False)
    out += encode_fixed_le(ub_angle, False)
    out += encode_fixed_le(ua_angle, False)
    # 电流角度: C/B/A 顺序
    out += encode_fixed_le(ic_angle, False)
    out += encode_fixed_le(ib_angle, False)
    out += encode_fixed_le(ia_angle, False)

    out += bytes([0xff])

    # 振幅字段全 0（不控制幅值）
    out += bytes([0x00, 0x00, 0x00, 0x00, 0xfc]) * 3   # 电压
    out += bytes([0x00, 0x00, 0x00, 0x00, 0xfa]) * 3   # 电流

    out += encode_fixed_le(freq, False)
    out += bytes([0x07, 0x03, 0x3f, 0x3f])
    out.append(calc_cs(out))
    return bytes(out)


def build_amplitude_update(
    ua=100.0, ub=100.0, uc=100.0,
    ia=1.0,  ib=1.0,  ic=1.0,
) -> bytes:
    """41B 幅值实时更新帧 (cmd A3 05 44 3F)"""
    out = bytearray([kHead, kRxId, kTxId, 0x29, 0xa3, 0x05, 0x44, 0x3f])
    out += encode_fixed_le(ua, False); out += bytes([0xfc])
    out += encode_fixed_le(ub, False); out += bytes([0xfc])
    out += encode_fixed_le(uc, False); out += bytes([0xfc])
    out += encode_fixed_le(ia, True);  out += bytes([0xfa])
    out += encode_fixed_le(ib, True);  out += bytes([0xfa])
    out += encode_fixed_le(ic, True);  out += bytes([0xfa])
    out += bytes([0x02, 0x3f])
    out.append(calc_cs(out))
    return bytes(out)


def build_freq_update(freq=50.0) -> bytes:
    """14B 频率更新帧 (cmd A3 05 04 C0)"""
    out = bytearray([kHead, kRxId, kTxId, 0x0e, 0xa3, 0x05, 0x04, 0xc0])
    out += encode_fixed_le(freq, False)
    out += bytes([0x07])
    out.append(calc_cs(out))
    return bytes(out)


# ─────────────────────────────── 响应判断 ────────────────────────────────

def looks_like_ok(rx: bytes) -> bool:
    if len(rx) > 4 and rx[4] == 0x30:
        return True
    if len(rx) > 15 and rx[15] == 0x30:
        return True
    if len(rx) > 177 and rx[4] == 0x50:
        return True
    return False


# ─────────────────────────────── 通信函数 ────────────────────────────────

def transact(ser: serial.Serial, frame: bytes, settle_ms=100, extra_wait_ms=200) -> bytes:
    """发送 -> 等待 extra_wait_ms -> 等队列稳定 -> 读取"""
    ser.reset_input_buffer()
    print(f"  TX [{len(frame):3d}B]: {frame.hex(' ').upper()}")
    ser.write(frame)

    time.sleep(extra_wait_ms / 1000.0)

    # 等待 RX 队列稳定
    prev = -1
    for _ in range(20):
        time.sleep(settle_ms / 1000.0)
        curr = ser.in_waiting
        if curr == prev and curr > 0:
            break
        prev = curr

    rx = ser.read(ser.in_waiting) if ser.in_waiting > 0 else b''
    if rx:
        print(f"  RX [{len(rx):3d}B]: {rx.hex(' ').upper()}")
    else:
        print(f"  RX: (无响应)")
    return rx


def decode_measurements(rx: bytes):
    """解析 readOutputs 返回的测量值"""
    min_len = 8 + 5 * 6 + 4
    if len(rx) < min_len:
        print(f"  数据长度不足 ({len(rx)} < {min_len}), 无法解析")
        return
    offset = 8
    try:
        uc_v = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        ub_v = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        ua_v = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        ic_a = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        ib_a = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        ia_a = struct.unpack_from('<i', rx, offset)[0] / 1e6; offset += 5
        freq  = struct.unpack_from('<i', rx, offset)[0] / 10000.0
        print(f"  UA={ua_v:.4f}V  UB={ub_v:.4f}V  UC={uc_v:.4f}V")
        print(f"  IA={ia_a:.6f}A  IB={ib_a:.6f}A  IC={ic_a:.6f}A")
        print(f"  Freq={freq:.4f} Hz")
    except struct.error as e:
        print(f"  解析错误: {e}")


# ─────────────────────────────── 主测试 ──────────────────────────────────

def main():
    port  = 'COM18'
    baud  = 9600
    config = dict(baudrate=baud, bytesize=8, parity='N', stopbits=1,
                  timeout=1.0, write_timeout=2.0, xonxoff=False, rtscts=False, dsrdtr=False)

    print("=" * 60)
    print(f"CL3021 串口测试  {port}  {baud}-8N1")
    print("=" * 60)

    try:
        ser = serial.Serial(port, **config)
        ser.setRTS(True)
        ser.setDTR(True)
    except serial.SerialException as e:
        print(f"\n无法打开 {port}: {e}")
        sys.exit(1)

    with ser:
        ser.reset_input_buffer()
        ser.reset_output_buffer()

        # ── 步骤 1: 联机 ───────────────────────────────────────────────
        print("\n─── 步骤1: 联机命令 ───")
        rx1 = transact(ser, build_connect(), settle_ms=200, extra_wait_ms=500)
        ok1 = looks_like_ok(rx1)
        print(f"  → {'SUCCESS' if ok1 else 'FAIL'}")
        if not ok1:
            print("  联机失败, 停止")
            return
        time.sleep(0.1)

        # ── 步骤 2: 切换 AC 输出版面 ──────────────────────────────────
        print("\n─── 步骤2: 切换 AC 输出版面 ───")
        rx_ac = transact(ser, build_ac_screen(), settle_ms=100, extra_wait_ms=200)
        print(f"  → {'OK' if looks_like_ok(rx_ac) else '无OK响应 (继续)'}")
        time.sleep(0.2)

        # ── 步骤 3: 设定线制/输出范围 ─────────────────────────────────
        print("\n─── 步骤3: 设定线制/输出范围 ───")
        rx_sl = transact(ser, build_set_line(), settle_ms=100, extra_wait_ms=200)
        print(f"  → {'OK' if looks_like_ok(rx_sl) else '无OK响应 (继续)'}")
        time.sleep(0.1)

        # ── 步骤 4a: 角度帧 (72B，振幅全0) ──────────────────────────
        print("\n─── 步骤4a: 角度更新 (72B，振幅=0) ───")
        sp72 = build_angle_update(ua_angle=0.0, ub_angle=240.0, uc_angle=120.0,
                                  ia_angle=0.0, ib_angle=240.0, ic_angle=120.0,
                                  freq=50.0)
        assert len(sp72) == 72, f"expected 72B got {len(sp72)}"
        rx_sp = transact(ser, sp72, settle_ms=100, extra_wait_ms=200)
        print(f"  → {'SUCCESS' if looks_like_ok(rx_sp) else 'FAIL'}")
        time.sleep(0.2)

        # ── 步骤 4b: 幅值实时更新 (41B) ─────────────────────────────
        print("\n─── 步骤4b: 幅值实时更新 (41B) ───")
        sp41 = build_amplitude_update(ua=100.0, ub=100.0, uc=100.0,
                                      ia=1.0,   ib=1.0,   ic=1.0)
        assert len(sp41) == 41, f"expected 41B got {len(sp41)}"
        rx_sp = transact(ser, sp41, settle_ms=100, extra_wait_ms=200)
        print(f"  → {'SUCCESS' if looks_like_ok(rx_sp) else 'FAIL'}")
        time.sleep(0.2)

        # ── 步骤 4c: 频率更新 (14B) ──────────────────────────────────
        print("\n─── 步骤4c: 频率更新 (14B) ───")
        sp14 = build_freq_update(50.0)
        assert len(sp14) == 14, f"expected 14B got {len(sp14)}"
        rx_sp = transact(ser, sp14, settle_ms=100, extra_wait_ms=200)
        print(f"  → {'SUCCESS' if looks_like_ok(rx_sp) else 'FAIL'}")

        # ── 步骤 5: 读取测量值 ────────────────────────────────────────
        print("\n─── 步骤5: 读取测量值 ───")
        rx_rd = transact(ser, build_read_command(), settle_ms=200, extra_wait_ms=500)
        ok_rd = looks_like_ok(rx_rd)
        print(f"  → {'SUCCESS' if ok_rd else 'FAIL'}")
        if ok_rd:
            decode_measurements(rx_rd)

        # ── 退出 ──────────────────────────────────────────────────────
        print("\n─── 退出 AC 版面 ───")
        transact(ser, build_exit_control(), settle_ms=100, extra_wait_ms=100)

    print("\n测试完成")


if __name__ == '__main__':
    main()
