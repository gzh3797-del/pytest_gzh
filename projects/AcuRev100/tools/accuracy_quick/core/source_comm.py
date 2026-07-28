"""
源控制模块（CL3021 标准源）
支持串口和 TCP/UDP 两种传输方式，使用相同的 CL3021 二进制协议

串口协议基于实际抓包验证（2026-06-03）：
  - 初始化: 联机(6B) → AC版面(10B) → 设定线制(10B)
  - 控制:   角度帧(72B) + 幅值帧(41B) + 频率帧(14B)，发两轮确保生效
  - 校验:   CS = XOR(byte[1] … byte[n-2])
"""
import json
import os
import socket
import struct
import serial
import logging
import time
from pathlib import Path

log = logging.getLogger(__name__)

# 帧级追踪(诊断用): 置 SRC_TRACE=1 后每发一帧打印帧名+时刻, 用于把"某个动作掉电"缩小到
# 具体哪一帧(2026-07-28: 用户目击批跑"中间掉两下", 而动作级探针只能定位到动作)。
TRACE_FRAMES = bool(os.environ.get("SRC_TRACE"))


# ══════════════════════════════════════════════════════════════
# 编码工具
# ══════════════════════════════════════════════════════════════

def _cs(data: bytes | bytearray) -> int:
    """CL3021 校验: XOR of byte[1]..byte[n-2]，跳过帧头和CS本身"""
    v = 0
    for b in data[1:]:
        v ^= b
    return v


def _enc_v(val: float) -> bytes:
    """电压/角度/频率: value × 10000 → LE int32"""
    return struct.pack('<i', round(val * 10000))


def _enc_i(val: float) -> bytes:
    """电流: value × 1000000 → LE int32"""
    return struct.pack('<i', round(val * 1000000))


# ══════════════════════════════════════════════════════════════
# 帧构建（已通过实际设备验证）
# ══════════════════════════════════════════════════════════════

def _build_connect() -> bytearray:
    f = bytearray([0x81, 0x01, 0x25, 0x06, 0xC9])
    f.append(_cs(f))
    return f


def _build_switch_screen_cmd(inter: int) -> bytearray:
    """inter=0x01 进入AC版面，inter=0x00 退出"""
    f = bytearray([0x81, 0x01, 0x25, 0x0A, 0xA3, 0x00, 0x10, 0x80, inter & 0xFF])
    f.append(_cs(f))
    return f


def _build_set_line() -> bytearray:
    f = bytearray([0x81, 0x01, 0x25, 0x0A, 0xA3, 0x00, 0x01, 0x20, 0x08])
    f.append(_cs(f))
    return f


def _build_angle_update(quc, qub, qua, qic, qib, qia, freq) -> bytearray:
    """
    72B 角度帧 (cmd A3 05 46 3F)
    只传相角，振幅字段固定为 0（与 TestBench 抓包完全一致）。
    实际输出幅值由 41B _build_amplitude_update 帧独立控制。
    若 72B 也写入振幅，设备会将两帧叠加导致输出翻倍。
    """
    f = bytearray([0x81, 0x01, 0x25, 0x48, 0xA3, 0x05, 0x46, 0x3F])
    # 电压角度 C/B/A
    f += _enc_v(quc) + _enc_v(qub) + _enc_v(qua)
    # 电流角度 C/B/A
    f += _enc_v(qic) + _enc_v(qib) + _enc_v(qia)
    f += bytes([0xFF])
    # 电压振幅字段：全 0（不控制幅值，由 41B 帧负责）
    f += bytes([0x00, 0x00, 0x00, 0x00, 0xFC]) * 3
    # 电流振幅字段：全 0
    f += bytes([0x00, 0x00, 0x00, 0x00, 0xFA]) * 3
    # 频率 + 尾部
    f += _enc_v(freq) + bytes([0x07, 0x03, 0x3F, 0x3F])
    f.append(_cs(f))
    assert len(f) == 72, f"角度帧长度错误: {len(f)}"
    return f


def _build_amplitude_update(uc, ub, ua, ic, ib, ia) -> bytearray:
    """
    41B 幅值实时更新帧 (cmd A3 05 44 3F)
    实际控制电压电流输出值的主命令
    """
    f = bytearray([0x81, 0x01, 0x25, 0x29, 0xA3, 0x05, 0x44, 0x3F])
    f += _enc_v(uc) + bytes([0xFC])
    f += _enc_v(ub) + bytes([0xFC])
    f += _enc_v(ua) + bytes([0xFC])
    f += _enc_i(ic) + bytes([0xFA])
    f += _enc_i(ib) + bytes([0xFA])
    f += _enc_i(ia) + bytes([0xFA])
    f += bytes([0x02, 0x3F])
    f.append(_cs(f))
    assert len(f) == 41, f"幅值帧长度错误: {len(f)}"
    return f


def _build_freq_update(freq: float) -> bytearray:
    """14B 频率更新帧 (cmd A3 05 04 C0)"""
    f = bytearray([0x81, 0x01, 0x25, 0x0E, 0xA3, 0x05, 0x04, 0xC0])
    f += _enc_v(freq) + bytes([0x07])
    f.append(_cs(f))
    assert len(f) == 14, f"频率帧长度错误: {len(f)}"
    return f


def _build_gear_mode(mode_bin: str = '00000000') -> bytearray:
    try:
        mode_val = int(mode_bin, 2)
    except ValueError:
        mode_val = 0
    cmd = [0x81, 0x01, 0x25, 0x0A, 0xA3, 0x05, 0x40, 0x04, mode_val]
    xor = 0
    for b in cmd[1:-1]:
        xor ^= b
    cmd.append(xor)
    return bytearray(cmd)


_V_RANGES = ((30.0, 5), (60.0, 4), (120.0, 3), (240.0, 2), (480.0, 1))
# 🔴 选档余量(2026-07-27 实机): 把 120V 打在 ≤120V 档(正好满刻度)时, 三相带载会让源报
#    "Ub/Uc 过载"; 钉死 480V 档跑同样点位一直正常 ⇒ 不要在量程满刻度附近输出。
#    取 80%: 本矩阵的测点档位与 90% 时相同(120V→240V档, 220/240/277V→480V档),
#    只把保活点 100V 也推到 240V 档(83%→42% 量程), 换更大安全余量。
#    仍远好于全程钉 480V 档(120V 只占 25% 量程 → 源自身误差被放大到 0.2% 量级)。
GEAR_HEADROOM = 0.8


def _get_voltage_gear(v: float) -> int:
    """按幅值选电压档，且留 GEAR_HEADROOM 余量；超最大量程返回最大档(1)。"""
    need = abs(v) / GEAR_HEADROOM
    for top, gear in _V_RANGES:
        if need <= top:
            return gear
    return 1


def _get_current_gear(i: float) -> int:
    if i <= 0.01: return 12
    if i <= 0.02: return 11
    if i <= 0.05: return 10
    if i <= 0.1:  return 9
    if i <= 0.2:  return 8
    if i <= 0.5:  return 7
    if i <= 1:    return 6
    if i <= 2:    return 5
    if i <= 5:    return 4
    if i <= 10:   return 3
    if i <= 20:   return 2
    if i <= 50:   return 1
    return 0


def _build_voltage_gear(uc, ub, ua) -> bytearray:
    cmd = [0x81, 0x01, 0x25, 0x0C, 0xA3, 0x02, 0x02, 0x07,
           _get_voltage_gear(uc), _get_voltage_gear(ub), _get_voltage_gear(ua)]
    xor = 0
    for b in cmd[1:]: xor ^= b
    cmd.append(xor)
    return bytearray(cmd)


def _build_current_gear(ic, ib, ia) -> bytearray:
    cmd = [0x81, 0x01, 0x25, 0x0C, 0xA3, 0x02, 0x02, 0x38,
           _get_current_gear(ic), _get_current_gear(ib), _get_current_gear(ia)]
    xor = 0
    for b in cmd[1:]: xor ^= b
    cmd.append(xor)
    return bytearray(cmd)


# ══════════════════════════════════════════════════════════════
# 响应判断
# ══════════════════════════════════════════════════════════════

def _is_ok(rx: bytes | bytearray) -> bool:
    if len(rx) > 4 and rx[4] == 0x30:   return True
    if len(rx) > 15 and rx[15] == 0x30: return True
    if len(rx) > 177 and rx[4] == 0x50: return True
    return False


# ══════════════════════════════════════════════════════════════
# 传输层
# ══════════════════════════════════════════════════════════════

class _SerialTransport:
    """串口传输：持久连接，带收发和响应读取"""

    def __init__(self, port: str, baudrate: int = 9600, timeout: float = 2.0):
        self._port = port
        self._baudrate = baudrate
        self._timeout = timeout
        self._ser: serial.Serial | None = None

    def connect(self) -> bool:
        self._ser = serial.Serial(
            port=self._port,
            baudrate=self._baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False,
            timeout=self._timeout,
            write_timeout=5.0,
        )
        # 对应 C++ dcb.fRtsControl = RTS_CONTROL_ENABLE
        self._ser.setRTS(True)
        self._ser.setDTR(True)
        self._ser.reset_input_buffer()
        self._ser.reset_output_buffer()
        return self._ser.is_open

    def transact(self, frame: bytearray,
                 extra_wait: float = 0.2,
                 settle: float = 0.1) -> bytes:
        """发送帧并等待响应返回，失败返回空字节"""
        if not (self._ser and self._ser.is_open):
            raise OSError(f"串口 {self._port} 未打开")
        self._ser.reset_input_buffer()
        log.debug("SER TX [%dB]: %s", len(frame), frame.hex(' ').upper())
        self._ser.write(bytes(frame))
        time.sleep(extra_wait)
        # 等待 RX 队列稳定
        prev = -1
        for _ in range(20):
            time.sleep(settle)
            n = self._ser.in_waiting
            if n == prev and n > 0:
                break
            prev = n
        n = self._ser.in_waiting
        rx = self._ser.read(n) if n > 0 else b''
        if rx:
            log.debug("SER RX [%dB]: %s", len(rx), rx.hex(' ').upper())
        return rx

    def send(self, data: bytearray):
        """仅发送，不读响应（兼容旧调用路径）"""
        if not (self._ser and self._ser.is_open):
            raise OSError(f"串口 {self._port} 未打开")
        log.debug("SER TX [%dB]: %s", len(data), data.hex(' ').upper())
        self._ser.reset_input_buffer()
        self._ser.write(bytes(data))

    def close(self):
        if self._ser and self._ser.is_open:
            self._ser.close()
            self._ser = None


class _TcpTransport:
    """UDP 传输：无状态，每帧 sendto"""

    def __init__(self, host: str, port: int,
                 local_ip: str = '', local_port: int = 0,
                 timeout: float = 3.0):
        self._host = host
        self._port = port
        self._local_ip = local_ip
        self._local_port = local_port
        self._timeout = timeout
        self._sock: socket.socket | None = None
        self._dest = (host, port)

    def connect(self) -> bool:
        try:
            self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self._sock.settimeout(self._timeout)
            self._sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._sock.bind((self._local_ip, self._local_port))
            actual = self._sock.getsockname()
            log.debug("UDP local bind: %s:%s → dest %s:%s",
                      actual[0], actual[1], self._host, self._port)
            return True
        except Exception as e:
            log.error("Source UDP connect error: %s", e)
            return False

    def send(self, data: bytearray):
        if self._sock:
            sent = self._sock.sendto(bytes(data), self._dest)
            log.debug("UDP TX %dB → %s:%s  %s",
                      sent, self._dest[0], self._dest[1],
                      data.hex(' ').upper())

    def close(self):
        if self._sock:
            self._sock.close()
            self._sock = None


# ══════════════════════════════════════════════════════════════
# 源控制主类
# ══════════════════════════════════════════════════════════════

class SourceComm:
    """
    CL3021 交流标准源控制。

    串口模式：
        connect() 完成完整初始化（联机→AC版面→设定线制），持久保持串口。
        set_ac() 依次发送 72B角度帧 + 41B幅值帧 + 14B频率帧，发两轮确保生效。
        disconnect() 发退出命令并关闭串口。

    TCP/UDP 模式：
        每帧通过 UDP sendto 发送，无状态，行为与原版一致。
    """

    def __init__(self, mode: str = 'tcp',
                 host: str = '192.168.0.50', port: int = 10003,
                 local_ip: str = '', local_port: int = 10005,
                 serial_port: str = 'COM1', baudrate: int = 9600):
        self.mode = mode.lower()
        self.host = host
        self.port = port
        self.local_ip = local_ip
        self.local_port = local_port
        self.serial_port = serial_port
        self.baudrate = baudrate
        self._connected = False
        self._ser_transport: _SerialTransport | None = None
        # 与 SourceUdp 对齐的字段（test_engine 按同一套接口驱动两种传输）
        self.send_gear_frames = True
        self.max_current_a = float('inf')    # 串口模式限幅由调用方(test_engine)前置校验
        self.last_point: dict | None = None

    # ── 连接 / 断开 ───────────────────────────────────────────

    def connect(self) -> bool:
        if self.mode == 'serial':
            return self._connect_serial()
        else:
            return self._connect_udp()

    def _connect_serial(self) -> bool:
        t = _SerialTransport(self.serial_port, self.baudrate)
        if not t.connect():
            raise OSError(f"无法打开串口 {self.serial_port}（端口不存在或被占用）")
        self._ser_transport = t

        # Step 1: 联机
        rx = t.transact(_build_connect(), extra_wait=0.5, settle=0.2)
        if not rx:
            t.close(); self._ser_transport = None
            raise OSError("联机无响应，请确认设备已上电且串口接线正确")
        if not _is_ok(rx):
            log.warning("联机响应非标准 OK（可能是设备ID串，继续）: %s", rx.hex(' '))

        # Step 2: 切换 AC 输出版面
        rx = t.transact(_build_switch_screen_cmd(0x01), extra_wait=0.2, settle=0.1)
        if not _is_ok(rx):
            log.warning("AC版面切换无OK响应，继续: %s", rx.hex(' ') if rx else '(空)')

        # Step 3: 设定线制
        rx = t.transact(_build_set_line(), extra_wait=0.2, settle=0.1)
        if not _is_ok(rx):
            log.warning("设定线制无OK响应，继续: %s", rx.hex(' ') if rx else '(空)')

        self._connected = True
        log.info("串口源 %s 初始化完成", self.serial_port)
        return True

    def _connect_udp(self) -> bool:
        try:
            self._send_raw(bytearray([0x81, 0x01, 0x25, 0x06, 0xC9, 0xEB]))
            self._connected = True
            return True
        except Exception as e:
            self._connected = False
            raise OSError(f"online 帧发送失败: {e}") from e

    def disconnect(self):
        if self.mode == 'serial' and self._ser_transport:
            try:
                # 退出 AC 版面
                self._ser_transport.transact(
                    _build_switch_screen_cmd(0x00), extra_wait=0.1, settle=0.1)
            except Exception:
                pass
            self._ser_transport.close()
            self._ser_transport = None
        self._connected = False

    @property
    def connected(self) -> bool:
        return self._connected

    # ── 底层发送 ──────────────────────────────────────────────

    def _send_raw(self, data: bytearray):
        """仅发送，不读响应。TCP/UDP 模式和内部初始化调用"""
        log.debug("SOURCE TX [%d]: %s", len(data), data.hex(' ').upper())
        try:
            if self.mode == 'serial':
                if not self._ser_transport:
                    raise OSError(f"串口未打开: {self.serial_port}")
                self._ser_transport.send(data)
            else:
                t = _TcpTransport(self.host, self.port, self.local_ip, self.local_port)
                if not t.connect():
                    raise OSError(f"无法连接 UDP {self.host}:{self.port}")
                t.send(data)
                t.close()
        except Exception:
            self._connected = False
            raise

    def _transact_serial(self, frame: bytearray,
                         extra_wait: float = 0.2,
                         settle: float = 0.1) -> bytes:
        """串口专用：发送 + 读响应"""
        if not self._ser_transport:
            raise OSError("串口未打开")
        try:
            rx = self._ser_transport.transact(frame, extra_wait, settle)
            return rx
        except Exception:
            self._connected = False
            raise

    # ── 公开接口 ──────────────────────────────────────────────

    def set_ac(self, quc: float, qub: float, qua: float,
               qic: float, qib: float, qia: float,
               uc: float, ub: float, ua: float,
               ic: float, ib: float, ia: float,
               f: float, settle_s: float = 5.0):
        """
        设置 AC 源输出。

        串口模式：发送三条帧（72B角度+幅值、41B幅值、14B频率），重复两轮。
        TCP 模式：发送 72B 帧（原有行为）。

        参数：
            quc/qub/qua  - C/B/A 相电压角度（°）
            qic/qib/qia  - C/B/A 相电流角度（°），PF=cos(theta), theta=ia-ua角度差
            uc/ub/ua     - C/B/A 相电压（V）
            ic/ib/ia     - C/B/A 相电流（A）
            f            - 频率（Hz）
            settle_s     - 稳定等待时间（s）
        """
        if self.mode == 'serial':
            self._set_ac_serial(quc, qub, qua, qic, qib, qia,
                                uc, ub, ua, ic, ib, ia, f)
        else:
            self._set_ac_udp(quc, qub, qua, qic, qib, qia,
                             uc, ub, ua, ic, ib, ia, f)
        if settle_s > 0:
            time.sleep(settle_s)

    def _set_ac_serial(self, quc, qub, qua, qic, qib, qia,
                       uc, ub, ua, ic, ib, ia, freq):
        """串口模式：三帧两轮控制"""
        ang = _build_angle_update(quc, qub, qua, qic, qib, qia, freq)
        amp = _build_amplitude_update(uc, ub, ua, ic, ib, ia)
        frq = _build_freq_update(freq)

        for round_n in range(2):
            rx = self._transact_serial(ang, extra_wait=0.2, settle=0.1)
            if not _is_ok(rx):
                log.debug("角度帧 round%d 无OK: %s", round_n+1,
                          rx.hex(' ') if rx else '空')

            rx = self._transact_serial(amp, extra_wait=0.2, settle=0.1)
            if not _is_ok(rx):
                log.debug("幅值帧 round%d 无OK: %s", round_n+1,
                          rx.hex(' ') if rx else '空')

            rx = self._transact_serial(frq, extra_wait=0.2, settle=0.1)
            if not _is_ok(rx):
                log.debug("频率帧 round%d 无OK: %s", round_n+1,
                          rx.hex(' ') if rx else '空')

    def _set_ac_udp(self, quc, qub, qua, qic, qib, qia,
                    uc, ub, ua, ic, ib, ia, freq):
        """UDP 模式：发送 72B 角度帧（振幅由 41B 控制，此帧振幅为 0）"""
        cmd = _build_angle_update(quc, qub, qua, qic, qib, qia, freq)
        self._send_raw(cmd)

    def set_zero(self, settle_s: float = 5.0):
        """关源（全零输出）"""
        self.set_ac(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50.0, settle_s)

    def switch_screen(self, inter: int):
        """切换界面：0x01 AC输出版面，0x00 主界面"""
        if self.mode == 'serial':
            self._transact_serial(_build_switch_screen_cmd(inter),
                                  extra_wait=0.2, settle=0.1)
        else:
            self._send_raw(_build_switch_screen_cmd(inter))

    def set_gear_mode(self, mode_bin: str = '00000000'):
        """设置档位切换模式"""
        self._send_raw(_build_gear_mode(mode_bin))

    def set_voltage_gear(self, uc: float, ub: float, ua: float):
        """根据电压值自动选档"""
        self._send_raw(_build_voltage_gear(uc, ub, ua))

    def set_current_gear(self, ic: float, ib: float, ia: float):
        """根据电流值自动选档"""
        self._send_raw(_build_current_gear(ic, ib, ia))

    # ── 与 SourceUdp 对齐的测点接口（串口模式用；UDP 模式请直接用 SourceUdp）──

    def init(self):
        """串口模式初始化已在 connect() 内完成，此处保持接口一致。"""
        return

    warm_init = init

    def ensure_gear_mode(self):
        """重发档位模式帧（切频率的前提）——与 SourceUdp 同接口，供引擎统一调用。"""
        self.set_gear_mode('00000000')

    def reinit_output(self):
        """串口模式重开输出：重切 AC 版面。"""
        self.switch_screen(0x01)

    def set_point(self, s: dict, settle_s: float, force: bool = False):
        """按测点字典下发（键同 SourceUdp.set_point）。force 仅为接口兼容，串口模式无跳发缓存。"""
        del force
        if self.send_gear_frames:
            self.set_voltage_gear(s["uc"], s["ub"], s["ua"])
            time.sleep(1.0)
            self.set_current_gear(s["ic"], s["ib"], s["ia"])
            time.sleep(0.5)
        self.set_ac(s["quc"], s["qub"], s["qua"], s["qic"], s["qib"], s["qia"],
                    s["uc"], s["ub"], s["ua"], s["ic"], s["ib"], s["ia"],
                    s["freq"], settle_s=settle_s)
        self.last_point = dict(s)


class SourceUdp:
    """CL3021 UDP 控源（逐帧带 ACK 校验）——**自供电电表台面的唯一有效路径**。

    2026-07-08 实测: SourceComm 的 udp 模式 set_ac 只发 72B 角度帧（幅值/频率帧不发、
    不校回执），源实际无输出；有效路径 = 逐帧发送 联机/切屏/档位/角度/幅值/频率 并核对
    6B ACK（与其 serial 模式同序）。本类原在 projects/AcuRev100/tests/helpers_accuracy.py
    内实现并在台面调通（2026-07-08~07-15 多轮实证），2026-07-27 下沉到本模块统一供
    tools/accuracy_quick 与 tests/ 共用，避免两份分叉。

    自供电电表要点（电表电源取自源 Va/Vn，源掉输出 = 电表断电）：
      - 🔴 **角度帧一发必掉源、电表必重启（2026-07-28 实测，6/6）**：
          实验——100V/0A 上只切角度（电压/频率/档位全程不动）连切 6 次：台面目击 6 次表全部
          重启；`DEVICE_RUN_TIME` 每次丢 0.7~1.9s（空窗本底仅 ±0.5s 量化噪声，均值 +0.04s），
          即每次重启的冷启动代价约 1s（该寄存器重启不清零，靠"丢秒"识别）。
          旧记录"只有频率切换才掉源、纯角度帧不掉"由此作废：005/006 是仅有的"角度≠标准三相"
          用例组，也正是仅有的频繁掉源组；002/003/004 全部命中角度缓存跳发角度帧 → 零重启。
          是"帧里带的 freq 字段被重写"还是"角度突变本身"尚未隔离，处置相同。
          ⇒ **一条用例发几个角度帧，电表就重启几次**，故本类：
          · 角度帧只发一轮（每多发一轮就多一次重启），且**先把电流降到 0 再切角度**；
          · 独立频率帧只在频率真变时发（同值不发，_last_freq 缓存）——纯减帧。
      - 切频率前源必须处于档位模式（ensure_gear_mode / init 都会发该帧）
      - 电流硬限幅 max_current_a + 逐相 max_current_a_phase（2026-07-09 烧板事故后的硬门禁）
    """

    def __init__(self, host: str, port: int, local_port: int, timeout: float = 3.0,
                 max_current_a: float = 0.1, send_gear_frames: bool = False,
                 pin_voltage_v: float = 0.0, pin_current_a: float = 0.0,
                 assume_angles: "tuple | None" = None,
                 max_current_a_phase: "dict | None" = None,
                 state_path: "str | None" = None):
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self._sock.bind(("0.0.0.0", local_port))
        self._sock.settimeout(timeout)
        self._dest = (host, port)
        self.host = host
        self.port = port
        self.local_port = local_port
        self.max_current_a = max_current_a   # 🔴 电流硬限幅(烧板事故后): 超限拒绝下发
        # 🔴 逐相上限(台体各相回路承载不同): 2026-07-27 实机 3E4WY 三相 20A 时源报 "Ic 过载"
        #    → config source.max_current_a_phase = {a:20, b:20, c:15}; 缺省=全相同全局限幅
        self.max_current_a_phase = {ph: float((max_current_a_phase or {}).get(ph, max_current_a))
                                    for ph in ("a", "b", "c")}
        self.send_gear_frames = send_gear_frames   # 2026-07-14 用户指示: 默认不主动切档(源自动档接管)
        self.pin_voltage_v = pin_voltage_v   # >0 = 档位钉死模式(config source.gear_pin, 见 _pin_gears)
        self.pin_current_a = pin_current_a
        self._inited = False
        self._connected = False              # GUI 状态灯用: _cmd 无回执/被拒即置 False
        self._ramped = False                 # 本会话是否已做过电流软启动
        # 🔴 暖启动角度缓存预置(2026-07-15 新固件缺陷规避): 瞬断会触发电表 ADC 异常测量恒0
        # (需恢复出厂设置才恢复, 见 knowledge context 疑似缺陷记录)。台面源已处于标准三相角度/50Hz
        # 时, config source.assume_angles 预置缓存 → warm_init 后首点跳发同值角度帧, 批内零瞬断。
        # ⚠️ 前提是台面实况与预置一致; 台面角度/频率被手动改过时必须删掉 config 该行。
        self.assume_angles = tuple(assume_angles) if assume_angles else None
        # 🔴 源角度状态落盘(2026-07-28): 角度帧一发电表必重启一次, 所以"下个进程首点要不要发
        #    角度帧"直接等于"要不要白送一次重启"。config 的 assume_angles 是写死的常量, 只在
        #    "台面收在标准三相"时成立, 逼得每条非标角度用例退场都得把角度拉回标准(=再赔一次
        #    重启)。改为每次角度帧成功下发即把角度/电压落盘, 下个进程照实况恢复缓存 ⇒ 退场
        #    不必还原, 进程崩了也不会像常量那样说谎。
        #    ⚠️ 落的是"我们发了什么", 不是"源在出什么": 有人手动拨面板照样看不见 ——
        #       这个洞由 helpers_accuracy 的电表反证(拿表实测相角对缓存)来堵。
        self.state_path = Path(state_path) if state_path else None
        self._last_vgear = None              # 上次电压档位号(ua,ub,uc 换算后); 同档不重发
        self._last_igear = None              # 上次电流档位号; 同档不重发
        self._last_angles = None             # 上次角度帧内容(6角+freq); 同值不重发——角度帧带 freq,
        #                                      一发即整波形重设 → 瞬断输出(见类注释)
        self._last_freq = (float(self.assume_angles[6])
                           if self.assume_angles and len(self.assume_angles) > 6 else None)
        #                                    ↑ 已生效频率; 同值不重发独立频率帧
        self.last_point = None               # 最近成功下发的完整测点(用例间 idle 保档位用)

    # ── 连接状态（供 GUI 复用 SourceComm 那套按钮逻辑）──

    def connect(self) -> bool:
        """联机探活: 发联机帧并核对回执。"""
        self._cmd("联机", _build_connect())
        self._connected = True
        return True

    def disconnect(self):
        self._connected = False
        self.close()

    @property
    def connected(self) -> bool:
        return self._connected

    def _cmd(self, name: str, frame, retries: int = 4):
        if TRACE_FRAMES:
            print(f"[frame {time.strftime('%H:%M:%S')}] {name}", flush=True)
        for attempt in range(1, retries + 1):
            self._sock.sendto(bytes(frame), self._dest)
            try:
                data, _ = self._sock.recvfrom(1024)
            except socket.timeout:
                if attempt == retries:
                    self._connected = False
                    raise RuntimeError(f"CL3021 无回执: {name}")
                time.sleep(0.8 * attempt)    # 退避重试(批跑中源偶发不应答)
                continue
            if _is_ok(data):
                time.sleep(0.15)             # 帧间留隙, 避免连发把源打蒙
                return
            if attempt == retries:
                self._connected = False
                raise RuntimeError(f"CL3021 拒绝命令 {name}: {data.hex(' ')[:40]}")

    @staticmethod
    def angles_of(s: dict) -> tuple:
        """测点的角度帧内容(6 角 + freq)——判"这一点要不要发角度帧"的唯一口径。"""
        return (s["qua"], s["qub"], s["quc"], s["qia"], s["qib"], s["qic"], s["freq"])

    def angles_pending(self, s: dict) -> bool:
        """该测点是否会触发角度帧下发(=会瞬断源输出、自供电表会重启)。

        调用方据此把角度切换安排成"0A 上切 + 等电表重启复活 + 再加载幅值"的可控动作,
        而不是让瞬断砸在读数窗口里(见 helpers_accuracy.prearm_angles)。
        """
        return self.angles_of(s) != self._last_angles

    def load_state(self) -> "dict | None":
        """读落盘的源状态: {"angles": [6角+freq], "volts": [ua,ub,uc], "updated": ...}。

        读不到/格式不对一律返回 None(调用方回退 config assume_angles), 状态文件坏了不该拖垮批跑。
        """
        if not self.state_path or not self.state_path.exists():
            return None
        try:
            st = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        ang = st.get("angles")
        if not isinstance(ang, list) or len(ang) != 7:
            return None
        return st

    def _save_state(self, s: dict):
        """角度帧成功下发后落盘(失败只告警不抛: 落盘是优化项, 不能让批跑失败)。"""
        if not self.state_path:
            return
        try:
            self.state_path.write_text(json.dumps({
                "angles": list(self.angles_of(s)),
                "volts": [s["ua"], s["ub"], s["uc"]],
                "updated": time.strftime("%Y-%m-%d %H:%M:%S"),
                "note": "CL3021 最近一次成功下发的角度/频率/电压(供下个进程免盲发角度帧; "
                        "台面被手动改动后本文件即失准, 由电表反证兜底)",
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError as exc:
            log.warning("源角度状态落盘失败(不影响本次运行): %s", exc)

    def current_angles(self) -> "tuple | None":
        """源当前角度缓存(6角+freq); None = 会话状态未知, 下一帧必然全发。

        与 last_point 的区别: 本会话还没下发过测点时 last_point 是 None, 但缓存可能已由
        落盘状态/assume_angles 预置 —— 保活点要沿用角度就得看这里, 看 last_point 会漏。
        """
        return self._last_angles

    def prime_angles(self, angles: "tuple | None"):
        """预置角度缓存(台面实况已知与之一致时): 使下一点跳发角度帧, 零瞬断。"""
        self._last_angles = tuple(angles) if angles else None
        if angles and len(angles) > 6:
            self._last_freq = float(angles[6])

    def warm_init(self):
        """暖初始化: 联机+档位模式, 不切屏(输出不断)。

        源必须处于档位模式才能平滑切频率——2026-07-13 实证: 跳过档位模式帧时
        50→60Hz 切换即掉输出; 带档位模式的扫描连切 7 频点零断电。切屏帧才会瞬断输出。
        """
        # 会话状态未知默认首点全帧; 有 assume_angles 预置(台面处于标准角度/50Hz)则跳发角度帧
        self.prime_angles(self.assume_angles)
        self._cmd("联机", _build_connect())
        time.sleep(0.5)
        self._cmd("档位模式", _build_gear_mode("00000000"))
        time.sleep(0.5)
        self._pin_gears()
        self._connected = True
        self._inited = True

    def _pin_gears(self):
        """档位钉死(2026-07-14 调通提速): 会话开始一次性把电压/电流档钉到批内最大量程。

        之后 _tx_point 全程不发档位帧。历史动因是"002/003 批源反复掉0+电表不断重启"，
        但 2026-07-27 澄清后已知**档位切换并不掉源输出、掉输出的只有频率切换** ⇒ 钉档不再是必需。
        代价明确: 低幅值点用大档输出、**源自身精度下降 0.2% 量级** ⇒ 精度测试不要开
        (tools/accuracy_quick 默认关, 见 config source.precision_tool_gear_pin)。
        """
        if not self.pin_voltage_v:
            return
        v = self.pin_voltage_v
        self._cmd("电压档位(钉死)", _build_voltage_gear(v, v, v))
        time.sleep(1.0)
        gv = _get_voltage_gear(v)
        self._last_vgear = (gv, gv, gv)
        gi = None
        if self.pin_current_a:
            i = self.pin_current_a
            self._cmd("电流档位(钉死)", _build_current_gear(i, i, i))
            time.sleep(0.5)
            gi = _get_current_gear(i)
            self._last_igear = (gi, gi, gi)
        log.info("[gear] 档位已钉死: 电压≤%sV(档%s)%s; 本批不再切档",
                 v, gv, f", 电流≤{self.pin_current_a}A(档{gi})" if gi else "")

    def ensure_gear_mode(self):
        """重发"档位模式"帧——切频率前的必要前提, 该帧本身不影响输出。

        2026-07-13 实证: 跳过档位模式帧时 50→60Hz 切换即掉输出; 处于档位模式时连切 7 个频点零断电。
        """
        self._cmd("档位模式", _build_gear_mode("00000000"))
        time.sleep(0.3)

    def mark_inited(self):
        """暖启动探测成功后调用: 源已就绪, 本会话免冷初始化。"""
        self._inited = True

    def is_inited(self) -> bool:
        return self._inited

    def init(self):
        self._last_angles = None             # 冷初始化后源状态重置, 首个测点全帧下发
        self._last_freq = None
        if self._inited:                     # 会话级只做一次(反复联机/切屏会把源搞失联)
            self._cmd("联机", _build_connect())
            return
        self._cmd("联机", _build_connect())
        time.sleep(1)
        self._cmd("切AC界面", _build_switch_screen_cmd(0x01))
        time.sleep(3)
        self._cmd("档位模式", _build_gear_mode("00000000"))
        time.sleep(0.5)
        self._pin_gears()
        self._inited = True
        self._connected = True

    def reinit_output(self):
        """源冷重臂: 输出被关闭(保护/幅值帧被吞停0)后, 唯一能重开输出的是切屏帧。

        2026-07-14 实证(003_01_case2 点2): 源输出掉0后只发保活幅值帧拉不回来,
        电表长时间无电; 冷初始化(联机→切屏→档位模式→钉档)一次即恢复、电表秒起。
        """
        log.info("[src] 源冷重臂: 联机→切屏(重开输出)→档位模式→钉档")
        self._inited = False
        self.init()

    def _tx_point(self, s: dict):
        # 🔑 掉源动作(2026-07-28 实测, 见类注释): 角度帧每发必掉、电表必重启(6/6) > 频率真切换
        #    >> 同值频率帧(002/004 每点发两轮从不掉源)。发几帧就重启几次, 减帧=减重启。
        #    故角度/频率切换一律: 源在档位模式 + 0A 上切 + 只发一轮 + 切后等复活验证。
        # 优先级: gear_pin 钉死(会话级一次切档; 代价是低幅值点源精度差, 精度测试不用)
        #        > send_gear_frames 逐点档位(按"档位号"缓存, 同档不发; 精度测试的默认选择)
        #        > 全不发(源自动档——2026-07-14 实证不接管, 且档位滞留会触发过载保护, 已弃用)。
        if self.pin_voltage_v:
            vmax = max(s["ua"], s["ub"], s["uc"])
            imax_pt = max(abs(s["ia"]), abs(s["ib"]), abs(s["ic"]))
            if vmax > self.pin_voltage_v or imax_pt > (self.pin_current_a or 0.0):
                raise RuntimeError(
                    f"测点幅值(U≤{vmax}V, I≤{imax_pt}A)超出钉死档量程"
                    f"(电压{self.pin_voltage_v}V/电流{self.pin_current_a}A), 大档打小幅值安全但"
                    f"反向会触发源过载保护(2026-07-14 10:05 台面目击)——请调大 config "
                    f"source.gear_pin 后重跑")
        elif self.send_gear_frames:
            vgear = (_get_voltage_gear(s["ua"]), _get_voltage_gear(s["ub"]),
                     _get_voltage_gear(s["uc"]))
            igear = (_get_current_gear(s["ia"]), _get_current_gear(s["ib"]),
                     _get_current_gear(s["ic"]))
            if (vgear != self._last_vgear or igear != self._last_igear) and self.last_point:
                # 🔴 切档一律在"0A + 两档都容得下的电压"上做(2026-07-28)。跨档本身实测不掉源,
                #    但已记录的过载锁存全部发生在"大幅值撞小量程"上:
                #      · 带载切档 → 电流灌进小电流档;
                #      · 降压点降档时若先发新档 → 旧的高电压(如 270V)留在 ≤240V 档里 = 超量程。
                #    故预置帧取**逐相 min(上一点电压, 本点电压) + 0A**: 降压时先降到新值(旧档更大,
                #    安全), 升压时保持旧值(切到更大的新档后再由正式幅值帧升上去)。
                prev = self.last_point
                self._cmd("幅值帧(切档前降流降压)", _build_amplitude_update(
                    min(prev["uc"], s["uc"]), min(prev["ub"], s["ub"]),
                    min(prev["ua"], s["ua"]), 0.0, 0.0, 0.0))
                time.sleep(0.3)
            if vgear != self._last_vgear:
                self._cmd("电压档位", _build_voltage_gear(s["uc"], s["ub"], s["ua"]))
                time.sleep(1.0)          # 档位继电器切换留时(切档后立发幅值帧疑被吞→输出停摆)
                self._last_vgear = vgear
            if igear != self._last_igear:    # 电流档位同理: 同档不重发
                self._cmd("电流档位", _build_current_gear(s["ic"], s["ib"], s["ia"]))
                self._last_igear = igear
            time.sleep(0.5)
        # 角度帧带 freq 字段 → 6角+freq 与上次一致时跳发(零瞬断); 真改相角/频率时才发。
        angles = self.angles_of(s)
        freq_changed = self._last_freq is None or float(s["freq"]) != self._last_freq
        if angles != self._last_angles:
            # 🟡 角度切换按"会瞬断"处理: ①先把电流降到 0 再切(切换发生在无载上, 不触发源过载保护);
            #    ②只发一轮——每多发一轮就多一次潜在瞬断、自供电表就多重启一次(2026-07-28 005/006 整改)。
            #    电表的重启等待由调用方负责(helpers_accuracy.prearm_angles 在 0A 上等复活再加载幅值)。
            if max(abs(s["ia"]), abs(s["ib"]), abs(s["ic"])) > 0:
                self._cmd("幅值帧(切角前降流)",
                          _build_amplitude_update(s["uc"], s["ub"], s["ua"], 0.0, 0.0, 0.0))
            self._cmd("角度帧", _build_angle_update(s["quc"], s["qub"], s["qua"],
                                                    s["qic"], s["qib"], s["qia"], s["freq"]))
            self._last_angles = angles       # 该帧已 ACK 即生效, 立刻记账(异常路径不会误跳发)
            self._save_state(s)              # 跨进程记账: 下个进程照实况恢复, 免白送一次重启
        for _rnd in (1, 2):                  # 幅值两轮确保生效(与 serial 模式一致)
            self._cmd("幅值帧", _build_amplitude_update(s["uc"], s["ub"], s["ua"],
                                                        s["ic"], s["ib"], s["ia"]))
            if freq_changed:                 # 频率同值不重发: 少一次"频率写入"就少一次瞬断风险
                self._cmd("频率帧", _build_freq_update(s["freq"]))
        self._last_freq = float(s["freq"])

    def set_point(self, s: dict, settle_s: float, force: bool = False):
        imax = max(abs(s["ia"]), abs(s["ib"]), abs(s["ic"]))
        if imax > self.max_current_a:        # 🔴 硬限幅: 直连 0.1A / 经CT 25A(见 config source 段)
            raise RuntimeError(
                f"电流 {imax}A 超过接线方式允许上限 {self.max_current_a}A, 拒绝下发 "
                f"(config source.current_injection 与实际接线核对后再跑)")
        for ph in ("a", "b", "c"):           # 🔴 逐相限幅: 台面 C 相回路只到 15A(20A 会过载)
            val = abs(s[f"i{ph}"])
            cap = self.max_current_a_phase[ph]
            if val > cap:
                raise RuntimeError(
                    f"{ph.upper()} 相电流 {val}A 超过该相上限 {cap}A, 拒绝下发 "
                    f"(config source.max_current_a_phase; 台面回路承载不足时源会报过载并可能锁存保护)")
        # 同点全跳发: 与上次完全相同的点重发本无信息量, 一帧不发即零风险。
        # (2026-07-27 澄清: 早前把"0A/低压幅值帧"列为瞬断第二元凶是误判——只要 Va 还在,
        #  电表就不掉电, 零电流输出本身无任何风险; 真正会打断 Va 的是**电压档位切换**。)
        # 救源类"重发本测点/拉保活"必须传 force=True 绕过跳发(那正是要重新断言输出的场合);
        # 冷初始化/冷重臂会清 _last_angles, 其后首点自动全帧下发不受影响。
        if not force and s == self.last_point and self._last_angles is not None:
            if settle_s > 0:
                time.sleep(settle_s)
            return
        if imax > 0.1 and not self._ramped:  # 会话首次出电流: 先 0.05A 软启动 2s
            warm = dict(s)
            warm["ia"] = warm["ib"] = warm["ic"] = min(0.05, imax)
            self._tx_point(warm)
            time.sleep(2)
            self._ramped = True
        self._tx_point(s)
        self.last_point = dict(s)            # 记住最近下发点(idle 保档位用)
        if settle_s > 0:
            time.sleep(settle_s)

    def close(self):
        self._sock.close()
