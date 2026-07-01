"""
源控制模块（CL3021 标准源）
支持串口和 TCP/UDP 两种传输方式，使用相同的 CL3021 二进制协议

串口协议基于实际抓包验证（2026-06-03）：
  - 初始化: 联机(6B) → AC版面(10B) → 设定线制(10B)
  - 控制:   角度帧(72B) + 幅值帧(41B) + 频率帧(14B)，发两轮确保生效
  - 校验:   CS = XOR(byte[1] … byte[n-2])
"""
import socket
import struct
import serial
import logging
import time

log = logging.getLogger(__name__)


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


def _get_voltage_gear(v: float) -> int:
    if v <= 30:  return 5
    if v <= 60:  return 4
    if v <= 120: return 3
    if v <= 240: return 2
    if v <= 480: return 1
    return 0


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
