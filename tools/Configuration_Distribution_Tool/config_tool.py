"""
样机配置写入工具
依赖: pip install pyserial
"""

import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import socket
import struct
import time
from datetime import datetime, timezone


# ─────────────────────────── Modbus 工具函数 ───────────────────────────

def crc16_modbus(data: bytes) -> int:
    crc = 0xFFFF
    for byte in data:
        crc ^= byte
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return crc


def bytes_to_hex(data: bytes) -> str:
    return ' '.join(f'{b:02x}' for b in data)


def verify_response_crc(resp: bytes) -> tuple[bool, str]:
    if len(resp) < 3:
        return False, f"{bytes_to_hex(resp)}  [报文过短]"
    payload  = resp[:-2]
    recv_crc = resp[-2:]
    calc_crc = crc16_modbus(payload)
    calc_lo  = calc_crc & 0xFF
    calc_hi  = (calc_crc >> 8) & 0xFF
    data_hex = bytes_to_hex(payload)
    crc_hex  = bytes_to_hex(recv_crc)
    if recv_crc[0] == calc_lo and recv_crc[1] == calc_hi:
        return True, f"{data_hex} | CRC: {crc_hex}  [OK]"
    else:
        expect = f"{calc_lo:02x} {calc_hi:02x}"
        return False, f"{data_hex} | CRC: {crc_hex}  [ERR 期望: {expect}]"


def parse_reg_addr(reg_str: str) -> int:
    """解析寄存器地址，支持 'F071' 或 'F0 71' 格式"""
    return int(reg_str.replace(' ', ''), 16)


def mac_to_ascii_bytes(mac_str: str) -> bytes:
    """MAC地址转12字节ASCII数据"""
    clean = mac_str.replace(':', '').replace('-', '').upper()
    if len(clean) != 12:
        raise ValueError(f"MAC地址格式错误: '{mac_str}'，应为12位十六进制（含或不含冒号）")
    return clean.encode('ascii')


def str_to_padded_bytes(s: str) -> bytes:
    """字符串转ASCII字节，不足偶数长度补0x00"""
    data = s.encode('ascii')
    if len(data) % 2 != 0:
        data += b'\x00'
    return data


def build_write_frame(slave_id: int, reg_addr: int, data: bytes) -> bytes:
    """构建写入帧 (FC=0x6A)，含CRC"""
    if len(data) % 2 != 0:
        data += b'\x00'
    reg_count  = len(data) // 2
    byte_count = len(data)
    frame = struct.pack('>BBHHB', slave_id, 0x6A, reg_addr, reg_count, byte_count)
    frame += data
    crc = crc16_modbus(frame)
    frame += struct.pack('<H', crc)
    return frame


def build_write_fc16_frame(slave_id: int, reg_addr: int, reg_count: int, value: int) -> bytes:
    """构建FC16写多寄存器帧，含CRC"""
    byte_count = reg_count * 2
    data = struct.pack('>H', value) * reg_count
    frame = struct.pack('>BBHHB', slave_id, 0x10, reg_addr, reg_count, byte_count)
    frame += data
    crc = crc16_modbus(frame)
    frame += struct.pack('<H', crc)
    return frame


def build_read_frame(slave_id: int, reg_addr: int, reg_count: int) -> bytes:
    """构建读取帧 FC=0x03，含CRC"""
    frame = struct.pack('>BBHH', slave_id, 0x03, reg_addr, reg_count)
    crc = crc16_modbus(frame)
    frame += struct.pack('<H', crc)
    return frame


def rtu_to_tcp_frame(rtu_frame: bytes, tid: int = 0) -> bytes:
    """RTU 帧 → Modbus TCP 帧：去掉 SlaveID/CRC，加 MBAP 头（事务ID/协议ID/长度/单元ID）"""
    unit_id = rtu_frame[0]
    pdu     = rtu_frame[1:-2]
    length  = 1 + len(pdu)
    return struct.pack('>HHHB', tid, 0, length, unit_id) + pdu


# ─────────────────────────── 主界面 ────────────────────────────────────

class ConfigTool:
    PARITY_MAP = {
        'none': serial.PARITY_NONE,
        'even': serial.PARITY_EVEN,
        'odd':  serial.PARITY_ODD,
    }

    CONFIG_FIELDS = [
        # (label,           default_value,        default_reg, dtype)
        ("MAC1",            "30:7A:57:01:F5:50",  "F071",      "mac"),
        ("MAC2",            "30:7A:57:01:F5:51",  "F081",      "mac"),
        ("MAC3",            "30:7A:57:01:F5:52",  "F056",      "mac"),
        ("SN",              "MDA250500082",        "F040",      "str"),
        ("HWversion",       "1.02",               "F050",      "str"),
        ("Function Model",  "2",                  "F052",      "int"),
        ("Current Type",    "1",                  "F054",      "int"),
    ]

    READ_DEFAULTS = [
        ("F071", "16"),
        ("F081", "16"),
        ("F056", "16"),
        ("F040", "16"),
        ("F050", "2"),
        ("F052", "1"),
        ("F054", "1"),
    ]

    DATA_TYPES = ["uint16_t", "uint8_t", "uint32_t", "int32_t", "uint64_t", "float", "timestamp"]

    SCALE_VALUES = ["1", "0.1", "0.01", "0.001", "0.0001"]

    # 连接探测：连上后向所配 Slave ID 读一帧，校验该从机是否真的在线
    PROBE_REG   = 0xF052   # 样机 Function Model 寄存器（READ_DEFAULTS 中已确认可读）
    PROBE_COUNT = 1

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("样机配置写入工具")
        self.root.resizable(True, True)
        self.serial_conn: serial.Serial | None = None
        self.tcp_conn: socket.socket | None = None
        self._build_ui()

    # ─── UI 构建 ───────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_conn_section()
        self._build_config_section()
        self._build_read_section()
        self._build_generic_section()

    def _build_conn_section(self):
        frame = tk.LabelFrame(self.root, text="连接配置",
                              font=('Microsoft YaHei', 9, 'bold'), padx=6, pady=4)
        frame.pack(fill='x', padx=10, pady=(8, 4))

        row = tk.Frame(frame)
        row.pack(fill='x', pady=2)

        # ── 模式选择 ──
        tk.Label(row, text="RTU/TCP").pack(side='left')
        self.mode_var = tk.StringVar(value="RTU")
        mode_cb = ttk.Combobox(row, textvariable=self.mode_var, width=5,
                               values=["RTU", "TCP"], state='readonly')
        mode_cb.pack(side='left', padx=(4, 10))
        mode_cb.bind('<<ComboboxSelected>>', lambda _: self._on_mode_change())

        # ── 协议参数动态区 ──
        self._proto_frame = tk.Frame(row)
        self._proto_frame.pack(side='left')

        # RTU 专用控件
        self._rtu_widgets = tk.Frame(self._proto_frame)
        tk.Label(self._rtu_widgets, text="波特率").pack(side='left')
        self.baud_var = tk.StringVar(value="19200")
        ttk.Combobox(self._rtu_widgets, textvariable=self.baud_var, width=8,
                     values=["9600", "19200", "38400", "57600", "115200"],
                     state='readonly').pack(side='left', padx=(4, 10))
        tk.Label(self._rtu_widgets, text="Com").pack(side='left')
        self.port_var = tk.StringVar(value="COM4")
        self.port_cb = ttk.Combobox(self._rtu_widgets, textvariable=self.port_var, width=9)
        self.port_cb.pack(side='left', padx=4)
        tk.Button(self._rtu_widgets, text="⟳", width=2,
                  command=self._refresh_ports).pack(side='left', padx=(0, 10))
        tk.Label(self._rtu_widgets, text="parity").pack(side='left')
        self.parity_var = tk.StringVar(value="none")
        ttk.Combobox(self._rtu_widgets, textvariable=self.parity_var, width=6,
                     values=["none", "even", "odd"], state='readonly').pack(side='left', padx=4)

        # TCP 专用控件
        self._tcp_widgets = tk.Frame(self._proto_frame)
        tk.Label(self._tcp_widgets, text="IP 地址").pack(side='left')
        self.tcp_ip_var = tk.StringVar(value="192.168.3.37")
        tk.Entry(self._tcp_widgets, textvariable=self.tcp_ip_var, width=16).pack(side='left', padx=(4, 10))
        tk.Label(self._tcp_widgets, text="端口").pack(side='left')
        self.tcp_port_var = tk.StringVar(value="502")
        tk.Entry(self._tcp_widgets, textvariable=self.tcp_port_var, width=7).pack(side='left', padx=4)

        # ── 公共控件：Slave ID + 按钮 + 状态 ──
        tk.Label(row, text="Slave ID").pack(side='left', padx=(10, 0))
        self.slave_var = tk.StringVar(value="1")
        tk.Entry(row, textvariable=self.slave_var, width=5).pack(side='left', padx=(4, 8))
        tk.Button(row, text="连接", width=6, bg='#4CAF50', fg='white',
                  activebackground='#388E3C', command=self._connect).pack(side='left', padx=2)
        tk.Button(row, text="断开", width=6, bg='#f44336', fg='white',
                  activebackground='#C62828', command=self._disconnect).pack(side='left', padx=2)
        self.conn_label = tk.Label(row, text="● 未连接", fg='#c0392b', font=('', 9, 'bold'))
        self.conn_label.pack(side='left', padx=8)

        self._refresh_ports()
        self._toggle_conn_mode()

    def _build_config_section(self):
        frame = tk.LabelFrame(self.root, text="写入样机配置",
                              font=('Microsoft YaHei', 9, 'bold'), padx=6, pady=4)
        frame.pack(fill='x', padx=10, pady=4)

        self.field_vars: dict[str, tuple] = {}

        for label, default_val, default_reg, dtype in self.CONFIG_FIELDS:
            row = tk.Frame(frame)
            row.pack(fill='x', pady=2)

            tk.Label(row, text=label, width=14, anchor='w',
                     fg='#2980b9', font=('', 9, 'bold')).pack(side='left')

            val_var = tk.StringVar(value=default_val)
            tk.Entry(row, textvariable=val_var, width=22).pack(side='left', padx=(0, 8))

            tk.Label(row, text="Reg(hex)").pack(side='left')
            reg_var = tk.StringVar(value=default_reg)
            tk.Entry(row, textvariable=reg_var, width=8).pack(side='left', padx=4)

            tk.Label(row, text="报文").pack(side='left', padx=(8, 2))
            rtu_var = tk.StringVar()
            tk.Entry(row, textvariable=rtu_var, width=60,
                     state='readonly', readonlybackground='#fafafa',
                     font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

            self.field_vars[label] = (val_var, reg_var, rtu_var, dtype)

        btn_row = tk.Frame(frame)
        btn_row.pack(fill='x', pady=(6, 2))

        tk.Button(btn_row, text="转换", width=10, bg='#2196F3', fg='white',
                  activebackground='#1565C0', font=('', 9, 'bold'),
                  command=self._convert).pack(side='left')

        tk.Button(btn_row, text="写入寄存器", width=14, bg='#FF9800', fg='white',
                  activebackground='#E65100', font=('', 9, 'bold'),
                  command=self._write_all).pack(side='right')

    def _build_read_section(self):
        frame = tk.LabelFrame(self.root, text="读取样机配置",
                              font=('Microsoft YaHei', 9, 'bold'), padx=6, pady=4)
        frame.pack(fill='x', padx=10, pady=(4, 4))

        self.read_vars: list[tuple] = []

        for reg_default, num_default in self.READ_DEFAULTS:
            row = tk.Frame(frame)
            row.pack(fill='x', pady=2)

            tk.Label(row, text="Reg(hex)").pack(side='left')
            reg_var = tk.StringVar(value=reg_default)
            tk.Entry(row, textvariable=reg_var, width=8).pack(side='left', padx=4)

            tk.Label(row, text="Reg num").pack(side='left', padx=(8, 2))
            num_var = tk.StringVar(value=num_default)
            tk.Entry(row, textvariable=num_var, width=5).pack(side='left', padx=4)

            tk.Label(row, text="返回报文").pack(side='left', padx=(8, 2))
            result_var = tk.StringVar()
            tk.Entry(row, textvariable=result_var, width=60,
                     readonlybackground='#fffde7', state='readonly',
                     font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

            self.read_vars.append((reg_var, num_var, result_var))

        btn_row = tk.Frame(frame)
        btn_row.pack(fill='x', pady=(6, 2))

        tk.Button(btn_row, text="读取配置信息", width=16, bg='#9C27B0', fg='white',
                  activebackground='#6A1B9A', font=('', 9, 'bold'),
                  command=self._read_all).pack(side='left')

    def _build_generic_section(self):
        frame = tk.LabelFrame(self.root, text="通用功能",
                              font=('Microsoft YaHei', 9, 'bold'), padx=6, pady=4)
        frame.pack(fill='x', padx=10, pady=(4, 8))

        # ════════ 写入区 ════════
        write_btn_row = tk.Frame(frame)
        write_btn_row.pack(fill='x', pady=(4, 2))
        tk.Button(write_btn_row, text="写寄存器", width=10, bg='#FF9800', fg='white',
                  activebackground='#E65100', font=('', 9, 'bold'),
                  command=self._generic_write).pack(side='left')

        write_row = tk.Frame(frame)
        write_row.pack(fill='x', pady=2)

        tk.Label(write_row, text="Reg(hex)").pack(side='left')
        self.gw_reg_var = tk.StringVar(value="1004")
        tk.Entry(write_row, textvariable=self.gw_reg_var, width=8).pack(side='left', padx=4)

        tk.Label(write_row, text="Reg num").pack(side='left', padx=(8, 2))
        self.gw_num_var = tk.StringVar(value="1")
        tk.Entry(write_row, textvariable=self.gw_num_var, width=5).pack(side='left', padx=4)

        tk.Label(write_row, text="Data").pack(side='left', padx=(8, 2))
        self.gw_data_var = tk.StringVar(value="1")
        tk.Entry(write_row, textvariable=self.gw_data_var, width=8).pack(side='left', padx=4)

        tk.Label(write_row, text="发送报文").pack(side='left', padx=(8, 2))
        self.gw_frame_var = tk.StringVar()
        tk.Entry(write_row, textvariable=self.gw_frame_var, width=32,
                 state='readonly', readonlybackground='#fafafa',
                 font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

        tk.Label(write_row, text="返回报文").pack(side='left', padx=(8, 2))
        self.gw_resp_var = tk.StringVar()
        tk.Entry(write_row, textvariable=self.gw_resp_var, width=32,
                 state='readonly', readonlybackground='#fffde7',
                 font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

        # ════════ 读取区 ════════
        read_btn_row = tk.Frame(frame)
        read_btn_row.pack(fill='x', pady=(10, 2))
        tk.Button(read_btn_row, text="读寄存器", width=10, bg='#9C27B0', fg='white',
                  activebackground='#6A1B9A', font=('', 9, 'bold'),
                  command=self._generic_read).pack(side='left')

        self.gr_rows: list[tuple] = []
        for _ in range(2):
            read_row = tk.Frame(frame)
            read_row.pack(fill='x', pady=2)

            tk.Label(read_row, text="Reg(hex)").pack(side='left')
            reg_var = tk.StringVar(value="1004")
            tk.Entry(read_row, textvariable=reg_var, width=8).pack(side='left', padx=4)

            tk.Label(read_row, text="Reg num").pack(side='left', padx=(8, 2))
            num_var = tk.StringVar(value="1")
            tk.Entry(read_row, textvariable=num_var, width=5).pack(side='left', padx=4)

            tk.Label(read_row, text="Data_type").pack(side='left', padx=(8, 2))
            dtype_var = tk.StringVar(value="uint16_t")
            ttk.Combobox(read_row, textvariable=dtype_var, width=9,
                         values=self.DATA_TYPES, state='readonly').pack(side='left', padx=4)

            tk.Label(read_row, text="scale").pack(side='left', padx=(8, 2))
            scale_var = tk.StringVar(value="1")
            ttk.Combobox(read_row, textvariable=scale_var, width=6,
                         values=self.SCALE_VALUES, state='readonly').pack(side='left', padx=4)

            tk.Label(read_row, text="发送报文").pack(side='left', padx=(8, 2))
            send_var = tk.StringVar()
            tk.Entry(read_row, textvariable=send_var, width=32,
                     state='readonly', readonlybackground='#fafafa',
                     font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

            tk.Label(read_row, text="返回报文").pack(side='left', padx=(8, 2))
            frame_var = tk.StringVar()
            tk.Entry(read_row, textvariable=frame_var, width=32,
                     state='readonly', readonlybackground='#fffde7',
                     font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

            tk.Label(read_row, text="data").pack(side='left', padx=(8, 2))
            data_var = tk.StringVar()
            tk.Entry(read_row, textvariable=data_var, width=24,
                     state='readonly', readonlybackground='#e8f5e9',
                     font=('Courier New', 9, 'bold')).pack(side='left', padx=4)

            self.gr_rows.append((reg_var, num_var, dtype_var, scale_var, send_var, frame_var, data_var))

            # 绑定读取输入变化自动刷新发送报文预览
            reg_var.trace_add('write', self._refresh_gr_frame)
            num_var.trace_add('write', self._refresh_gr_frame)

        # 绑定写入输入变化自动刷新报文预览
        for var in (self.gw_data_var, self.gw_reg_var, self.gw_num_var):
            var.trace_add('write', self._refresh_gw_frame)

        # Slave ID 变化时联动刷新写入 / 读取发送报文预览
        self.slave_var.trace_add('write', self._refresh_gw_frame)
        self.slave_var.trace_add('write', self._refresh_gr_frame)

        self._refresh_gw_frame()
        self._refresh_gr_frame()

    # ─── 连接模式切换 ─────────────────────────────────────────────────

    def _toggle_conn_mode(self):
        if self.mode_var.get() == "RTU":
            self._tcp_widgets.pack_forget()
            self._rtu_widgets.pack(side='left')
        else:
            self._rtu_widgets.pack_forget()
            self._tcp_widgets.pack(side='left')

    def _on_mode_change(self):
        """切换通信方式：调整控件并按新格式刷新写入 / 读取报文预览"""
        self._toggle_conn_mode()
        self._refresh_gw_frame()
        self._refresh_gr_frame()

    # ─── 串口 / TCP 操作 ───────────────────────────────────────────────

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_cb['values'] = ports if ports else ["COM1"]
        if ports and self.port_var.get() not in ports:
            self.port_var.set(ports[0])

    def _connect(self):
        if self.mode_var.get() == "TCP":
            self._connect_tcp()
        else:
            self._connect_rtu()

    def _connect_rtu(self):
        if self.serial_conn and self.serial_conn.is_open:
            messagebox.showinfo("提示", "串口已处于连接状态")
            return
        try:
            self.serial_conn = serial.Serial(
                port=self.port_var.get(),
                baudrate=int(self.baud_var.get()),
                bytesize=8,
                parity=self.PARITY_MAP.get(self.parity_var.get(), serial.PARITY_NONE),
                stopbits=1,
                timeout=1.0,
            )
        except Exception as e:
            messagebox.showerror("连接失败", str(e))
            return
        slave_id = self._get_slave_id()
        ok, info = self._probe_slave(slave_id)
        if ok:
            self.conn_label.config(text=f"● 已连接 {self.port_var.get()} (从机 {slave_id})",
                                   fg='#27ae60')
        else:
            self.serial_conn.close()
            self.serial_conn = None
            self.conn_label.config(text="● 未连接", fg='#c0392b')
            messagebox.showerror("连接失败", info)

    def _connect_tcp(self):
        if self.tcp_conn:
            try:
                self.tcp_conn.close()
            except Exception:
                pass
            self.tcp_conn = None
        try:
            ip   = self.tcp_ip_var.get().strip()
            port = int(self.tcp_port_var.get().strip())
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(3.0)
            sock.connect((ip, port))
            sock.settimeout(2.0)
            self.tcp_conn = sock
        except Exception as e:
            self.tcp_conn = None
            messagebox.showerror("TCP 连接失败", str(e))
            return
        slave_id = self._get_slave_id()
        ok, info = self._probe_slave(slave_id)
        if ok:
            self.conn_label.config(text=f"● 已连接 {ip}:{port} (从机 {slave_id})", fg='#27ae60')
        else:
            try:
                self.tcp_conn.close()
            except Exception:
                pass
            self.tcp_conn = None
            self.conn_label.config(text="● 未连接", fg='#c0392b')
            messagebox.showerror("连接失败", info)

    def _disconnect(self):
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.serial_conn = None
        if self.tcp_conn:
            try:
                self.tcp_conn.close()
            except Exception:
                pass
        self.tcp_conn = None
        self.conn_label.config(text="● 未连接", fg='#c0392b')

    def _check_connected(self) -> bool:
        if self.mode_var.get() == "TCP":
            if not self.tcp_conn:
                messagebox.showwarning("提示", "请先连接 Modbus TCP")
                return False
            return True
        if not self.serial_conn or not self.serial_conn.is_open:
            messagebox.showwarning("提示", "请先连接串口")
            return False
        return True

    def _get_slave_id(self) -> int:
        try:
            return int(self.slave_var.get())
        except ValueError:
            return 1

    # ─── 传输层：RTU / TCP 统一收发 ───────────────────────────────────

    def _is_tcp(self) -> bool:
        return self.mode_var.get() == "TCP"

    def _display_frame(self, rtu_frame: bytes) -> bytes:
        """按当前通信方式返回用于显示的帧：TCP 模式转 MBAP，RTU 模式原样"""
        return rtu_to_tcp_frame(rtu_frame) if self._is_tcp() else rtu_frame

    def _resp_pdu(self, resp: bytes) -> bytes:
        """从返回报文中取出 PDU（功能码起始）：TCP 跳过 7 字节 MBAP 头，RTU 去掉 1 字节 SlaveID + 2 字节 CRC"""
        if self._is_tcp():
            return resp[7:]
        return resp[1:-2] if len(resp) >= 4 else b''

    def _format_resp(self, resp: bytes) -> str:
        """返回报文展示：RTU 附带 CRC 校验结果，TCP 直接展示原始字节"""
        if not resp:
            return "无响应（超时）"
        if self._is_tcp():
            return bytes_to_hex(resp)
        _, display = verify_response_crc(resp)
        return display

    def _check_write_resp(self, resp: bytes) -> tuple[bool, str]:
        """判断写入返回报文是否成功（异常码 / RTU 额外校验 CRC）"""
        if not resp:
            return False, "无响应（超时）"
        pdu = self._resp_pdu(resp)
        if not pdu:
            return False, f"返回报文过短: {bytes_to_hex(resp)}"
        if pdu[0] & 0x80:
            exc = pdu[1] if len(pdu) > 1 else 0
            return False, f"设备异常 0x{exc:02X}"
        if not self._is_tcp():
            ok, _ = verify_response_crc(resp)
            if not ok:
                return False, f"CRC 校验失败: {bytes_to_hex(resp)}"
        return True, bytes_to_hex(resp)

    # ─── 连接探测：校验所配 Slave ID 是否真的在线 ─────────────────────

    def _probe_slave(self, slave_id: int) -> tuple[bool, str]:
        """连接后向所配 Slave ID 发一帧读探测，判断该从机是否在线"""
        try:
            frame = build_read_frame(slave_id, self.PROBE_REG, self.PROBE_COUNT)
            resp  = self._send_recv(frame, write_sleep=0.1)
        except Exception as e:
            return False, f"探测异常: {e}"
        return self._classify_probe_resp(resp, self._is_tcp(), slave_id)

    @staticmethod
    def _classify_probe_resp(resp: bytes, is_tcp: bool, slave_id: int) -> tuple[bool, str]:
        """判定探测响应：从机任何应答（正常数据或数据级异常码）即视为在线；
        超时、RTU CRC 错、网关路由异常(0x0A/0x0B) 均视为该从机不可达"""
        if not resp:
            return False, f"从机 {slave_id} 无响应（超时）"
        if not is_tcp:
            ok, _ = verify_response_crc(resp)
            if not ok:
                return False, f"从机 {slave_id} 响应 CRC 校验失败"
        pdu = resp[7:] if is_tcp else (resp[1:-2] if len(resp) >= 4 else b'')
        if len(pdu) < 1:
            return False, f"从机 {slave_id} 响应过短"
        if pdu[0] & 0x80:
            exc = pdu[1] if len(pdu) > 1 else 0
            if exc in (0x0A, 0x0B):
                return False, f"从机 {slave_id} 不可达（网关异常 0x{exc:02X}）"
            # 其它数据级异常（非法功能/地址/值等）说明从机已应答 → 在线
        return True, ""

    def _send_recv(self, frame: bytes, write_sleep: float = 0.05) -> bytes:
        """发送 RTU 帧并返回原始响应；TCP 模式自动完成 MBAP 封装并返回原始 TCP 响应"""
        if self._is_tcp():
            return self._tcp_send_recv(frame)
        self.serial_conn.reset_input_buffer()      # 清除上次残留
        self.serial_conn.write(frame)
        time.sleep(write_sleep)
        raw = self.serial_conn.read(256)
        # 自动剥离 RS485 半双工本地回显：若收到数据以发送帧开头则丢弃前 N 字节
        if len(raw) >= len(frame) and raw[:len(frame)] == frame:
            raw = raw[len(frame):]
        return raw

    def _tcp_send_recv(self, rtu_frame: bytes) -> bytes:
        """RTU 帧 → Modbus TCP 发送，返回原始 TCP 响应（MBAP头 + PDU，无 CRC）"""
        tcp_frame = rtu_to_tcp_frame(rtu_frame)
        try:
            self.tcp_conn.sendall(tcp_frame)
            header = self._tcp_recv_exact(6)
            if len(header) < 6:
                return b''
            resp_len = struct.unpack('>H', header[4:6])[0]
            payload  = self._tcp_recv_exact(resp_len)
            if len(payload) < resp_len:
                return b''
        except socket.timeout:
            return b''
        return header + payload

    def _tcp_recv_exact(self, n: int) -> bytes:
        data = b''
        try:
            while len(data) < n:
                chunk = self.tcp_conn.recv(n - len(data))
                if not chunk:
                    break
                data += chunk
        except socket.timeout:
            pass
        return data

    # ─── 业务逻辑：写入样机配置 ────────────────────────────────────────

    def _get_field_data(self, name: str) -> tuple[int, bytes]:
        val_var, reg_var, _, dtype = self.field_vars[name]
        value    = val_var.get().strip()
        reg_addr = parse_reg_addr(reg_var.get().strip())
        if dtype == "mac":
            data = mac_to_ascii_bytes(value)
        elif dtype == "int":
            data = struct.pack('>H', int(value))
        else:
            data = str_to_padded_bytes(value)
        return reg_addr, data

    def _convert(self):
        slave_id = self._get_slave_id()
        errors   = []
        for name, (val_var, reg_var, rtu_var, _dtype) in self.field_vars.items():
            value = val_var.get().strip()
            reg   = reg_var.get().strip()
            if not value or not reg:
                rtu_var.set("")
                continue
            try:
                reg_addr, data = self._get_field_data(name)
                frame = build_write_frame(slave_id, reg_addr, data)
                rtu_var.set(bytes_to_hex(self._display_frame(frame)))
            except Exception as e:
                rtu_var.set(f"[错误] {e}")
                errors.append(f"{name}: {e}")
        if errors:
            messagebox.showerror("转换错误", '\n'.join(errors))

    def _write_all(self):
        if not self._check_connected():
            return
        slave_id      = self._get_slave_id()
        errors        = []
        success_count = 0
        skipped       = 0
        for name, (val_var, reg_var, _rtu_var, _dtype) in self.field_vars.items():
            value = val_var.get().strip()
            reg   = reg_var.get().strip()
            if not value or not reg:
                skipped += 1
                continue
            try:
                reg_addr, data = self._get_field_data(name)
                frame = build_write_frame(slave_id, reg_addr, data)
                resp  = self._send_recv(frame, write_sleep=0.05)
                ok, info = self._check_write_resp(resp)
                if ok:
                    success_count += 1
                else:
                    errors.append(f"{name}: {info}")
            except Exception as e:
                errors.append(f"{name}: {e}")
        if errors:
            messagebox.showerror("写入结果",
                                 f"成功 {success_count} 项，失败 {len(errors)} 项:\n" +
                                 '\n'.join(errors))
        else:
            msg = f"全部 {success_count} 项寄存器写入完成"
            if skipped:
                msg += f"\n（已跳过 {skipped} 项空字段）"
            messagebox.showinfo("写入成功", msg)

    # ─── 业务逻辑：读取样机配置 ────────────────────────────────────────

    def _read_all(self):
        if not self._check_connected():
            return
        slave_id = self._get_slave_id()
        for reg_var, num_var, result_var in self.read_vars:
            reg = reg_var.get().strip()
            num = num_var.get().strip()
            if not reg or not num:
                result_var.set("")
                continue
            try:
                reg_addr  = parse_reg_addr(reg)
                reg_count = int(num)
                frame = build_read_frame(slave_id, reg_addr, reg_count)
                resp  = self._send_recv(frame, write_sleep=0.1)
                result_var.set(self._format_resp(resp))
            except Exception as e:
                result_var.set(f"[错误] {e}")

    # ─── 通用功能：写寄存器 ────────────────────────────────────────────

    def _refresh_gw_frame(self, *_):
        """写入参数变化时自动刷新报文预览"""
        slave_id = self._get_slave_id()
        try:
            reg_str  = self.gw_reg_var.get().strip()
            num_str  = self.gw_num_var.get().strip()
            data_str = self.gw_data_var.get().strip()
            if not reg_str or not num_str or not data_str:
                self.gw_frame_var.set("")
                return
            reg_addr  = parse_reg_addr(reg_str)
            reg_count = int(num_str)
            value     = int(data_str)
            frame = build_write_fc16_frame(slave_id, reg_addr, reg_count, value)
            self.gw_frame_var.set(bytes_to_hex(self._display_frame(frame)))
        except Exception:
            self.gw_frame_var.set("")

    def _generic_write(self):
        slave_id = self._get_slave_id()
        try:
            reg_str  = self.gw_reg_var.get().strip()
            num_str  = self.gw_num_var.get().strip()
            data_str = self.gw_data_var.get().strip()
            if not reg_str or not num_str or not data_str:
                messagebox.showwarning("提示", "Reg / Reg num / Data 不能为空")
                return
            reg_addr  = parse_reg_addr(reg_str)
            reg_count = int(num_str)
            value     = int(data_str)
            frame = build_write_fc16_frame(slave_id, reg_addr, reg_count, value)
            self.gw_frame_var.set(bytes_to_hex(self._display_frame(frame)))
            if not self._check_connected():
                return
            resp = self._send_recv(frame)
            self.gw_resp_var.set(self._format_resp(resp))
            ok, info = self._check_write_resp(resp)
            if ok:
                messagebox.showinfo("写入成功", f"寄存器写入成功\n返回报文: {info}")
            else:
                messagebox.showwarning("写入结果", f"响应异常\n{info}")
        except Exception as e:
            self.gw_frame_var.set(f"[错误] {e}")
            messagebox.showerror("错误", str(e))

    # ─── 通用功能：读寄存器 ────────────────────────────────────────────

    def _refresh_gr_frame(self, *_):
        """读取参数变化时自动刷新发送报文预览（FC03 读帧，按当前模式格式化）"""
        slave_id = self._get_slave_id()
        for reg_var, num_var, _dtype_var, _scale_var, send_var, _frame_var, _data_var in self.gr_rows:
            reg = reg_var.get().strip()
            num = num_var.get().strip()
            if not reg or not num:
                send_var.set("")
                continue
            try:
                reg_addr  = parse_reg_addr(reg)
                reg_count = int(num)
                frame = build_read_frame(slave_id, reg_addr, reg_count)
                send_var.set(bytes_to_hex(self._display_frame(frame)))
            except Exception:
                send_var.set("")

    def _generic_read(self):
        if not self._check_connected():
            return
        slave_id = self._get_slave_id()
        for reg_var, num_var, dtype_var, scale_var, _send_var, frame_var, data_var in self.gr_rows:
            reg = reg_var.get().strip()
            num = num_var.get().strip()
            if not reg or not num:
                frame_var.set("")
                data_var.set("")
                continue
            try:
                reg_addr  = parse_reg_addr(reg)
                reg_count = int(num)
                req_frame = build_read_frame(slave_id, reg_addr, reg_count)

                resp = self._send_recv(req_frame, write_sleep=0.1)
                if not resp:
                    frame_var.set("无响应（超时）")
                    data_var.set("")
                    continue

                frame_var.set(bytes_to_hex(resp))   # 显示返回报文（按当前模式的真实字节）

                # PDU = FC(1) + byte_count(1) + data；按模式自动定位
                pdu = self._resp_pdu(resp)
                if len(pdu) < 2:
                    data_var.set("[响应过短]")
                    continue

                # 异常响应：FC | 0x80
                if pdu[0] & 0x80:
                    exc = pdu[1] if len(pdu) > 1 else 0
                    data_var.set(f"[设备异常 0x{exc:02X}]")
                    continue

                byte_count = pdu[1]
                raw = pdu[2:2 + byte_count]
                if len(raw) < byte_count:
                    data_var.set(f"[数据截断: 期望{byte_count}B 实收{len(raw)}B]")
                    continue
                data_var.set(self._parse_typed(raw, dtype_var.get(), scale_var.get()))
            except Exception as e:
                frame_var.set(f"[错误] {e}")
                data_var.set("")

    @staticmethod
    def _apply_scale(value, scale_str: str) -> str:
        """对数值应用缩放系数：data = 原始值 × scale"""
        try:
            scale = float(scale_str)
        except (TypeError, ValueError):
            scale = 1.0
        if scale == 1.0:
            return str(value)
        # round 规避浮点误差（如 574 × 0.1 = 57.400000000000006）
        return f"{round(value * scale, 6):g}"

    @staticmethod
    def _parse_typed(raw: bytes, dtype: str, scale_str: str = "1") -> str:
        """将寄存器原始字节解析为指定类型的数值字符串；数值类型按 scale 缩放"""
        try:
            if dtype == "uint16_t":
                if len(raw) >= 2:
                    return ConfigTool._apply_scale(struct.unpack('>H', raw[:2])[0], scale_str)
            elif dtype == "uint8_t":
                # 一个寄存器含 2 个字节，逐字节解析为 uint8（如 IP 段 192 168），scale 不适用
                if len(raw) >= 1:
                    return ' '.join(str(b) for b in raw)
            elif dtype == "uint32_t":
                if len(raw) >= 4:
                    return ConfigTool._apply_scale(struct.unpack('>I', raw[:4])[0], scale_str)
            elif dtype == "int32_t":
                if len(raw) >= 4:
                    return ConfigTool._apply_scale(struct.unpack('>i', raw[:4])[0], scale_str)
            elif dtype == "timestamp":
                # 32 位 Unix 时间戳（秒），按 UTC 格式化，scale 不适用
                if len(raw) >= 4:
                    ts = struct.unpack('>I', raw[:4])[0]
                    dt = datetime.fromtimestamp(ts, timezone.utc)
                    return dt.strftime('%Y-%m-%d %H:%M:%S UTC')
            elif dtype == "uint64_t":
                if len(raw) >= 8:
                    return ConfigTool._apply_scale(struct.unpack('>Q', raw[:8])[0], scale_str)
            elif dtype == "float":
                if len(raw) >= 4:
                    val = struct.unpack('>f', raw[:4])[0] * float(scale_str or "1")
                    return f"{val:.6g}"
            return "[数据不足]"
        except Exception as e:
            return f"[解析错误: {e}]"


# ─────────────────────────── 入口 ──────────────────────────────────────

def main():
    root = tk.Tk()
    root.configure(bg='#f5f5f5')
    ConfigTool(root)
    root.mainloop()


if __name__ == '__main__':
    main()
