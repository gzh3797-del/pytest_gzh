"""
样机配置写入工具
依赖: pip install pyserial
"""

import tkinter as tk
from tkinter import ttk, messagebox
import serial
import serial.tools.list_ports
import struct
import time


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
    """MAC地址转12字节ASCII数据，如 '30:7A:57:01:F5:50' -> b'307A5701F550'"""
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

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("样机配置写入工具")
        self.root.resizable(True, True)
        self.serial_conn: serial.Serial | None = None
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

        tk.Label(row, text="波特率").pack(side='left')
        self.baud_var = tk.StringVar(value="19200")
        ttk.Combobox(row, textvariable=self.baud_var, width=8,
                     values=["9600", "19200", "38400", "57600", "115200"],
                     state='readonly').pack(side='left', padx=(4, 12))

        tk.Label(row, text="Com").pack(side='left')
        self.port_var = tk.StringVar(value="COM4")
        self.port_cb = ttk.Combobox(row, textvariable=self.port_var, width=9)
        self.port_cb.pack(side='left', padx=4)
        tk.Button(row, text="⟳", width=2, command=self._refresh_ports).pack(side='left', padx=(0, 12))

        tk.Label(row, text="Slave ID").pack(side='left')
        self.slave_var = tk.StringVar(value="1")
        tk.Entry(row, textvariable=self.slave_var, width=5).pack(side='left', padx=(4, 12))

        tk.Label(row, text="parity").pack(side='left')
        self.parity_var = tk.StringVar(value="none")
        ttk.Combobox(row, textvariable=self.parity_var, width=6,
                     values=["none", "even", "odd"], state='readonly').pack(side='left', padx=(4, 20))

        tk.Button(row, text="连接", width=6, bg='#4CAF50', fg='white',
                  activebackground='#388E3C', command=self._connect).pack(side='left', padx=2)
        tk.Button(row, text="断开", width=6, bg='#f44336', fg='white',
                  activebackground='#C62828', command=self._disconnect).pack(side='left', padx=2)

        self.conn_label = tk.Label(row, text="● 未连接", fg='#c0392b', font=('', 9, 'bold'))
        self.conn_label.pack(side='left', padx=8)

        self._refresh_ports()

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

            tk.Label(row, text="RTU报文").pack(side='left', padx=(8, 2))
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

        # ── 写入行 ──
        write_row = tk.Frame(frame)
        write_row.pack(fill='x', pady=2)

        tk.Label(write_row, text="Data").pack(side='left')
        self.generic_data_var = tk.StringVar(value="1")
        tk.Entry(write_row, textvariable=self.generic_data_var, width=6).pack(side='left', padx=(4, 8))

        tk.Label(write_row, text="Reg(hex)").pack(side='left')
        self.generic_write_reg_var = tk.StringVar(value="1004")
        tk.Entry(write_row, textvariable=self.generic_write_reg_var, width=8).pack(side='left', padx=4)

        tk.Label(write_row, text="Reg num").pack(side='left', padx=(8, 2))
        self.generic_write_num_var = tk.StringVar(value="1")
        tk.Entry(write_row, textvariable=self.generic_write_num_var, width=5).pack(side='left', padx=4)

        tk.Label(write_row, text="RTU报文").pack(side='left', padx=(8, 2))
        self.generic_write_rtu_var = tk.StringVar()
        tk.Entry(write_row, textvariable=self.generic_write_rtu_var, width=50,
                 state='readonly', readonlybackground='#fafafa',
                 font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

        tk.Button(write_row, text="写入寄存器", width=12, bg='#FF9800', fg='white',
                  activebackground='#E65100', font=('', 9, 'bold'),
                  command=self._generic_write).pack(side='left', padx=4)

        # ── 读取行 ──
        read_row = tk.Frame(frame)
        read_row.pack(fill='x', pady=2)

        # 空白占位，与写入行的 Data 标签 + 输入框对齐
        tk.Label(read_row, text=" " * 12).pack(side='left')

        tk.Label(read_row, text="Reg(hex)").pack(side='left')
        self.generic_read_reg_var = tk.StringVar(value="1004")
        tk.Entry(read_row, textvariable=self.generic_read_reg_var, width=8).pack(side='left', padx=4)

        tk.Label(read_row, text="Reg num").pack(side='left', padx=(8, 2))
        self.generic_read_num_var = tk.StringVar(value="1")
        tk.Entry(read_row, textvariable=self.generic_read_num_var, width=5).pack(side='left', padx=4)

        tk.Label(read_row, text="返回报文").pack(side='left', padx=(8, 2))
        self.generic_read_result_var = tk.StringVar()
        tk.Entry(read_row, textvariable=self.generic_read_result_var, width=50,
                 readonlybackground='#fffde7', state='readonly',
                 font=('Courier New', 9)).pack(side='left', padx=4, fill='x', expand=True)

        tk.Button(read_row, text="读取寄存器", width=12, bg='#9C27B0', fg='white',
                  activebackground='#6A1B9A', font=('', 9, 'bold'),
                  command=self._generic_read).pack(side='left', padx=4)

        # 绑定输入变化自动刷新 RTU 预览
        for var in (self.generic_data_var, self.generic_write_reg_var, self.generic_write_num_var):
            var.trace_add('write', self._refresh_generic_write_frame)
        self._refresh_generic_write_frame()

    # ─── 串口操作 ──────────────────────────────────────────────────────

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_cb['values'] = ports if ports else ["COM1"]
        if ports and self.port_var.get() not in ports:
            self.port_var.set(ports[0])

    def _connect(self):
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
            self.conn_label.config(text=f"● 已连接 {self.port_var.get()}", fg='#27ae60')
        except Exception as e:
            messagebox.showerror("连接失败", str(e))

    def _disconnect(self):
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.serial_conn = None
        self.conn_label.config(text="● 未连接", fg='#c0392b')

    def _check_connected(self) -> bool:
        if not self.serial_conn or not self.serial_conn.is_open:
            messagebox.showwarning("提示", "请先连接串口")
            return False
        return True

    def _get_slave_id(self) -> int:
        try:
            return int(self.slave_var.get())
        except ValueError:
            return 1

    # ─── 业务逻辑 ──────────────────────────────────────────────────────

    def _get_field_data(self, name: str) -> tuple[int, bytes]:
        """返回 (reg_addr, data_bytes)，按 dtype 编码"""
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
        """转换：生成所有 RTU 报文（含 CRC，仅供查看）；数据或寄存器为空时跳过"""
        slave_id = self._get_slave_id()
        errors   = []

        for name, (val_var, reg_var, rtu_var, dtype) in self.field_vars.items():
            value = val_var.get().strip()
            reg   = reg_var.get().strip()

            if not value or not reg:
                rtu_var.set("")
                continue

            try:
                reg_addr, data = self._get_field_data(name)
                frame = build_write_frame(slave_id, reg_addr, data)
                rtu_var.set(bytes_to_hex(frame))
            except Exception as e:
                rtu_var.set(f"[错误] {e}")
                errors.append(f"{name}: {e}")

        if errors:
            messagebox.showerror("转换错误", '\n'.join(errors))

    def _write_all(self):
        """写入所有寄存器；数据或寄存器地址为空的条目自动跳过"""
        if not self._check_connected():
            return

        slave_id      = self._get_slave_id()
        errors        = []
        success_count = 0
        skipped       = 0

        for name, (val_var, reg_var, rtu_var, dtype) in self.field_vars.items():
            value = val_var.get().strip()
            reg   = reg_var.get().strip()

            if not value or not reg:
                skipped += 1
                continue

            try:
                reg_addr, data = self._get_field_data(name)
                frame = build_write_frame(slave_id, reg_addr, data)
                self.serial_conn.write(frame)
                time.sleep(0.05)
                resp = self.serial_conn.read(256)
                if not resp:
                    errors.append(f"{name}: 无响应")
                else:
                    success_count += 1
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

    def _read_all(self):
        """读取所有寄存器；Reg 或 Reg num 为空的条目自动跳过"""
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
                self.serial_conn.write(frame)
                time.sleep(0.1)

                expected_len = 3 + reg_count * 2 + 2
                resp = self.serial_conn.read(expected_len)

                if resp:
                    _, display = verify_response_crc(resp)
                    result_var.set(display)
                else:
                    result_var.set("无响应（超时）")

            except Exception as e:
                result_var.set(f"[错误] {e}")

    # ─── 通用功能 ──────────────────────────────────────────────────────

    def _refresh_generic_write_frame(self, *_):
        """输入变化时自动刷新通用写入 RTU 报文预览"""
        slave_id = self._get_slave_id()
        try:
            reg_str  = self.generic_write_reg_var.get().strip()
            num_str  = self.generic_write_num_var.get().strip()
            data_str = self.generic_data_var.get().strip()

            if not reg_str or not num_str or not data_str:
                self.generic_write_rtu_var.set("")
                return

            reg_addr  = parse_reg_addr(reg_str)
            reg_count = int(num_str)
            value     = int(data_str)

            frame = build_write_fc16_frame(slave_id, reg_addr, reg_count, value)
            self.generic_write_rtu_var.set(bytes_to_hex(frame))
        except Exception:
            self.generic_write_rtu_var.set("")

    def _generic_write(self):
        """通用写入：FC16，Data 为 16 位整数；字段为空时不发送"""
        slave_id = self._get_slave_id()
        try:
            reg_str  = self.generic_write_reg_var.get().strip()
            num_str  = self.generic_write_num_var.get().strip()
            data_str = self.generic_data_var.get().strip()

            if not reg_str or not num_str or not data_str:
                self.generic_write_rtu_var.set("[字段为空，已跳过]")
                return

            reg_addr  = parse_reg_addr(reg_str)
            reg_count = int(num_str)
            value     = int(data_str)

            frame = build_write_fc16_frame(slave_id, reg_addr, reg_count, value)
            self.generic_write_rtu_var.set(bytes_to_hex(frame))

            if not self._check_connected():
                return

            self.serial_conn.write(frame)
            time.sleep(0.05)
            resp = self.serial_conn.read(256)
            if not resp:
                messagebox.showwarning("写入结果", "无响应（超时）")
            else:
                messagebox.showinfo("写入结果", f"响应: {bytes_to_hex(resp)}")

        except Exception as e:
            self.generic_write_rtu_var.set(f"[错误] {e}")
            messagebox.showerror("错误", str(e))

    def _generic_read(self):
        """通用读取：FC03；Reg 或 Reg num 为空时不发送"""
        if not self._check_connected():
            return

        slave_id = self._get_slave_id()
        try:
            reg_str = self.generic_read_reg_var.get().strip()
            num_str = self.generic_read_num_var.get().strip()

            if not reg_str or not num_str:
                self.generic_read_result_var.set("[字段为空，已跳过]")
                return

            reg_addr  = parse_reg_addr(reg_str)
            reg_count = int(num_str)

            frame = build_read_frame(slave_id, reg_addr, reg_count)
            self.serial_conn.write(frame)
            time.sleep(0.1)

            expected_len = 3 + reg_count * 2 + 2
            resp = self.serial_conn.read(expected_len)

            if resp:
                _, display = verify_response_crc(resp)
                self.generic_read_result_var.set(display)
            else:
                self.generic_read_result_var.set("无响应（超时）")

        except Exception as e:
            self.generic_read_result_var.set(f"[错误] {e}")


# ─────────────────────────── 入口 ──────────────────────────────────────

def main():
    root = tk.Tk()
    root.configure(bg='#f5f5f5')
    app = ConfigTool(root)
    root.mainloop()


if __name__ == '__main__':
    main()
