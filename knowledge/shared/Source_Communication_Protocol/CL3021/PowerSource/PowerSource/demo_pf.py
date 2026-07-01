"""
CL3021 功率因数演示: PF=1.0 / PF=0.8 / PF=0.5
运行: python demo_pf.py
"""

from test_cl3021_com18 import *
import serial, time, struct, math, ctypes
import serial.win32 as sw

ser = serial.Serial('COM18', baudrate=9600, bytesize=8, parity='N', stopbits=1,
                    timeout=1.0, write_timeout=5.0, rtscts=False, dsrdtr=False)

# 强制禁用 CTS 流控
dcb = sw.DCB()
ctypes.windll.kernel32.GetCommState(int(ser._port_handle), ctypes.byref(dcb))
dcb.fOutxCtsFlow = 0
dcb.fRtsControl  = 2   # RTS_CONTROL_ENABLE
ctypes.windll.kernel32.SetCommState(int(ser._port_handle), ctypes.byref(dcb))
ser.setRTS(True)
ser.setDTR(True)
ctypes.windll.kernel32.PurgeComm(int(ser._port_handle), 0x0F)


def send(frame, wait=0.4):
    ser.reset_input_buffer()
    ser.write(frame)
    time.sleep(wait)
    n = ser.in_waiting
    return ser.read(n) if n else b''


def set_output(ua, ia, freq, pf):
    theta = math.degrees(math.acos(max(0.0, min(1.0, pf))))
    send(build_angle_update(0.0, 240.0, 120.0,
                            theta, theta, theta, freq), 0.35)
    send(build_amplitude_update(ua, ua, ua, ia, ia, ia), 0.35)
    send(build_freq_update(freq), 0.35)


def read_meas():
    rx = send(build_read_command(), 0.9)
    if len(rx) < 42:
        print(f'  读取失败 ({len(rx)}B)')
        return
    o = 8
    uc  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    ub  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    ua  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    ic  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    ib  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    ia  = struct.unpack_from('<i', rx, o)[0] / 1e6; o += 5
    freq = struct.unpack_from('<i', rx, o)[0] / 1e4
    print(f'  测量: UA={ua:6.2f}V  IA={ia:.4f}A  f={freq:.1f}Hz')


print('初始化...')
send(build_connect(), 0.8)
send(build_ac_screen(), 0.5)
send(build_set_line(), 0.5)

UA = 50.0
IA = 1.0
FREQ = 50.0

cases = [
    (1.0, 'PF=1.0  θ=  0.0°  纯阻性，电流与电压同相'),
    (0.8, 'PF=0.8  θ= 36.9°  感性负载'),
    (0.5, 'PF=0.5  θ= 60.0°  感性负载'),
]

for pf, label in cases:
    theta = math.degrees(math.acos(pf))
    print(f'\n=== {label} ===')
    print(f'  设定: {UA}V / {IA}A / {FREQ}Hz / 电流滞后 {theta:.1f}°')
    set_output(UA, IA, FREQ, pf)
    set_output(UA, IA, FREQ, pf)
    time.sleep(1.5)
    read_meas()
    print('  观察设备面板 (停留5秒)...')
    time.sleep(5.0)

print('\n=== 归零并退出 ===')
set_output(0, 0, FREQ, 1.0)
send(build_exit_control(), 0.5)
ser.close()
print('完成')
