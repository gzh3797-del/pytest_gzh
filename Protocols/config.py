# -*- coding: utf-8 -*-
"""
BACnet/IP 测试配置
所有可配项集中于此，无需修改其他模块。
"""
import os as _os
_BASE = _os.path.dirname(_os.path.abspath(__file__))

# ── 当前测试设备 ──────────────────────────────────────────────────────────────
DEVICE_NAME   = "AcuRev4100"          # 用于报告文件名
DEVICE_MODULE = "devices.acurev4100"  # 设备 Modbus 地址映射模块

# ── 网关（BACnet Server） ─────────────────────────────────────────────────────
GATEWAY_IP       = "192.168.2.159"   # 网关 IP
GATEWAY_PORT     = 49000             # 网关 BACnet/IP UDP 端口
GATEWAY_NAME     = "web2"  # Device Object Name（用于连接验证）
DEVICE_INSTANCE  = 4194302            # 设备实例号（None = 运行时自动发现）

# ── 本机（BACnet Client） ─────────────────────────────────────────────────────
LOCAL_IP   = "192.168.2.45"  # 本机网卡 IP（BAC0 监听地址）
LOCAL_PORT = 47808            # 本机 BACnet 监听端口

# ── Modbus 通信方式 ────────────────────────────────────────────────────────────
# "tcp"：通过网络直连设备（默认）
# "rtu"：通过串口 RS485 连接设备（设备与电脑不同网段时使用）
MODBUS_MODE = "tcp"

# ── Modbus RTU（串口）配置，仅 MODBUS_MODE="rtu" 时生效 ──────────────────────
MODBUS_RTU_PORT     = "COM11"   # 串口号（Windows: "COM3"，Linux: "/dev/ttyUSB0"）
MODBUS_RTU_BAUDRATE = 19200
MODBUS_RTU_PARITY   = "N"      # N=无校验  E=偶校验  O=奇校验
MODBUS_RTU_STOPBITS = 1
MODBUS_RTU_BYTESIZE = 8

# ── 下挂设备（Modbus TCP Client） ────────────────────────────────────────────
# TCP 模式：host/port/unit_id 均生效
# RTU 模式：仅 unit_id（slave 地址）生效，host/port 忽略
# 格式："设备名": ("IP", 端口, Unit ID)
MODBUS_DEVICE_MAP: dict = {
    "AcuRev4100": ("192.168.2.30", 502, 102),
    "AcuRev2100": ("192.168.2.64", 502, 101),
    "AcuvimIIW":  ("192.168.2.27",   502,   2),
    "AcuvimIIR":  ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuVIM3":    ("192.168.2.32", 502,   1),
    "AcuIOM01":   ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuIOM02":   ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuIOM03":   ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuIOM04":   ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
}
# 以下三项由运行时从 MODBUS_DEVICE_MAP 填入，无需手动修改
MODBUS_HOST = "192.168.2.10"
MODBUS_PORT = 502
MODBUS_UNIT = 102

# ── 参数模板目录（各设备 blockParams xlsx，替代 EPICS 文件） ─────────────────
TEMPLATE_DIR = _os.path.join(_BASE, "template")

# ── 报告输出目录（自动创建） ──────────────────────────────────────────────────
REPORT_DIR  = _os.path.join(_BASE, "reports")

# ── 读取参数 ──────────────────────────────────────────────────────────────────
READ_TIMEOUT   = 30   # 单次读取超时（秒）——网关性能较差时适当调大
MAX_RETRIES    = 4    # 超时重试次数
RETRY_WAIT     = 1.0  # 每次重试前等待（秒）
CONNECT_WAIT   = 3.0  # BAC0 启动后等待网关就绪（秒）
BATCH_SIZE     = 20   # 每批读取对象数（影响日志粒度）
WHOIS_TIMEOUT  = 10   # Who-Is 等待响应时间（秒）

# ── 数值比对容差（Modbus vs BACnet） ─────────────────────────────────────────
TOLERANCE_PERCENT  = 1.0   # ±1%（相对容差，用于较大值）
TOLERANCE_ABSOLUTE = 0.05  # ±0.05（绝对容差，用于接近零的值）

# ── AcuCloud 历史数据对比容差（快照 vs 实时 Modbus） ──────────────────────────
# 因存在时序差异，容差设置较大
CLOUD_TOLERANCE_PERCENT  = 5.0   # ±5%（相对容差）
CLOUD_TOLERANCE_ABSOLUTE = 1.0   # ±1.0（绝对容差）

# ── AcuCloud 数据目录 ─────────────────────────────────────────────────────────
CLOUD_DATA_DIR = _os.path.join(_BASE, "Datas", "acuclouddatas")

# ── MQTT 数据目录 ─────────────────────────────────────────────────────────────
MQTT_DATA_DIR = _os.path.join(_BASE, "MQTT")

# ── MQTT 快照比对容差（与 AcuCloud 相同，补偿时序差异） ──────────────────────
MQTT_TOLERANCE_PERCENT  = 5.0   # ±5%（相对容差）
MQTT_TOLERANCE_ABSOLUTE = 1.0   # ±1.0（绝对容差）

# ── Datalog 数据目录 ──────────────────────────────────────────────────────────
DATALOG_DATA_DIR = _os.path.join(_BASE, "Datas", "DatalogDatas")

# ── Datalog 快照比对容差 ───────────────────────────────────────────────────────
DATALOG_TOLERANCE_PERCENT  = 5.0    # ±5%（相对容差）
DATALOG_TOLERANCE_ABSOLUTE = 0.05   # ±0.05（绝对容差，与 BACnet 一致；防止小值误判通过）
