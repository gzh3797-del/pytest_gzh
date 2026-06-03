# -*- coding: utf-8 -*-
"""
config.py — 网关协议测试统一配置

结构：
  1. 当前测试设备
  2. 共享路径
  3. Modbus（各协议共用的下行读取层）
  4. BACnet/IP
  5. EtherNet/IP
  6. AcuCloud
  7. MQTT
  8. 设备网页自动化（Playwright）
  9. Datalog
"""
import os as _os
_BASE = _os.path.dirname(_os.path.abspath(__file__))


# ══════════════════════════════════════════════════════════════════════════════
# 1. 当前测试设备
# ══════════════════════════════════════════════════════════════════════════════

DEVICE_NAME   = "AcuRev4100"          # 用于报告文件名
DEVICE_MODULE = "devices.acurev4100"  # 设备 Modbus 地址映射模块


# ══════════════════════════════════════════════════════════════════════════════
# 2. 共享路径
# ══════════════════════════════════════════════════════════════════════════════

# 主参数模板目录（各设备 blockParams xlsx）
TEMPLATE_DIR = _os.path.join(_BASE, "..", "knowledge", "shared", "templates", "raw")

# 报告输出目录（自动创建）
REPORT_DIR = _os.path.join(_BASE, "reports")


# ══════════════════════════════════════════════════════════════════════════════
# 3. Modbus（所有协议的下行读取层）
# ══════════════════════════════════════════════════════════════════════════════

# 通信方式："tcp" 直连设备 | "rtu" 串口 RS485（设备与电脑不同网段时使用）
MODBUS_MODE = "tcp"

# ── RTU（串口）参数，仅 MODBUS_MODE="rtu" 时生效 ──────────────────────────
MODBUS_RTU_PORT     = "COM11"   # Windows: "COM3"，Linux: "/dev/ttyUSB0"
MODBUS_RTU_BAUDRATE = 19200
MODBUS_RTU_PARITY   = "N"       # N=无校验  E=偶校验  O=奇校验
MODBUS_RTU_STOPBITS = 1
MODBUS_RTU_BYTESIZE = 8

# ── 下挂设备地址表 ────────────────────────────────────────────────────────────
# TCP 模式：host / port / unit_id 均生效
# RTU 模式：仅 unit_id（slave 地址）生效，host / port 忽略
# 格式："设备名": ("IP", 端口, Unit ID)
MODBUS_DEVICE_MAP: dict = {
    "AcuRev4100": ("192.168.2.242", 502,   1),
    "AcuRev2100": ("192.168.2.64",  502, 101),
    "AcuvimIIW":  ("192.168.2.27",  502,   2),
    "AcuvimIIR":  ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuVIM3":    ("192.168.2.32",  502,   1),
    "AcuRev1300": ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit（AcuRev1300/PXM350）
    "AcuIOM01":   ("192.168.2.10",  502,   6),  # ← 请填写实际 IP/Unit
    "AcuIOM02":   ("192.168.2.10",  502,   1),  # ← 请填写实际 IP/Unit
    "AcuIOM03":   ("192.168.2.10",  502,   4),  # ← 请填写实际 IP/Unit
    "AcuIOM04":   ("192.168.2.10",  502,   2),  # ← 请填写实际 IP/Unit
}

# 以下三项由运行时从 MODBUS_DEVICE_MAP 填入，无需手动修改
MODBUS_HOST = "192.168.2.10"
MODBUS_PORT = 502
MODBUS_UNIT = 102

# ── 读取参数（Modbus + BACnet 共用） ─────────────────────────────────────────
READ_TIMEOUT  = 30    # 单次读取超时（秒）——网关性能较差时适当调大
MAX_RETRIES   = 4     # 超时重试次数
RETRY_WAIT    = 1.0   # 每次重试前等待（秒）
BATCH_SIZE    = 20    # 每批读取对象数（影响日志粒度）


# ══════════════════════════════════════════════════════════════════════════════
# 4. BACnet/IP
# ══════════════════════════════════════════════════════════════════════════════

# ── 网关（BACnet Server） ─────────────────────────────────────────────────────
GATEWAY_IP       = "192.168.3.9"    # 网关 IP
GATEWAY_PORT     = 49000            # 网关 BACnet/IP UDP 端口
GATEWAY_NAME     = "web2"           # Device Object Name（用于连接验证）
DEVICE_INSTANCE  = 4194302          # 设备实例号（None = 运行时自动发现）

# ── 本机（BACnet Client） ─────────────────────────────────────────────────────
LOCAL_IP   = "192.168.2.45"  # 本机网卡 IP（BAC0 监听地址）
LOCAL_PORT = 47808           # 本机 BACnet 监听端口

# ── BACnet 读取专用参数 ───────────────────────────────────────────────────────
CONNECT_WAIT  = 3.0   # BAC0 启动后等待网关就绪（秒）
WHOIS_TIMEOUT = 10    # Who-Is 等待响应时间（秒）

# ── BACnet 参数范围过滤标记（AcuIOM 专用） ───────────────────────────────────
# 空字符串：使用主模板 BACnetIP 列过滤（默认）
# "8" ：AcuIOM-01/02（AI 模拟量通道）
# "10"：AcuIOM-03/04（DI 数字量通道）
BACNET_RANGE_MARKER = ""

# ── BACnet vs Modbus 数值比对容差 ─────────────────────────────────────────────
TOLERANCE_PERCENT  = 1.0    # ±1%（相对容差）
TOLERANCE_ABSOLUTE = 0.05   # ±0.05（绝对容差，防止接近零的值误判）


# ══════════════════════════════════════════════════════════════════════════════
# 5. EtherNet/IP
# ══════════════════════════════════════════════════════════════════════════════

ENIP_HOST = "192.168.3.9"                              # EIP 网关 IP（端口固定 44818）
ENIP_SLOT = 0                                          # CIP slot（通常为 0）
EDS_DIR   = _os.path.join(_BASE, "EtherNetIP", "eds")  # EDS 文件目录（单表模式回退查找用）

# EDS 文件路径（网关配置快照，每次测试按实际 EDS 填写）
# --all 多表模式必须填写；空字符串时回退到 EDS_DIR 按设备名模糊匹配（单表模式兼容）
ENIP_EDS_PATH: str = _os.path.join(_BASE, "EtherNetIP", "eds", "AcuRev-4100-WEB2.eds")

# EtherNet/IP vs Modbus 数值比对容差（与 BACnet 相同）
ENIP_TOLERANCE_PERCENT  = 1.0
ENIP_TOLERANCE_ABSOLUTE = 0.05


# ── EtherNet/IP 多表测试（--all 模式） ────────────────────────────────────────
# 格式：(eds_label, device_name, modbus_host, modbus_port, modbus_unit)
#   eds_label   — 设备 SN（即 EDS ConnectionN 的 help string，web2 生成 EDS 时固定写入 SN）
#                 脚本据此在 EDS 中找到该设备对应的所有 Assembly 实例，实现参数隔离
#   device_name — 设备型号，对应 devices/ 模块与参数模板（如 "AcuRev4100"、"AcuIOM01"）
#   modbus_host — 各设备自身的 Modbus TCP IP（每台不同）
#   modbus_port — Modbus TCP 端口（4100 非标准端口为 2000，请按实际填写）
#   modbus_unit — Modbus Unit ID（从设备网页配置获取）
# 注：EIP 网关地址统一使用 ENIP_HOST，无需在此重复填写
ENIP_MULTI_DEVICES: list[tuple] = [
    # ("41002242", "AcuRev4100", "192.168.2.242", 502, 1),
    ("4100229",  "AcuRev4100", "192.168.2.29",  502, 202),
    # ("4100150138", "AcuRev4100", "192.168.150.138", 502, 1),
]



# ══════════════════════════════════════════════════════════════════════════════
# 6. AcuCloud
# ══════════════════════════════════════════════════════════════════════════════

# AcuCloud 专属模板目录（含 paramType_AcuCloud 列，为参数范围基准）
ACUCLOUD_TEMPLATE_DIR = _os.path.join(
    _BASE, "..", "knowledge", "shared", "templates", "raw", "AcuCloud 模板适配"
)

# AcuCloud 快照数据目录
CLOUD_DATA_DIR = _os.path.join(_BASE, "Datas", "acuclouddatas")

# AcuCloud 快照 vs 实时 Modbus 容差（时序差异较大，容差设宽）
CLOUD_TOLERANCE_PERCENT  = 5.0   # ±5%（相对容差）
CLOUD_TOLERANCE_ABSOLUTE = 1.0   # ±1.0（绝对容差）


# ══════════════════════════════════════════════════════════════════════════════
# 7. MQTT
# ══════════════════════════════════════════════════════════════════════════════

# MQTT 快照数据目录
MQTT_DATA_DIR = _os.path.join(_BASE, "MQTT")

# MQTT 快照 vs 实时 Modbus 容差
MQTT_TOLERANCE_PERCENT  = 5.0   # ±5%（相对容差）
MQTT_TOLERANCE_ABSOLUTE = 1.0   # ±1.0（绝对容差）


# ── 设备侧 Broker 地址（填入设备网页的地址，与本机监听地址分开） ─────────────
# 设备连接的目标 Broker 地址（如 www.accu.com 或本机 IP）
WEB_MQTT_BROKER_ADDRESS = "www.accu.com"
WEB_MQTT_BROKER_PORT    = 1883

# MQTT 内嵌 Broker（实时采集模式，--live）
MQTT_BROKER_HOST     = "0.0.0.0"  # Broker 监听地址（0.0.0.0 允许外部设备连入）
MQTT_BROKER_PORT     = 1883       # MQTT 端口
MQTT_COLLECT_TIMEOUT = 60         # 等待设备数据的秒数
MQTT_TOPIC           = "#"        # 订阅 Topic

# MQTT SSL / mTLS 证书路径（--ssl 模式使用）
MQTT_SSL_CERT_DIR    = _os.path.join(_BASE, "MQTT", "certs")
MQTT_SSL_CA_CERT     = _os.path.join(MQTT_SSL_CERT_DIR, "ca.crt")
MQTT_SSL_SERVER_CERT = _os.path.join(MQTT_SSL_CERT_DIR, "server.crt")
MQTT_SSL_SERVER_KEY  = _os.path.join(MQTT_SSL_CERT_DIR, "server.key")
MQTT_SSL_CLIENT_CERT = _os.path.join(MQTT_SSL_CERT_DIR, "client.crt")
MQTT_SSL_CLIENT_KEY  = _os.path.join(MQTT_SSL_CERT_DIR, "client.key")
MQTT_SSL_PORT        = 8883


# ══════════════════════════════════════════════════════════════════════════════
# 8. 设备网页自动化（Playwright）
# ══════════════════════════════════════════════════════════════════════════════

WEB_URL      = "http://192.168.2.9"   # 设备网页地址
WEB_USERNAME = "admin"                # 登录用户名
WEB_PASSWORD = "Admin@110001"         # 登录密码

# 保存配置后等待设备重连的时间（秒）
WEB_RECONNECT_WAIT = 20

# 是否无头模式运行浏览器（True=后台，False=可见，调试时设 False）
WEB_HEADLESS = False


# ══════════════════════════════════════════════════════════════════════════════
# 9. Datalog
# ══════════════════════════════════════════════════════════════════════════════

# Datalog 快照数据目录
DATALOG_DATA_DIR = _os.path.join(_BASE, "Datas", "DatalogDatas")

# Datalog 快照 vs 实时 Modbus 容差
# 判断逻辑：max(绝对容差, 相对容差)——IEC/ANSI 惯例
DATALOG_TOLERANCE_PERCENT  = 5.0    # ±5%（相对容差）
DATALOG_TOLERANCE_ABSOLUTE = 0.05   # ±0.05（绝对容差下限，防止极小值误判）


# ══════════════════════════════════════════════════════════════════════════════
# 10. Datalog 推送服务器（Post Channel 验证用）
# ══════════════════════════════════════════════════════════════════════════════

# 本机 IP（网关需能访问此地址）
DATALOG_SERVER_HOST = "192.168.2.149"

# ── FTP ───────────────────────────────────────────────────────────────────────
DATALOG_FTP_PORT = 21
DATALOG_FTP_USER = "datalog"
DATALOG_FTP_PASS = "datalog123"

# ── SFTP ──────────────────────────────────────────────────────────────────────
DATALOG_SFTP_PORT = 22
DATALOG_SFTP_USER = "datalog"
DATALOG_SFTP_PASS = "datalog123"

# ── HTTP / HTTPS ──────────────────────────────────────────────────────────────
DATALOG_HTTP_PORT  = 8080
DATALOG_HTTPS_PORT = 8443
# HTTPS 自签名证书（留空则使用 HTTP）
DATALOG_SSL_CERT = ""
DATALOG_SSL_KEY  = ""

# ── 网关 Web 登录（Selenium 用） ──────────────────────────────────────────────
GATEWAY_WEB_URL  = "https://192.168.2.9/#/login"
GATEWAY_WEB_USER = "admin"
GATEWAY_WEB_PASS = "Admin@110001"

# ── 等待文件推送的超时时间（秒） ──────────────────────────────────────────────
DATALOG_PUSH_TIMEOUT = 300
