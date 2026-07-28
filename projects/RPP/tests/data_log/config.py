# -*- coding: utf-8 -*-
from pathlib import Path
import yaml

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent.parent.parent  # AcuHMI-1-7/

_cfg = yaml.safe_load((_HERE / "config.yaml").read_text(encoding="utf-8"))

DEVICE_NAME   = _cfg["datalog"].get("default_device", "AcuRev4100")
DEVICE_MODULE = f"devices.{DEVICE_NAME.lower()}"

TEMPLATE_DIR = str(_PROJECT_ROOT / "knowledge" / "shared" / "templates" / "raw")
REPORT_DIR   = str(_HERE / "reports")

_modbus = _cfg.get("modbus", {})

# ── 统一设备映射：{设备名: {"mode": "tcp"|"rtu", ...连接参数}} ────────────────
# TCP 设备：{"mode":"tcp", "ip":str, "port":int, "unit":int}
# RTU 设备：{"mode":"rtu", "serial_port":str, "baudrate":int, "parity":str,
#            "stopbits":int, "bytesize":int, "unit":int}
MODBUS_DEVICE_MAP: dict = {}

for _name, _dev in _modbus.get("tcp", {}).get("devices", {}).items():
    MODBUS_DEVICE_MAP[_name] = {
        "mode":  "tcp",
        "ip":    _dev["ip"],
        "port":  int(_dev["port"]),
        "unit":  int(_dev["unit"]),
    }

# rtu 为列表，每项对应一条独立串口线路
for _line in _modbus.get("rtu", []):
    _serial_params = {
        "serial_port": _line["port"],
        "baudrate":    int(_line.get("baudrate", 9600)),
        "parity":      _line.get("parity",   "N"),
        "stopbits":    int(_line.get("stopbits", 1)),
        "bytesize":    int(_line.get("bytesize", 8)),
    }
    for _name, _dev in _line.get("devices", {}).items():
        MODBUS_DEVICE_MAP[_name] = {
            "mode": "rtu",
            "unit": int(_dev["unit"]),
            **_serial_params,
        }

# 兼容旧代码：暴露第一条 RTU 线路的参数作为默认值
_first_rtu = next(iter(_modbus.get("rtu", [])), {})
MODBUS_RTU_PORT     = _first_rtu.get("port",     "COM6")
MODBUS_RTU_BAUDRATE = int(_first_rtu.get("baudrate", 19200))
MODBUS_RTU_PARITY   = _first_rtu.get("parity",   "N")
MODBUS_RTU_STOPBITS = int(_first_rtu.get("stopbits", 1))
MODBUS_RTU_BYTESIZE = int(_first_rtu.get("bytesize", 8))

READ_TIMEOUT  = 30
MAX_RETRIES   = 4
RETRY_WAIT    = 1.0
BATCH_SIZE    = 20

DATALOG_DATA_DIR           = str(_HERE / "data")
DATALOG_TOLERANCE_PERCENT  = float(_cfg["datalog"]["tolerance_pct"])
DATALOG_TOLERANCE_ABSOLUTE = float(_cfg["datalog"]["tolerance_abs"])
DATALOG_PUSH_TIMEOUT       = int(_cfg["datalog"]["push_timeout"])

DATALOG_SERVER_HOST = _cfg["server"]["host"]
DATALOG_FTP_PORT  = int(_cfg["server"]["ftp"]["port"])
DATALOG_FTP_USER  = _cfg["server"]["ftp"]["user"]
DATALOG_FTP_PASS  = _cfg["server"]["ftp"]["password"]
DATALOG_SFTP_PORT = int(_cfg["server"]["sftp"]["port"])
DATALOG_SFTP_USER = _cfg["server"]["sftp"]["user"]
DATALOG_SFTP_PASS = _cfg["server"]["sftp"]["password"]
DATALOG_HTTP_PORT  = int(_cfg["server"]["http"]["port"])
DATALOG_HTTPS_PORT = int(_cfg["server"]["https"]["port"])
DATALOG_SSL_CERT   = _cfg["server"]["https"].get("cert", "")
DATALOG_SSL_KEY    = _cfg["server"]["https"].get("key", "")

GATEWAY_WEB_URL  = _cfg["gateway"]["url"]
GATEWAY_WEB_USER = _cfg["gateway"]["username"]
GATEWAY_WEB_PASS = _cfg["gateway"]["password"]

# 运行时由 conftest driver fixture 填充：网关动态发现的下挂 Modbus TCP 设备列表
DISCOVERED_DEVICES: list = []
