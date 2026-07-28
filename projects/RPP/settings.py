"""项目配置适配层：从 framework 分层配置 + .env 暴露常量，替代旧 config/settings.py。"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

from framework.config.loader import load_config

_REPO_ROOT = Path(__file__).resolve().parents[2]
load_dotenv(_REPO_ROOT / "configs" / ".env")

_cfg = load_config("RPP")


def _as_bool(value: object, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return default


BROWSER = _cfg.get("browser", "chromium")
HEADLESS = _as_bool(_cfg.get("headless"), False)
SLOW_MO = int(_cfg.get("slow_mo", 300))
TIMEOUT = int(_cfg.get("timeout", 30000))
BASE_URL = os.getenv("BASE_URL", _cfg.get("base_url"))
# RPP 真机账号密码跟 AcuHMI-1-7 不是同一套，不能共用 WEB_USERNAME/WEB_PASSWORD，
# 单独用 RPP_WEB_USERNAME/RPP_WEB_PASSWORD（同样只由 configs/.env 提供，不硬编码入库）。
DEFAULT_USERNAME = os.getenv("RPP_WEB_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("RPP_WEB_PASSWORD", "")
# View（只读）账号：权限类用例使用，密码同样只由 configs/.env 提供，禁止硬编码入库。
VIEW_USERNAME = os.getenv("VIEW_USERNAME", "view")
VIEW_PASSWORD = os.getenv("VIEW_PASSWORD", "")


def get_screenshot_dir() -> Path:
    """失败截图目录（惰性求值）。

    优先用 runner / 根 conftest 注入的 SCREENSHOT_DIR（指向本次运行的 run 目录），
    否则回退到 reports/RPP/_adhoc。仅在真正要截图时调用、用时才 mkdir，
    避免 import settings 就在磁盘上凭空建目录。
    """
    path = Path(os.getenv("SCREENSHOT_DIR",
                          str(_REPO_ROOT / "reports" / "RPP" / "_adhoc")))
    path.mkdir(parents=True, exist_ok=True)
    return path

# ── HMI 设备常量（迁移自 test_case/AcuHMI_1_7/config.py）──
HMI_IP = _cfg.get("hmi_ip")
HMI_URL = os.getenv("HMI_URL", _cfg.get("hmi_url"))
# HMI 设备登录账号统一复用 Web UI 凭据（DEFAULT_USERNAME/DEFAULT_PASSWORD，
# 即 .env 的 WEB_USERNAME/WEB_PASSWORD），不再单独维护硬编码 HMI 账号。
# 仍暴露 HMI_USERNAME/HMI_PASSWORD 别名，供 MQTT comparator 等模块引用，
# 默认回退到 Web 凭据，可由 .env 的 HMI_USERNAME/HMI_PASSWORD 单独覆盖。
HMI_USERNAME = os.getenv("HMI_USERNAME", DEFAULT_USERNAME)
HMI_PASSWORD = os.getenv("HMI_PASSWORD", DEFAULT_PASSWORD)
HMI_DEVICE_NAME = _cfg.get("hmi_device_name", "Acu4100")
# 被测表设备名（接线检查 + 直连 Modbus 测试共用）。此名同时是：
#   1) device_modbus 段查 ip/port/unit 的键名；
#   2) HMI Wiring Check 页面 Device 下拉显示名。
# 详细解析见下方 DEVICE_MODBUS_MAP 之后。
# 强制转 str：YAML 会把纯数字设备名（如 4100229）解析为 int，导致与发现/页面的字符串
# 设备名比较（d.name == WIRING_DEVICE_NAME）恒不相等、下拉 has_text 也异常，故统一为字符串。
METER_DEVICE_NAME = str(_cfg.get("meter_device_name", "Acurev4100242"))
# 兼容别名：接线检查模块仍以 WIRING_DEVICE_NAME 引用同一台表。
WIRING_DEVICE_NAME = METER_DEVICE_NAME

# ── BACnet/IP 客户端（BAC0）──
BACNET_PORT = int(_cfg.get("bacnet_port", 47808))
LOCAL_IP = _cfg.get("local_ip", "192.168.2.45")
LOCAL_PORT = int(_cfg.get("local_port", 47810))
BACNET_CONNECT_WAIT = float(_cfg.get("bacnet_connect_wait", 4.0))
BACNET_READ_TIMEOUT = float(_cfg.get("bacnet_read_timeout", 20.0))
BACNET_RESTART_WAIT = float(_cfg.get("bacnet_restart_wait", 6.0))
# 保存配置触发 BACnet 服务重启后，轮询等待服务重新可达的总预算（秒）。
# 隔离单设备 + 开启全量 Polling 后，网关重建大量对象耗时常超过固定 BACNET_RESTART_WAIT，
# 故用轮询 can_connect() 直到服务回来，避免过早读取导致超时误判为"客户端无法连接"。
BACNET_SERVICE_READY_TIMEOUT = float(_cfg.get("bacnet_service_ready_timeout", 90.0))

# ── 接线检查 / Modbus 测试参数 ──
# 直连 Modbus 目标表（METER_TCP_IP/PORT/MODBUS_SLAVE）已统一改由 device_modbus 段解析，
# 见下方 DEVICE_MODBUS_MAP 之后的「直连 Modbus 默认目标表」。
NOMINAL_VOLTAGE = _cfg.get("nominal_voltage", 100)
NORMAL_CURRENT = _cfg.get("normal_current", 1.0)
CHECK_TIMEOUT = _cfg.get("check_timeout", 30)
MODBUS_CMP_TOLERANCE_PERCENT = float(_cfg.get("modbus_cmp_tolerance_percent", 1.0))
MODBUS_CMP_TOLERANCE_ABSOLUTE = float(_cfg.get("modbus_cmp_tolerance_absolute", 0.05))
MODBUS_CMP_TIMEOUT = float(_cfg.get("modbus_cmp_timeout", 10.0))
MODBUS_CMP_MAX_RETRIES = int(_cfg.get("modbus_cmp_max_retries", 2))


def _build_device_modbus_map() -> dict[str, Optional[tuple[str, int, int]]]:
    """从 config.yaml 的 device_modbus 段构建 DEVICE_MODBUS_MAP。

    config.yaml 格式（device_modbus 段每条可为 null 或含 ip/port/unit 子键）：
        device_modbus:
          AcuRev4100:
            ip: 192.168.2.242
            port: 502
            unit: 4
          AcuvimIIR: null

    返回格式与旧硬编码一致：
        {"DeviceName": ("ip", port, unit)} 或 {"DeviceName": None}
    """
    raw: object = _cfg.get("device_modbus", {})
    result: dict[str, Optional[tuple[str, int, int]]] = {}
    if not isinstance(raw, dict):
        return result
    for dev_name, entry in raw.items():
        if entry is None:
            result[dev_name] = None
        elif isinstance(entry, dict):
            ip_val = entry.get("ip")
            port_raw = entry.get("port")
            unit_raw = entry.get("unit")
            # ip/port/unit 子键全空（如 YAML 写成 `ip:` `port:` `unit:`）等同未接入，
            # 与显式 `设备名: null` 一致按 None 处理，避免 int(None) 在导入期崩溃。
            if ip_val in (None, "") and port_raw is None and unit_raw is None:
                result[dev_name] = None
            else:
                port_val = int(port_raw) if port_raw is not None else 502
                unit_val = int(unit_raw) if unit_raw is not None else 1
                result[dev_name] = (str(ip_val or ""), port_val, unit_val)
    return result


# ── BACnet 上传值 vs Modbus 实时值比对的设备直连参数 ──
# 从 config.yaml 的 device_modbus 段读取，不再硬编码。
# 格式："设备名": ("设备真实 IP", 端口, Unit ID)，或 None（未接入 → 段4 自动跳过）。
DEVICE_MODBUS_MAP: dict[str, Optional[tuple[str, int, int]]] = _build_device_modbus_map()


def _resolve_device(name: str) -> tuple[str, int, int]:
    """按设备名从 device_modbus 段取 (ip, port, unit)。

    查无此名、或该名在 device_modbus 段被显式标记为 null（未接入）时直接抛错，
    并列出可用设备名——避免静默回退到一个错误地址，导致 fixture 深处出现难以
    定位的 Modbus 超时（曾因 meter_device_name 写错而静默连到默认表）。
    """
    if name not in DEVICE_MODBUS_MAP:
        available = ", ".join(sorted(DEVICE_MODBUS_MAP)) or "（device_modbus 段为空）"
        raise ValueError(
            f"meter_device_name='{name}' 不在 config.yaml 的 device_modbus 段中。"
            f"可用设备名：{available}。请把 meter_device_name 改为其中之一，"
            f"或在 device_modbus 段新增一条 '{name}: {{ip, port, unit}}'。")
    entry = DEVICE_MODBUS_MAP[name]
    if entry is None:
        raise ValueError(
            f"meter_device_name='{name}' 在 device_modbus 段被标记为 null（未接入），"
            f"不能作为直连 Modbus / 接线检查目标。请改指向一台已配置 ip/port/unit 的设备。")
    return entry


# ── 被测表的直连 Modbus 参数（惰性解析）──
# 按 meter_device_name 从 device_modbus 段取 ip/port/unit（name+ip+port+unit 单一来源），
# 接线检查与直连 Modbus 测试（parameter_settings/acurev4100、mqtt 数值比对等）共用这台表。
# 这些常量惰性求值：仅当真正读取时才解析，meter_device_name 写错/为 null 时立即抛出
# 可读错误（见 _resolve_device），而不再静默回退。不读这些常量的模块（如纯 BACnet
# 用例只用 DEVICE_MODBUS_MAP）导入 settings 不受影响。
# 暴露三组等价别名：
#   METER_TCP_IP / METER_TCP_PORT / MODBUS_SLAVE            —— 通用直连
#   WIRING_METER_IP / WIRING_METER_PORT / WIRING_MODBUS_SLAVE —— 接线检查模块
#   MODBUS_HOST / MODBUS_PORT / MODBUS_UNIT                  —— MQTT comparator 兼容
_METER_IP_ALIASES   = {"METER_TCP_IP", "WIRING_METER_IP", "MODBUS_HOST"}
_METER_PORT_ALIASES = {"METER_TCP_PORT", "WIRING_METER_PORT", "MODBUS_PORT"}
_METER_UNIT_ALIASES = {"MODBUS_SLAVE", "WIRING_MODBUS_SLAVE", "MODBUS_UNIT"}


def __getattr__(name: str):
    """PEP 562 惰性属性：按需解析被测表 Modbus 参数，名字错/为 null 时显式报错。"""
    if name in _METER_IP_ALIASES:
        return _resolve_device(METER_DEVICE_NAME)[0]
    if name in _METER_PORT_ALIASES:
        return _resolve_device(METER_DEVICE_NAME)[1]
    if name in _METER_UNIT_ALIASES:
        return _resolve_device(METER_DEVICE_NAME)[2]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

# ── MQTT ──────────────────────────────────────────────────────────────────────
WEB_MQTT_BROKER_ADDRESS = os.getenv("WEB_MQTT_BROKER_ADDRESS", _cfg.get("mqtt_broker_address", "www.accu.com"))
WEB_MQTT_BROKER_PORT    = int(_cfg.get("mqtt_broker_port", 1883))
MQTT_SSL_PORT           = int(_cfg.get("mqtt_ssl_port", 8883))
MQTT_BROKER_HOST        = "0.0.0.0"
MQTT_BROKER_PORT        = WEB_MQTT_BROKER_PORT
MQTT_COLLECT_TIMEOUT    = int(_cfg.get("mqtt_collect_timeout", 60))
MQTT_URL_PATH           = _cfg.get("mqtt_url_path", "/#/protocols/mqtt/deviceToPublish")
MQTT_DEFAULT_DEVICE     = _cfg.get("mqtt_default_device", "Acurev4100")
MQTT_TOLERANCE_PERCENT  = float(_cfg.get("mqtt_tolerance_percent", 1.0))
MQTT_TOLERANCE_ABSOLUTE = float(_cfg.get("mqtt_tolerance_absolute", 0.05))

# MQTT SSL 证书路径（gen_certs.py 生成后自动存入 mqtt/certs/）
_MQTT_CERT_DIR       = _REPO_ROOT / "projects" / "RPP" / "tests" / "mqtt" / "certs"
MQTT_SSL_CA_CERT     = str(_MQTT_CERT_DIR / "ca.crt")
MQTT_SSL_SERVER_CERT = str(_MQTT_CERT_DIR / "server.crt")
MQTT_SSL_SERVER_KEY  = str(_MQTT_CERT_DIR / "server.key")
MQTT_SSL_CLIENT_CERT = str(_MQTT_CERT_DIR / "client.crt")
MQTT_SSL_CLIENT_KEY  = str(_MQTT_CERT_DIR / "client.key")

# RPP 是单一产品，没有 AcuHMI-1-7 那种 WEB2/HMI17 双网关切换需求，直接暴露真机地址。
# ── MQTT comparator 兼容别名（mqtt_comparator.py 通过 config.XXX 读取） ──────
GATEWAY_WEB_URL  = HMI_URL
GATEWAY_WEB_USER = HMI_USERNAME
GATEWAY_WEB_PASS = HMI_PASSWORD
DEVICE_NAME         = MQTT_DEFAULT_DEVICE
# MODBUS_HOST / MODBUS_PORT / MODBUS_UNIT 由模块级 __getattr__ 惰性提供（见上）。
MODBUS_DEVICE_MAP = DEVICE_MODBUS_MAP
TEMPLATE_DIR     = str(_REPO_ROOT / "knowledge" / "shared" / "templates" / "raw")
REPORT_DIR       = str(_REPO_ROOT / "reports" / "RPP" / "mqtt")
