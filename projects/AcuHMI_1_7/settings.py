"""项目配置适配层：从 framework 分层配置 + .env 暴露常量，替代旧 config/settings.py。"""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

from framework.config.loader import load_config

_REPO_ROOT = Path(__file__).resolve().parents[2]
load_dotenv(_REPO_ROOT / "configs" / ".env")

_cfg = load_config("AcuHMI_1_7")


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
# 密码不留硬编码兜底（settings.py 入库），必须由 configs/.env 提供。
DEFAULT_USERNAME = os.getenv("WEB_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("WEB_PASSWORD", "")
# View（只读）账号：权限类用例使用，密码同样只由 configs/.env 提供，禁止硬编码入库。
VIEW_USERNAME = os.getenv("VIEW_USERNAME", "view")
VIEW_PASSWORD = os.getenv("VIEW_PASSWORD", "")

def get_screenshot_dir() -> Path:
    """失败截图目录（惰性求值）。

    优先用 runner / 根 conftest 注入的 SCREENSHOT_DIR（指向本次运行的 run 目录），
    否则回退到 reports/AcuHMI_1_7/_adhoc。仅在真正要截图时调用、用时才 mkdir，
    避免 import settings 就在磁盘上凭空建目录。
    """
    path = Path(os.getenv("SCREENSHOT_DIR",
                          str(_REPO_ROOT / "reports" / "AcuHMI_1_7" / "_adhoc")))
    path.mkdir(parents=True, exist_ok=True)
    return path

# ── HMI 设备常量（迁移自 test_case/AcuHMI_1_7/config.py）──
HMI_IP = _cfg.get("hmi_ip")
HMI_URL = os.getenv("HMI_URL", _cfg.get("hmi_url"))
# HMI 设备登录账号统一复用 Web UI 凭据（DEFAULT_USERNAME/DEFAULT_PASSWORD，
# 即 .env 的 WEB_USERNAME/WEB_PASSWORD），不再单独维护 HMI_USERNAME/HMI_PASSWORD。
HMI_DEVICE_NAME = _cfg.get("hmi_device_name", "Acu4100")

# ── BACnet/IP 客户端（BAC0）──
BACNET_PORT = int(_cfg.get("bacnet_port", 47808))
LOCAL_IP = _cfg.get("local_ip", "192.168.2.45")
LOCAL_PORT = int(_cfg.get("local_port", 47810))
BACNET_CONNECT_WAIT = float(_cfg.get("bacnet_connect_wait", 4.0))
BACNET_READ_TIMEOUT = float(_cfg.get("bacnet_read_timeout", 20.0))
BACNET_RESTART_WAIT = float(_cfg.get("bacnet_restart_wait", 6.0))

# ── 接线检查 / Modbus ──
METER_TCP_IP = _cfg.get("meter_tcp_ip", "192.168.2.242")
METER_TCP_PORT = int(_cfg.get("meter_tcp_port", 502))
MODBUS_SLAVE = int(_cfg.get("modbus_slave", 1))
NOMINAL_VOLTAGE = _cfg.get("nominal_voltage", 100)
NORMAL_CURRENT = _cfg.get("normal_current", 1.0)
CHECK_TIMEOUT = _cfg.get("check_timeout", 30)
MODBUS_CMP_TOLERANCE_PERCENT = float(_cfg.get("modbus_cmp_tolerance_percent", 1.0))
MODBUS_CMP_TOLERANCE_ABSOLUTE = float(_cfg.get("modbus_cmp_tolerance_absolute", 0.05))
MODBUS_CMP_TIMEOUT = float(_cfg.get("modbus_cmp_timeout", 10.0))
MODBUS_CMP_MAX_RETRIES = int(_cfg.get("modbus_cmp_max_retries", 2))

# ── BACnet 上传值 vs Modbus 实时值比对的设备直连参数 ──
# 直连各下挂设备自身真实的 Modbus TCP（每台设备各有 IP + Unit ID，非网关聚合）。
# 值为 None = 未配置 → 该设备用例自动跳过数值比对，仅保留范围一致 + 可读验证。
# 格式："设备名": ("设备真实 IP", 端口, Unit ID)
DEVICE_MODBUS_MAP = {
    "AcuRev4100": (METER_TCP_IP, METER_TCP_PORT, MODBUS_SLAVE),  # 复用接线检查的 4100 直连参数
    "AcuvimIIR": None,   # PXE1，填 ("192.168.x.x", 502, unit)
    "AcuvimIIW": ("192.168.2.27", 502, 2),   # PXE2
    "AcuRev1300": None,   # PXM350
    "AcuVIM3": ("192.168.2.32", 502, 1),
    "AcuRev2100": ("192.168.2.64", 502, 101),
}
