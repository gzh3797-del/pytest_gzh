# ── Wiring Check 寄存器地址 ──────────────────────────────────────────────────
# Basic Setting
REG_SERVICE_CONFIG      = 0x1042   # R/W uint16: 0=1E2W 1=2E3W1P 2=2E3WDelta 3=2E3WNet 4=3E4WY
REG_PHASE_ORDER         = 0x10DC   # R/W uint16: 0=ABC 1=ACB

# Waveform/PQ — Nominal Voltage (uint32, 两个不连续寄存器)
REG_NOMINAL_VOLTAGE_L   = 0x6501   # R/W uint16: low  word
REG_NOMINAL_VOLTAGE_H   = 0x651A   # R/W uint16: high word

# Wiring Check 控制 / 状态
REG_WIRE_CHECK_START    = 0x1300   # R/W uint16: 0=Idle 1=Start
REG_WIRE_CHECK_STATUS   = 0x1301   # R   uint16: 0=NotStarted 1=InProgress 2=Completed

# 结果 — 电压 Error Code
REG_VOLTAGE_ERROR       = 0x1302   # R   uint16: 9 bits，见下方位定义

# 结果 — 电流 Error Code，User1=0x1303 … User12=0x130E
REG_CURRENT_ERROR_BASE  = 0x1303   # R   uint16 × 12
NUM_USER_CHANNELS       = 12

# ── Service Configuration 取值 ───────────────────────────────────────────────
SERVICE_1E2W        = 0
SERVICE_2E3W_1PHASE = 1
SERVICE_2E3W_DELTA  = 2
SERVICE_2E3W_NET    = 3
SERVICE_3E4WY       = 4

# ── Phase Order 取值 ─────────────────────────────────────────────────────────
PHASE_ABC = 0
PHASE_ACB = 1

# ── 电压 Error Code 位定义 ───────────────────────────────────────────────────
V_BIT_VA_MISSING        = 0   # Va Wiring Missing
V_BIT_VB_MISSING        = 1   # Vb Wiring Missing
V_BIT_VC_MISSING        = 2   # Vc Wiring Missing
V_BIT_VA_REVERSED       = 3   # Va-Vn Reversed
V_BIT_VB_REVERSED       = 4   # Vb-Vn Reversed
V_BIT_VC_REVERSED       = 5   # Vc-Vn Reversed
V_BIT_VB_PHASE_SHIFT    = 6   # Vb Phase Shift
V_BIT_VC_PHASE_SHIFT    = 7   # Vc Phase Shift
V_BIT_PHASE_ORDER_ERR   = 8   # Phase Order Error

# ── 电流 Error Code 位定义（3E4WY / 2E3W1P / Delta / Network）───────────────
I_BIT_IA_MISSING        = 0
I_BIT_IB_MISSING        = 1
I_BIT_IC_MISSING        = 2
I_BIT_IA_REVERSED       = 3
I_BIT_IB_REVERSED       = 4
I_BIT_IC_REVERSED       = 5
I_BIT_IA_PHASE_SHIFT    = 6
I_BIT_IB_PHASE_SHIFT    = 7
I_BIT_IC_PHASE_SHIFT    = 8

# ── 各接线方式活跃 User Channel 数 ───────────────────────────────────────────
ACTIVE_USER_CHANNELS = {
    SERVICE_3E4WY:       8,   # User 1-8，每个 A/B/C
    SERVICE_2E3W_NET:   12,   # User 1-12，每个含 2 相
    SERVICE_2E3W_DELTA: 12,   # User 1-12，每个 A/C
    SERVICE_2E3W_1PHASE:12,   # User 1-12，每个 A/C
    SERVICE_1E2W:       24,   # Input Channel 当 User Channel，每个只有 Phase A
}

# 各接线方式电流列显示的相（1E2W 只有 A 相）
CHANNEL_PHASES = {
    SERVICE_3E4WY:       ('A', 'B', 'C'),
    SERVICE_2E3W_NET:    ('A', 'B', 'C'),
    SERVICE_2E3W_DELTA:  ('A', 'C'),
    SERVICE_2E3W_1PHASE: ('A', 'C'),
    SERVICE_1E2W:        ('A',),
}

# 2E3W Network 每个 User Channel 实际分配的相（AB/CA/BC 循环 × 4）
NETWORK_CHANNEL_PHASE_MAP = (('A', 'B'), ('C', 'A'), ('B', 'C')) * 4

# ── 测试环境参数（统一维护于 projects/AcuHMI_1_7/settings.py）────────────────
from projects.AcuHMI_1_7.settings import (  # noqa: E402
    HMI_IP,
    DEFAULT_USERNAME as HMI_USER,
    DEFAULT_PASSWORD as HMI_PASS,
    HMI_DEVICE_NAME,
    WIRING_DEVICE_NAME,
    NOMINAL_VOLTAGE,
    NORMAL_CURRENT,
    CHECK_TIMEOUT,
)

from typing import Optional  # noqa: E402

# ── 被测表直连 Modbus 参数（运行时由 meter_connection fixture 从网关动态填入）──
# import 时为 None；conftest 的 meter_connection（session autouse）会在用例运行前调
# apply_discovered_connection() 填好。meter_modbus / report 仍按 cfg.METER_TCP_IP 读取。
METER_TCP_IP: Optional[str] = None
METER_TCP_PORT: Optional[int] = None
MODBUS_SLAVE: Optional[int] = None


def apply_discovered_connection(ip: str, port: int, unit: int) -> None:
    """由 meter_connection fixture 调用，填入从网关发现的被测表连接参数。"""
    global METER_TCP_IP, METER_TCP_PORT, MODBUS_SLAVE
    METER_TCP_IP, METER_TCP_PORT, MODBUS_SLAVE = ip, port, unit


def ensure_meter_connection(headless: bool = True) -> None:
    """独立运行（非 pytest）入口：若连接参数尚未填入，则起一个临时浏览器登录网关、
    动态发现被测表连接参数并填入 cfg。pytest 路径由 meter_connection fixture 负责，不会走到这里。
    """
    if METER_TCP_IP is not None:
        return
    from playwright.sync_api import sync_playwright
    from projects.AcuHMI_1_7.helpers.physical_devices_reader import (
        discover_modbus_tcp_devices,
    )
    from projects.AcuHMI_1_7.tests.Wiring_check.core.wiring_check_page import WiringCheckPage

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        try:
            page = browser.new_page(ignore_https_errors=True)
            wcp = WiringCheckPage(page, device_name=WIRING_DEVICE_NAME)
            wcp.login_if_needed()
            devices = discover_modbus_tcp_devices(page)
        finally:
            browser.close()
    target = next((d for d in devices if d.name == WIRING_DEVICE_NAME), None)
    if target is None or not target.online:
        names = [d.name for d in devices if d.online]
        raise RuntimeError(
            f"meter_device_name={WIRING_DEVICE_NAME!r} 未在网关在线 Modbus TCP 设备中找到。"
            f"当前在线设备：{names}。"
        )
    apply_discovered_connection(target.ip, target.port, target.unit)
