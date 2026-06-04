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

# ── 测试环境参数（统一维护于 test_case/AcuHMI_1_7/config.py）────────────────
from test_case.AcuHMI_1_7.config import (  # noqa: E402
    METER_TCP_IP,
    METER_TCP_PORT,
    MODBUS_SLAVE,
    HMI_IP,
    HMI_USERNAME as HMI_USER,
    HMI_PASSWORD as HMI_PASS,
    HMI_DEVICE_NAME,
    NOMINAL_VOLTAGE,
    NORMAL_CURRENT,
    CHECK_TIMEOUT,
)
