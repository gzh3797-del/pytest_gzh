"""
AcuRev4100 SNMP OID → Modbus 地址映射

SNMP OID 结构（enterprise: .1.3.6.1.4.1.39604）：
  - RT 数据：.9.3.1.3.1.{param_idx}.{dev_data_idx}
  - RT 标识：.9.3.1.2.1.{field}.{dev_data_idx}
  - 设备列表：.9.2.1.{field}.{dev_list_idx}
  - 设备计数：.9.1.0

AcuRev4100 参数分段（以 SNMP param_idx 划分）：
  2   ~ 33    基础测量参数（float，8192+）
  34  ~ 177   输入通道 1-24 基础参数（float，8264+）
  178 ~ 237   用户通道 s001-s012 基础参数（float，8600+）
  238 ~ 261   需量：系统相 A/B/C + 系统（float，8960+）
  262 ~ 405   需量：输入通道 1-24（float，9008+）
  406 ~ 477   需量：用户通道 s001-s012（float，9296+）
  478 ~ 484   电压序分量（float，9728+）
  485 ~ 502   电压 THD（float，9742/9812/9882/9952）
  503 ~ 598   输入通道电流 THD（float，9958+）
  599 ~ 682   用户通道序分量（float，11590+）
  683 ~ 1042  电能（double 4寄存器，12288+）
  1043 ~ 1060 DI/DO/RO（非保持寄存器，返回 None）
"""

import json
import os
import re

ENTERPRISE_OID = ".1.3.6.1.4.1.39604"

# ── Excel 参数模板（可选依赖，不可用时降级到 legacy 硬编码）──────────────────
try:
    from projects.AcuHMI_1_7.helpers.excel_param_map import load_device_params as _load_excel
except ImportError:
    _load_excel = None


def _excel_desc(device_name: str, mib_names: dict, param_idx: int) -> str:
    """从 Excel blockParams 的 descrption 列获取参数中文/英文描述。"""
    if _load_excel is None:
        return ""
    mib_name = mib_names.get(param_idx)
    if not mib_name:
        return ""
    try:
        params = _load_excel(device_name)
        entry = params.get(mib_name)
        if entry is None and mib_name.endswith('_Percent'):
            entry = params.get(mib_name[:-len('_Percent')] + '_%')
        if entry:
            return entry.get("desc", "")
    except Exception:
        pass
    return ""


def _excel_resolve(device_name: str, mib_names: dict, param_idx: int) -> tuple:
    """
    通过 Excel 参数模板解析 param_idx → (addr, type_str, scale)。
    链路：param_idx → MIB 对象名 → Excel paramType → addr/type/scale。
    任一环节未命中时返回 (None, None, None)，由调用方降级到 legacy 逻辑。
    """
    if _load_excel is None:
        return None, None, None
    mib_name = mib_names.get(param_idx)
    if not mib_name:
        return None, None, None
    try:
        params = _load_excel(device_name)
        entry = params.get(mib_name)
        if entry is None and mib_name.endswith('_Percent'):
            # MIB uses _Percent suffix; Excel uses _% suffix
            entry = params.get(mib_name[:-len('_Percent')] + '_%')
        if entry:
            addr, typ, scale = entry["addr"], entry["type"], entry["scale"]
            # Guard: word/word_signed with scale > 10 is almost certainly a template error
            # (16-bit integers × 100 → values up to 6.5M, impossible for any physical unit).
            # Fall back to legacy hard-coded mapping instead.
            if typ in ('word', 'word_signed') and scale > 10.0:
                return None, None, None
            return addr, typ, scale
    except Exception:
        pass
    return None, None, None

# ─── 设备注册表（从 config/devices.yaml 加载，降级到内置默认值）─────────────────

def _build_device_registry() -> list:
    """从 config/devices.yaml 构建 DEVICE_REGISTRY，加载失败时使用内置默认值。"""
    try:
        import yaml as _yaml
        _cfg_path = os.path.normpath(
            os.path.join(os.path.dirname(__file__), "devices.yaml")
        )
        _raw = _yaml.safe_load(open(_cfg_path, encoding="utf-8").read()).get("devices", {})
        _reg = []
        for name, info in _raw.items():
            entry = {
                "name":        name,
                "modbus_host": str(info["ip"]),
                "modbus_port": int(info["port"]),
                "modbus_unit": int(info["unit"]),
            }
            if "model_type" in info:
                entry["model_type"] = str(info["model_type"])
            if "dev_data_idx" in info:
                entry["dev_data_idx"] = int(info["dev_data_idx"])
            _reg.append(entry)
        return _reg
    except Exception:
        pass
    # 内置默认值（devices.yaml 不可用时的兜底）
    return [
        {"name": "AcuRev4100a", "dev_data_idx": 1, "modbus_host": "192.168.2.29", "modbus_port": 502, "modbus_unit": 203},
        {"name": "AcuRev4100",  "modbus_host": "192.168.2.30", "modbus_port": 502, "modbus_unit": 102},
        {"name": "AcuRev4100b", "modbus_host": "192.168.3.62", "modbus_port": 502, "modbus_unit": 1},
        {"name": "AcuRev2100",  "modbus_host": "192.168.2.64", "modbus_port": 502, "modbus_unit": 101},
        {"name": "Acuvim3",     "modbus_host": "192.168.2.32", "modbus_port": 502, "modbus_unit": 1},
        {"name": "AcuvimIIW",   "modbus_host": "192.168.2.27", "modbus_port": 502, "modbus_unit": 2},
        {"name": "AcuvimIIR",   "modbus_host": "192.168.3.29", "modbus_port": 502, "modbus_unit": 2},
        {"name": "AcuRev1300",  "modbus_host": "192.168.2.8",  "modbus_port": 502, "modbus_unit": 102},
    ]


DEVICE_REGISTRY = _build_device_registry()

_MIB_MAPPING_PATH = os.path.join(os.path.dirname(__file__), "mib_mapping.json")


def get_model_type(device_name: str) -> str | None:
    """
    通过设备名称在 mib_mapping.json 中查找 model_type。
    mib_manager 在每次 session 开始时从 Template 字段自动推导并写入。
    返回 None 表示 mib_mapping.json 不存在或该设备无 Template 记录。
    """
    try:
        if not os.path.exists(_MIB_MAPPING_PATH):
            return None
        mapping = json.loads(open(_MIB_MAPPING_PATH, encoding="utf-8").read())
        return mapping.get(device_name, {}).get("model_type")
    except Exception:
        return None

# ─── AcuRev4100 基础参数 Modbus 地址（snmp_idx 2-33）────────────────────────
# {param_idx: modbus_start_addr}
# Load Nature（uint32）不出现在 SNMP 中，地址连续但有跳空
_BASIC_PARAMS = {
    2:  8192,   # System Frequency
    3:  8194,   # Phase A L-N Voltage
    4:  8196,   # Phase B L-N Voltage
    5:  8198,   # Phase C L-N Voltage
    6:  8200,   # Average L-N Voltage
    7:  8202,   # Phase AB L-L Voltage
    8:  8204,   # Phase BC L-L Voltage
    9:  8206,   # Phase CA L-L Voltage
    10: 8208,   # Average L-L Voltage
    11: 8210,   # Phase A Voltage Angle
    12: 8212,   # Phase B Voltage Angle
    13: 8214,   # Phase C Voltage Angle
    14: 8216,   # Phase A Current
    15: 8218,   # Phase A Active Power
    16: 8220,   # Phase A Reactive Power
    17: 8222,   # Phase A Apparent Power
    # 8224 = Phase A Load Nature (uint32, skipped in SNMP)
    18: 8226,   # Phase A Power Factor
    19: 8228,   # Phase B Current
    20: 8230,   # Phase B Active Power
    21: 8232,   # Phase B Reactive Power
    22: 8234,   # Phase B Apparent Power
    # 8236 = Phase B Load Nature (skipped)
    23: 8238,   # Phase B Power Factor
    24: 8240,   # Phase C Current
    25: 8242,   # Phase C Active Power
    26: 8244,   # Phase C Reactive Power
    27: 8246,   # Phase C Apparent Power
    # 8248 = Phase C Load Nature (skipped)
    28: 8250,   # Phase C Power Factor
    29: 8252,   # System Average Current
    30: 8254,   # System Active Power
    31: 8256,   # System Reactive Power
    32: 8258,   # System Apparent Power
    # 8260 = System Load Nature (skipped)
    33: 8262,   # System Power Factor
}

# 输入通道内 6 个参数的 Modbus 偏移（Load Nature 跳过 +8/+9）
_INPUT_CH_OFFSETS = [0, 2, 4, 6, 10, 12]

# 用户通道内 5 个参数的 Modbus 偏移
_USER_CH_OFFSETS  = [0, 2, 4, 6, 10]


# ─── MIB 参数名解析 ───────────────────────────────────────────────────────────

_SNMP_DIR = os.path.dirname(os.path.abspath(__file__))

# 各设备类型对应的实例名前缀（用于在 mib_mapping.json 中查找对应设备）
_DEVICE_TYPE_PREFIXES: dict[str, list[str]] = {
    "AcuRev4100": ["AcuRev4100", "PXB"],
    "AcuRev2100": ["AcuRev2100"],
    "Acuvim3":    ["AcuVIM3", "Acuvim3"],
    "AcuvimIIW":  ["AcuvimIIW", "PXE2"],
    "AcuvimIIR":  ["AcuvimIIR", "PXE1"],
    "AcuRev1300": ["AcuRev1300", "PXM350"],
}

# 历史 MIB 文件的兜底路径（无 mib_mapping.json 时回退使用）
# 所有路径均指向 mib/ 子目录（由 mib_manager 从 HMI 页面下载解压）
_MIB_DIR = os.path.join(_SNMP_DIR, "mib")
_LEGACY_MIB: dict[str, tuple[str, str]] = {
    "AcuRev4100": (
        os.path.join(_MIB_DIR, "pXB-M24-XMV-GENModbus_v1.03p05.MIB"),
        "pXB_M24_XMV_GENModbus_v1_03p05DeviceReadingEntry",
    ),
    "AcuRev2100": (
        os.path.join(_MIB_DIR, "AcuRev-2100_Modbus_v1.20.MIB"),
        "AcuRev_2100_Modbus_v1_20DeviceReadingEntry",
    ),
    "Acuvim3": (
        os.path.join(_MIB_DIR, "Acuvim3_Modbus_v1.03p19.MIB"),
        "Acuvim3_Modbus_v1_03p19DeviceReadingEntry",
    ),
    "AcuvimIIW": (
        os.path.join(_MIB_DIR, "pXE2Modbus_v6.36.MIB"),
        "pXE2Modbus_v6_36DeviceReadingEntry",
    ),
}


def _parse_mib_reading_names(mib_path: str, entry_name: str) -> dict:
    """
    从 MIB 文件解析 {param_idx: name}。
    提取所有 ::= { <entry_name> N } 对应的 OBJECT-TYPE 名称。
    """
    result = {}
    if not os.path.exists(mib_path):
        return result
    current_name = None
    pattern = re.compile(
        r"::=\s*\{\s*" + re.escape(entry_name) + r"\s+(\d+)\s*\}"
    )
    with open(mib_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            stripped = line.strip()
            m = re.match(r"^(\w+)\s+OBJECT-TYPE\s*$", stripped)
            if m:
                current_name = m.group(1)
            if current_name:
                m = pattern.match(stripped)
                if m:
                    result[int(m.group(1))] = current_name
                    current_name = None
    return result


class _LazyMibParams(dict):
    """
    懒加载的 MIB 参数名字典。首次访问时从 mib_mapping.json 加载对应 MIB 文件；
    mib_mapping.json 不存在或无匹配时，回退到历史硬编码 MIB 文件。
    """

    def __init__(self, device_type: str) -> None:
        super().__init__()
        self._device_type = device_type
        self._loaded = False

    def _load(self) -> None:
        if self._loaded:
            return
        self._loaded = True

        # 1. 尝试 mib_mapping.json
        mapping_path = os.path.join(_SNMP_DIR, "mib_mapping.json")
        if os.path.exists(mapping_path):
            try:
                import json as _json
                mapping = _json.loads(
                    open(mapping_path, encoding="utf-8").read()
                )
                prefixes = _DEVICE_TYPE_PREFIXES.get(self._device_type, [self._device_type])
                for dev_name, info in mapping.items():
                    if any(dev_name == p or dev_name.startswith(p) for p in prefixes):
                        mib_rel = info.get("mib_file")
                        entry_name = info.get("entry_name")
                        if mib_rel and entry_name:
                            mib_abs = os.path.join(_SNMP_DIR, mib_rel)
                            data = _parse_mib_reading_names(mib_abs, entry_name)
                            if data:
                                self.update(data)
                                return
            except Exception:
                pass

        # 2. 回退到历史文件
        legacy = _LEGACY_MIB.get(self._device_type)
        if legacy:
            mib_path, entry_name = legacy
            self.update(_parse_mib_reading_names(mib_path, entry_name))

    # ── dict 接口代理（触发懒加载）──────────────────────────────────────────

    def get(self, key, default=None):
        self._load()
        return super().get(key, default)

    def __contains__(self, key):
        self._load()
        return super().__contains__(key)

    def __len__(self):
        self._load()
        return super().__len__()

    def __iter__(self):
        self._load()
        return super().__iter__()

    def items(self):
        self._load()
        return super().items()

    def values(self):
        self._load()
        return super().values()

    def keys(self):
        self._load()
        return super().keys()


_MIB_PARAM_NAMES:            _LazyMibParams = _LazyMibParams("AcuRev4100")
_MIB_2100_PARAM_NAMES:       _LazyMibParams = _LazyMibParams("AcuRev2100")
_MIB_ACUVIM3_PARAM_NAMES:    _LazyMibParams = _LazyMibParams("Acuvim3")
_MIB_ACUVIMIIIW_PARAM_NAMES: _LazyMibParams = _LazyMibParams("AcuvimIIW")
_MIB_ACUVIMIIIR_PARAM_NAMES: _LazyMibParams = _LazyMibParams("AcuvimIIR")
_MIB_ACUREV1300_PARAM_NAMES: _LazyMibParams = _LazyMibParams("AcuRev1300")


def mib_loaded() -> bool:
    """返回 AcuRev4100 MIB 是否成功加载（向后兼容）。"""
    return len(_MIB_PARAM_NAMES) > 0


# ─── 公开 API ─────────────────────────────────────────────────────────────────

def acurev4100_snmp_to_modbus(param_idx: int) -> int | None:
    """
    将 AcuRev4100 SNMP param_idx 转换为对应的 Modbus 起始寄存器地址。
    返回 None 表示无对应 Modbus 地址（DI/DO/RO 使用非保持寄存器协议）。
    优先查 Excel 参数模板；未命中时降级到以下公式映射。
    """
    addr, _, _ = _excel_resolve("AcuRev4100", _MIB_PARAM_NAMES, param_idx)
    if addr is not None:
        return addr
    # ── legacy 兜底：DI/DO/RO 等 Excel 不覆盖的参数 ────────────────────────
    # ── 基础参数（float，2 regs）───────────────────────────────────────────
    if param_idx in _BASIC_PARAMS:
        return _BASIC_PARAMS[param_idx]

    # ── 输入通道 1-24 基础（float，14 regs/ch，Load Nature 在 +8/+9 跳空）─
    if 34 <= param_idx <= 177:
        ch  = (param_idx - 34) // 6
        pos = (param_idx - 34) % 6
        return 8264 + ch * 14 + _INPUT_CH_OFFSETS[pos]

    # ── 用户通道 s001-s012 基础（float，12 regs/ch）───────────────────────
    if 178 <= param_idx <= 237:
        ch  = (param_idx - 178) // 5
        pos = (param_idx - 178) % 5
        return 8600 + ch * 12 + _USER_CH_OFFSETS[pos]

    # ── 需量：系统相 A/B/C + 系统（float，12 regs/group，6 params: I/IMP_P/EXP_P/IMP_Q/EXP_Q/S）
    if 238 <= param_idx <= 261:
        group = (param_idx - 238) // 6   # 0=A, 1=B, 2=C, 3=Sys
        pos   = (param_idx - 238) % 6
        return 8960 + group * 12 + pos * 2

    # ── 需量：输入通道 1-24（float，12 regs/ch）──────────────────────────
    if 262 <= param_idx <= 405:
        ch  = (param_idx - 262) // 6
        pos = (param_idx - 262) % 6
        return 9008 + ch * 12 + pos * 2

    # ── 需量：用户通道 s001-s012（float，12 regs/ch）─────────────────────
    if 406 <= param_idx <= 477:
        ch  = (param_idx - 406) // 6
        pos = (param_idx - 406) % 6
        return 9296 + ch * 12 + pos * 2

    # ── 电压序分量（float，7 params）────────────────────────────────────
    if 478 <= param_idx <= 484:
        return 9728 + (param_idx - 478) * 2

    # ── 电压 THD：相 A/B/C 各 5 params（含 30 个谐波跳空 60 regs），avg 3 params
    if 485 <= param_idx <= 499:
        group = (param_idx - 485) // 5   # 0=A, 1=B, 2=C
        pos   = (param_idx - 485) % 5
        return 9742 + group * 70 + pos * 2
    if 500 <= param_idx <= 502:
        return 9952 + (param_idx - 500) * 2

    # ── 输入通道电流 THD（float，68 regs/ch：4 SNMP + 30 谐波跳空 60 regs）
    if 503 <= param_idx <= 598:
        ch  = (param_idx - 503) // 4
        pos = (param_idx - 503) % 4
        return 9958 + ch * 68 + pos * 2

    # ── 用户通道序分量（float，14 regs/ch，7 params: MAG_POS/ZERO/NEG + ANG_POS/ZERO/NEG + UNBL）
    if 599 <= param_idx <= 682:
        ch  = (param_idx - 599) // 7
        pos = (param_idx - 599) % 7
        return 11590 + ch * 14 + pos * 2

    # ── 电能：系统相 A/B/C + 系统（double，4 regs，9 params/group: EP×4 EQ×4 ES）
    if 683 <= param_idx <= 718:
        group = (param_idx - 683) // 9   # 0=A, 1=B, 2=C, 3=Sys
        pos   = (param_idx - 683) % 9
        return 12288 + group * 36 + pos * 4

    # ── 电能：输入通道 1-24（double，36 regs/ch）─────────────────────────
    if 719 <= param_idx <= 934:
        ch  = (param_idx - 719) // 9
        pos = (param_idx - 719) % 9
        return 12432 + ch * 36 + pos * 4

    # ── 电能：用户通道 s001-s012（double，36 regs/ch）────────────────────
    if 935 <= param_idx <= 1042:
        ch  = (param_idx - 935) // 9
        pos = (param_idx - 935) % 9
        return 13296 + ch * 36 + pos * 4

    # ── DI 脉冲计数（保持寄存器 03H，uint32，2 regs）────────────────────
    # idx 1043-1046: DI_PC_001~004
    if 1043 <= param_idx <= 1046:
        return 25600 + (param_idx - 1043) * 2

    # ── DI 状态（离散输入 02H，bit，addr 0-3）────────────────────────────
    # idx 1047-1050: DI_ST_001~004
    if 1047 <= param_idx <= 1050:
        return param_idx - 1047

    # ── DO 状态（线圈 01H，bit，addr 0-7）───────────────────────────────
    # idx 1051-1058: DO_ST_001~008
    if 1051 <= param_idx <= 1058:
        return param_idx - 1051

    # ── RO 状态（线圈 01H，bit，addr 32-33）─────────────────────────────
    # idx 1059-1060: RO_ST_001~002
    if 1059 <= param_idx <= 1060:
        return 32 + (param_idx - 1059)

    return None


def acurev4100_param_type(param_idx: int) -> str | None:
    """
    返回 AcuRev4100 SNMP 参数的 Modbus 读取类型。
    优先查 Excel 参数模板；未命中时降级到以下规则。
      'float'     → 2 寄存器，03H，ABCD 大端序 float
      'double'    → 4 寄存器，03H，ABCDEFGH 大端序 double
      'uint32'    → 2 寄存器，03H，ABCD 大端序 uint32（DI 脉冲计数）
      'bit_di'    → 1 bit，02H 离散输入（DI 状态）
      'bit_coil'  → 1 bit，01H 线圈（DO/RO 状态）
      None        → 无对应地址
    """
    # DI_ST/DO_ST/RO_ST: Excel 把这些标为 'word'，但正确 FC 是 02H/01H，
    # 必须在 Excel 查询前强制覆盖，否则走 FC=03 保持寄存器必然失败。
    if 1047 <= param_idx <= 1050:
        return 'bit_di'
    if 1051 <= param_idx <= 1060:
        return 'bit_coil'
    _, typ, _ = _excel_resolve("AcuRev4100", _MIB_PARAM_NAMES, param_idx)
    if typ is not None:
        return typ
    # legacy
    if acurev4100_snmp_to_modbus(param_idx) is None:
        return None
    if 683 <= param_idx <= 1042:
        return 'double'
    if 1043 <= param_idx <= 1046:
        return 'uint32'
    if 1047 <= param_idx <= 1050:
        return 'bit_di'
    if 1051 <= param_idx <= 1060:
        return 'bit_coil'
    return 'float'


def acurev4100_modbus_scale(param_idx: int) -> float:
    """AcuRev4100 换算系数（大多数参数 float32 直接物理量，scale=1.0）。"""
    _, _, scale = _excel_resolve("AcuRev4100", _MIB_PARAM_NAMES, param_idx)
    return scale if scale is not None else 1.0


def acurev4100_param_name(param_idx: int) -> str:
    """
    返回 AcuRev4100 SNMP 参数名称。
    优先使用 MIB 文件中的官方名称（如 FREQ_Hz、VLN_a_V），
    MIB 未加载时回退到公式推导的描述字符串。
    """
    if param_idx in _MIB_PARAM_NAMES:
        return _MIB_PARAM_NAMES[param_idx]

    # MIB 未找到时的回退
    _fallback_basic = {
        2: "System Frequency",       3: "Phase A L-N Voltage",
        4: "Phase B L-N Voltage",    5: "Phase C L-N Voltage",
        6: "Average L-N Voltage",    7: "Phase AB L-L Voltage",
        8: "Phase BC L-L Voltage",   9: "Phase CA L-L Voltage",
        10: "Average L-L Voltage",   11: "Phase A Voltage Angle",
        12: "Phase B Voltage Angle", 13: "Phase C Voltage Angle",
        14: "Phase A Current",       15: "Phase A Active Power",
        16: "Phase A Reactive Power",17: "Phase A Apparent Power",
        18: "Phase A Power Factor",  19: "Phase B Current",
        20: "Phase B Active Power",  21: "Phase B Reactive Power",
        22: "Phase B Apparent Power",23: "Phase B Power Factor",
        24: "Phase C Current",       25: "Phase C Active Power",
        26: "Phase C Reactive Power",27: "Phase C Apparent Power",
        28: "Phase C Power Factor",  29: "System Average Current",
        30: "System Active Power",   31: "System Reactive Power",
        32: "System Apparent Power", 33: "System Power Factor",
    }
    if param_idx in _fallback_basic:
        return _fallback_basic[param_idx]
    if 34 <= param_idx <= 177:
        ch  = (param_idx - 34) // 6 + 1
        pos = (param_idx - 34) % 6
        names = ["Current", "Active Power", "Reactive Power",
                 "Apparent Power", "Power Factor", "Current Phase Angle"]
        return f"Input Ch{ch} {names[pos]}"
    if param_idx >= 178:
        ch  = (param_idx - 178) // 5 + 1
        pos = (param_idx - 178) % 5
        names = ["Current", "Active Power", "Reactive Power",
                 "Apparent Power", "Power Factor"]
        return f"User Ch{ch} {names[pos]}"
    return f"Param#{param_idx}"


# ─── OID 构造工具 ─────────────────────────────────────────────────────────────

def build_rt_oid(param_idx: int, dev_data_idx: int = 1, subtree: int = 1) -> str:
    return f"{ENTERPRISE_OID}.9.3.{subtree}.3.1.{param_idx}.{dev_data_idx}"


def build_device_list_oid(field_idx: int, dev_idx: int) -> str:
    return f"{ENTERPRISE_OID}.9.2.1.{field_idx}.{dev_idx}"


def build_rt_info_oid(field_idx: int, dev_data_idx: int = 1, subtree: int = 1) -> str:
    return f"{ENTERPRISE_OID}.9.3.{subtree}.2.1.{field_idx}.{dev_data_idx}"


def discover_dev_data_idx(device_name: str, snmp_data: dict):
    """从 SNMP walk 结果中动态查找设备所在的 (subtree, dev_data_idx)。
    扫描所有 .9.3.{subtree}.2.1.{field}.{idx} 模式，大小写不敏感。
    返回 (subtree_int, dev_data_idx_int) 或 (None, None)。
    两轮搜索：先精确匹配，再允许后跟非字母数字字符的子串匹配，
    防止 "AcuRev4100" 误匹配 "AcuRev4100b"。
    """
    name_lower = device_name.lower()
    pat = re.compile(
        re.escape(ENTERPRISE_OID) + r"\.9\.3\.(\d+)\.2\.1\.(\d+)\.(\d+)"
    )

    def _name_matches(val: str) -> bool:
        v = val.lower()
        if v == name_lower:
            return True
        pos = v.find(name_lower)
        if pos < 0:
            return False
        after = pos + len(name_lower)
        return after >= len(v) or not v[after].isalnum()

    for oid, val in snmp_data.items():
        m = pat.match(oid)
        if m and val and _name_matches(val):
            return int(m.group(1)), int(m.group(3))
    return None, None


def dump_device_oids(snmp_data: dict) -> str:
    """返回 walk 数据中所有设备标识相关 OID 的诊断字符串（用于排查 dev_data_idx 未找到）。"""
    pat = re.compile(
        re.escape(ENTERPRISE_OID) + r"\.9\.(1|2|3\.\d+\.(2|3\.1\.1))\."
    )
    lines = [f"  {'OID':<65} 值"]
    lines.append("  " + "-" * 90)
    for oid in sorted(snmp_data):
        if pat.match(oid):
            lines.append(f"  {oid:<65} {snmp_data[oid]}")
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
# AcuRev2100 SNMP → Modbus 映射
# 参数 idx 基于地址表 AcuRev2100_ Modbus Address_v1.02_20260406.xlsx 推算；
# 实际 param_idx 需通过 SNMP walk 验证后按需修正。
# ═══════════════════════════════════════════════════════════════════════════════

# 基础 inline 参数（Float，2寄存器，0x2000/8192起）
_ACUREV2100_BASIC = {
    2:  8192,   # Freq
    3:  8194,   # U1
    4:  8196,   # U2
    5:  8198,   # U3
    6:  8200,   # Uavg
    7:  8202,   # U12
    8:  8204,   # U23
    9:  8206,   # U31
    10: 8208,   # Ulavg
    11: 8210,   # IL1
    12: 8212,   # IL2
    13: 8214,   # IL3
    14: 8216,   # Iavg
    15: 8218,   # Pin-S
    16: 8220,   # Qin-S
    17: 8222,   # Sin-S
    18: 8224,   # PFin-S  [8226=LN-S skip]
    19: 8228,   # Pin-A
    20: 8230,   # Pin-B
    21: 8232,   # Pin-C
    22: 8234,   # Qin-A
    23: 8236,   # Qin-B
    24: 8238,   # Qin-C
    25: 8240,   # Sin-A
    26: 8242,   # Sin-B
    27: 8244,   # Sin-C
    28: 8246,   # PFin-A
    29: 8248,   # PFin-B
    30: 8250,   # PFin-C   [8252/54/56=LN-A/B/C skip]
}

# 每通道偏移：stride=12，取 I/P/Q/S/PF（跳过 LN at +10）
_ACUREV2100_CH_OFFSETS  = [0, 2, 4, 6, 8]
# 用户通道偏移：stride=10，取 Ps/Qs/Ss/PFs（跳过 LN at +8）
_ACUREV2100_UCH_OFFSETS = [0, 2, 4, 6]


def acurev2100_snmp_to_modbus(param_idx: int) -> int | None:
    """AcuRev2100 SNMP param_idx → Modbus 保持寄存器起始地址。
    优先查 Excel 参数模板；未命中时降级到以下公式映射。"""
    addr, _, _ = _excel_resolve("AcuRev2100", _MIB_2100_PARAM_NAMES, param_idx)
    if addr is not None:
        return addr
    # legacy
    if param_idx in _ACUREV2100_BASIC:
        return _ACUREV2100_BASIC[param_idx]

    # ── 独立通道 1-18（idx 31-120，5 params/ch，stride 12）────────────────
    if 31 <= param_idx <= 120:
        ch  = (param_idx - 31) // 5
        pos = (param_idx - 31) % 5
        return 8448 + ch * 12 + _ACUREV2100_CH_OFFSETS[pos]

    # ── 用户通道 1-9（idx 121-156，4 params/ch，stride 10）───────────────
    if 121 <= param_idx <= 156:
        ch  = (param_idx - 121) // 4
        pos = (param_idx - 121) % 4
        return 8664 + ch * 10 + _ACUREV2100_UCH_OFFSETS[pos]

    # EP_IMP A/B/C/S idx 157-160; ch 161-178; user s001-s009 idx 179-187
    if 157 <= param_idx <= 160:
        return [9472, 9474, 9476, 9478][param_idx - 157]   # A, B, C, S
    if 161 <= param_idx <= 178:
        return 9480 + (param_idx - 161) * 2
    if 179 <= param_idx <= 187:
        return 9516 + (param_idx - 179) * 2

    # EQ_IMP A/B/C/S idx 188-191; ch 192-209; user s001-s006 idx 210-215; s007-s009 idx 244-246
    if 188 <= param_idx <= 191:
        return [11008, 11010, 11012, 11014][param_idx - 188]  # A, B, C, S
    if 192 <= param_idx <= 209:
        return 11016 + (param_idx - 192) * 2
    if 210 <= param_idx <= 215:
        return 11052 + (param_idx - 210) * 2   # EQ s001-s006
    if 244 <= param_idx <= 246:
        return 11120 + (param_idx - 244) * 2   # EQ s007-s009

    # ES A/B/C/S idx 216-219; ch 220-237; user s001-s006 idx 238-243; s007-s009 idx 247-249
    if 216 <= param_idx <= 219:
        return [11064, 11066, 11068, 11070][param_idx - 216]  # A, B, C, S
    if 220 <= param_idx <= 237:
        return 11072 + (param_idx - 220) * 2
    if 238 <= param_idx <= 243:
        return 11108 + (param_idx - 238) * 2   # ES s001-s006
    if 247 <= param_idx <= 249:
        return 11126 + (param_idx - 247) * 2   # ES s007-s009

    # Demand S (idx 250-255): P/Q/S demand+pred at base 11520, offsets [0,2,9,11,18,20]
    if 250 <= param_idx <= 255:
        return 11520 + [0, 2, 9, 11, 18, 20][param_idx - 250]

    # Demand Phase A/B/C (idx 256-279): 3 phases × 8 params, offsets [0,2,9,11,18,20,27,29]
    if 256 <= param_idx <= 279:
        phase = (param_idx - 256) // 8   # 0=A, 1=B, 2=C
        pos   = (param_idx - 256) % 8
        bases = [11547, 11583, 11619]
        offs  = [0, 2, 9, 11, 18, 20, 27, 29]
        return bases[phase] + offs[pos]

    # Demand ch 1-18 (idx 280-423): 18 ch × 8 params, base 11655, stride 36
    if 280 <= param_idx <= 423:
        ch  = (param_idx - 280) // 8
        pos = (param_idx - 280) % 8
        return 11655 + ch * 36 + [0, 2, 9, 11, 18, 20, 27, 29][pos]

    # Demand user ch 1-9 (idx 424-477): 9 ch × 6 params, base 12303, stride 27
    if 424 <= param_idx <= 477:
        ch  = (param_idx - 424) // 6
        pos = (param_idx - 424) % 6
        return 12303 + ch * 27 + [0, 2, 9, 11, 18, 20][pos]

    # PQ (Word, 1 register each)
    if param_idx == 478:
        return 12800                          # UNBL_V
    if 479 <= param_idx <= 482:
        return 12801 + (param_idx - 479)      # THD_V a/b/c/avg
    if param_idx == 483:
        return 12895                          # UNBL_I
    if 484 <= param_idx <= 495:
        return 12896 + (param_idx - 484)      # OTHD/ETHD/CF/THFF V per phase
    if 496 <= param_idx <= 549:
        return 12960 + (param_idx - 496)      # OTHD/ETHD/KF I ch 1-18 (sequential)
    if 550 <= param_idx <= 567:
        return 13056 + (param_idx - 550) * 31  # THD_I ch 1-18 (stride 31)
    if 568 <= param_idx <= 576:
        return 13614 + (param_idx - 568)      # UNBL_I_s001-s009
    if 577 <= param_idx <= 578:
        return 13824 + (param_idx - 577)      # ANG_VLN_b, ANG_VLN_c
    if 579 <= param_idx <= 596:
        return 13826 + (param_idx - 579)      # ANG_I_001-018

    return None


def acurev2100_param_type(param_idx: int) -> str | None:
    """AcuRev2100 参数的 Modbus 读取类型。
    优先查 Excel 参数模板；未命中时降级到以下规则。"""
    _, typ, _ = _excel_resolve("AcuRev2100", _MIB_2100_PARAM_NAMES, param_idx)
    if typ is not None:
        return typ
    # legacy
    if acurev2100_snmp_to_modbus(param_idx) is None:
        return None
    if (157 <= param_idx <= 215) or (216 <= param_idx <= 249):
        return 'uint32'
    if 250 <= param_idx <= 477:
        return 'float'
    if 478 <= param_idx <= 596:
        return 'word'
    return 'float'


def acurev2100_modbus_scale(param_idx: int) -> float:
    """AcuRev2100 Modbus 原始值换算系数。优先查 Excel，未命中降级到以下规则。"""
    _, _, scale = _excel_resolve("AcuRev2100", _MIB_2100_PARAM_NAMES, param_idx)
    if scale is not None:
        return scale
    # legacy
    if 157 <= param_idx <= 249:
        return 0.1    # 电能 Dword × 0.1 = kWh/kvarh/kVAh
    # PQ scaling（Word 类型，按参数种类不同系数）
    if 478 <= param_idx <= 596:
        # UNBL_V(478): ×0.1
        if param_idx == 478:
            return 0.1
        # THD_V a/b/c/avg(479-482): ×0.01
        if 479 <= param_idx <= 482:
            return 0.01
        # UNBL_I(483): ×0.1
        if param_idx == 483:
            return 0.1
        # OTHD/ETHD/CF/THFF voltage(484-495): CF=×0.001, THFF=×0.01, others=×0.01
        if 484 <= param_idx <= 495:
            pos = (param_idx - 484) % 4   # 0=OTHD,1=ETHD,2=CF,3=THFF
            return 0.001 if pos == 2 else 0.01
        # OTHD/ETHD current ch(496-549): ×0.01; KF: ×0.1
        if 496 <= param_idx <= 549:
            pos = (param_idx - 496) % 3   # 0=OTHD,1=ETHD,2=KF
            return 0.1 if pos == 2 else 0.01
        # THD_I(550-567): ×0.01
        if 550 <= param_idx <= 567:
            return 0.01
        # UNBL_I_s(568-576): ×0.1
        if 568 <= param_idx <= 576:
            return 0.1
        # Angles(577-596): ×0.1
        if 577 <= param_idx <= 596:
            return 0.1
    return 1.0


def acurev4100_param_desc(param_idx: int) -> str:
    """AcuRev4100 参数字段描述（来自 Excel descrption 列）。"""
    return _excel_desc("AcuRev4100", _MIB_PARAM_NAMES, param_idx)


def acurev2100_param_name(param_idx: int) -> str:
    """AcuRev2100 参数描述名称。优先返回 MIB 官方名称。"""
    if param_idx in _MIB_2100_PARAM_NAMES:
        return _MIB_2100_PARAM_NAMES[param_idx]
    _BASIC_NAMES = {
        2: "Freq",    3: "U1",     4: "U2",     5: "U3",    6: "Uavg",
        7: "U12",     8: "U23",    9: "U31",    10: "Ulavg",
        11: "IL1",   12: "IL2",   13: "IL3",   14: "Iavg",
        15: "Pin-S", 16: "Qin-S", 17: "Sin-S", 18: "PFin-S",
        19: "Pin-A", 20: "Pin-B", 21: "Pin-C",
        22: "Qin-A", 23: "Qin-B", 24: "Qin-C",
        25: "Sin-A", 26: "Sin-B", 27: "Sin-C",
        28: "PFin-A",29: "PFin-B",30: "PFin-C",
    }
    if param_idx in _BASIC_NAMES:
        return _BASIC_NAMES[param_idx]
    if 31 <= param_idx <= 120:
        ch = (param_idx - 31) // 5 + 1
        return f"Ch{ch:02d}_{['I','P','Q','S','PF'][(param_idx-31)%5]}"
    if 121 <= param_idx <= 156:
        ch = (param_idx - 121) // 4 + 1
        return f"UCh{ch}_{['Ps','Qs','Ss','PFs'][(param_idx-121)%4]}"
    if 157 <= param_idx <= 160:
        return ["Epin-S","Epin-A","Epin-B","Epin-C"][param_idx - 157]
    if 161 <= param_idx <= 178:
        return f"Epin-Ch{param_idx - 160:02d}"
    if 179 <= param_idx <= 187:
        return f"Epin-UCh{param_idx - 178}"
    if 188 <= param_idx <= 191:
        return ["Eqin-S","Eqin-A","Eqin-B","Eqin-C"][param_idx - 188]
    if 192 <= param_idx <= 209:
        return f"Eqin-Ch{param_idx - 191:02d}"
    if 210 <= param_idx <= 218:
        return f"Eqin-UCh{param_idx - 209}"
    if 219 <= param_idx <= 222:
        return ["Esin-S","Esin-A","Esin-B","Esin-C"][param_idx - 219]
    if 223 <= param_idx <= 240:
        return f"Esin-Ch{param_idx - 222:02d}"
    if 241 <= param_idx <= 249:
        return f"Esin-UCh{param_idx - 240}"
    if 250 <= param_idx <= 267:
        return f"DI{param_idx - 249:02d}_PC"
    if 268 <= param_idx <= 285:
        return f"DI{param_idx - 267:02d}_ST"
    if 286 <= param_idx <= 287:
        return f"RO{param_idx - 285}_ST"
    return f"Param#{param_idx}"


# ─── Acuvim3 完整映射（param_idx → (modbus_addr, type, scale)）────────────────
# float32: idx 2-119；double(4寄存器): idx 120-155；scale 全部=1.0（已是物理量）
_ACUVIM3_MAP: dict[int, tuple] = {
      2: (8468,  "float",  1.0),   # FREQ_Hz
      3: (8470,  "float",  1.0),   # VLN_a_V
      4: (8472,  "float",  1.0),   # VLN_b_V
      5: (8474,  "float",  1.0),   # VLN_c_V
      6: (8476,  "float",  1.0),   # VLN_avg_V
      7: (8478,  "float",  1.0),   # VLL_ab_V
      8: (8480,  "float",  1.0),   # VLL_bc_V
      9: (8482,  "float",  1.0),   # VLL_ca_V
     10: (8484,  "float",  1.0),   # VLL_avg_V
     11: (8486,  "float",  1.0),   # I_a_A
     12: (8488,  "float",  1.0),   # I_b_A
     13: (8490,  "float",  1.0),   # I_c_A
     14: (8492,  "float",  1.0),   # I_n_A
     15: (8494,  "float",  1.0),   # I_avg_A
     16: (8496,  "float",  1.0),   # P_a_kW
     17: (8498,  "float",  1.0),   # P_b_kW
     18: (8500,  "float",  1.0),   # P_c_kW
     19: (8502,  "float",  1.0),   # P_kW
     20: (8504,  "float",  1.0),   # Q_a_kvar
     21: (8506,  "float",  1.0),   # Q_b_kvar
     22: (8508,  "float",  1.0),   # Q_c_kvar
     23: (8510,  "float",  1.0),   # Q_kvar
     24: (8512,  "float",  1.0),   # S_a_kVA
     25: (8514,  "float",  1.0),   # S_b_kVA
     26: (8516,  "float",  1.0),   # S_c_kVA
     27: (8518,  "float",  1.0),   # S_kVA
     28: (8528,  "float",  1.0),   # PF_a
     29: (8530,  "float",  1.0),   # PF_b
     30: (8532,  "float",  1.0),   # PF_c
     31: (8534,  "float",  1.0),   # PF
     32: (8536,  "float",  1.0),   # LEAD_PF_a
     33: (8538,  "float",  1.0),   # LEAD_PF_b
     34: (8540,  "float",  1.0),   # LEAD_PF_c
     35: (8542,  "float",  1.0),   # LEAD_PF
     36: (8544,  "float",  1.0),   # LAG_PF_a
     37: (8546,  "float",  1.0),   # LAG_PF_b
     38: (8548,  "float",  1.0),   # LAG_PF_c
     39: (8550,  "float",  1.0),   # LAG_PF
     40: (8552,  "float",  1.0),   # ANG_VLN_a
     41: (8554,  "float",  1.0),   # ANG_VLN_b
     42: (8556,  "float",  1.0),   # ANG_VLN_c
     43: (8558,  "float",  1.0),   # ANG_VLL_ab
     44: (8560,  "float",  1.0),   # ANG_VLL_bc
     45: (8562,  "float",  1.0),   # ANG_VLL_ca
     46: (8564,  "float",  1.0),   # ANG_I_a
     47: (8566,  "float",  1.0),   # ANG_I_b
     48: (8568,  "float",  1.0),   # ANG_I_c
     49: (8570,  "float",  1.0),   # MAG_SEQ_POS_V
     50: (8572,  "float",  1.0),   # MAG_SEQ_ZERO_V
     51: (8574,  "float",  1.0),   # MAG_SEQ_NEG_V
     52: (8576,  "float",  1.0),   # SEQ_ZERO_RATIO_V_%
     53: (8578,  "float",  1.0),   # UNBL_V_%
     54: (8580,  "float",  1.0),   # MAG_SEQ_POS_I
     55: (8582,  "float",  1.0),   # MAG_SEQ_ZERO_I
     56: (8584,  "float",  1.0),   # MAG_SEQ_NEG_I
     57: (8586,  "float",  1.0),   # SEQ_ZERO_RATIO_I_%
     58: (8588,  "float",  1.0),   # UNBL_I_%
     59: (8590,  "float",  1.0),   # ANG_SEQ_POS_V
     60: (8592,  "float",  1.0),   # ANG_SEQ_ZERO_V
     61: (8594,  "float",  1.0),   # ANG_SEQ_NEG_V
     62: (8596,  "float",  1.0),   # ANG_SEQ_POS_I
     63: (8598,  "float",  1.0),   # ANG_SEQ_ZERO_I
     64: (8600,  "float",  1.0),   # ANG_SEQ_NEG_I
     65: (8602,  "float",  1.0),   # THD_V_a_%
     66: (8604,  "float",  1.0),   # THD_V_b_%
     67: (8606,  "float",  1.0),   # THD_V_c_%
     68: (8608,  "float",  1.0),   # OTHD_V_a_%
     69: (8610,  "float",  1.0),   # OTHD_V_b_%
     70: (8612,  "float",  1.0),   # OTHD_V_c_%
     71: (8614,  "float",  1.0),   # ETHD_V_a_%
     72: (8616,  "float",  1.0),   # ETHD_V_b_%
     73: (8618,  "float",  1.0),   # ETHD_V_c_%
     74: (8620,  "float",  1.0),   # CF_V_a_%
     75: (8622,  "float",  1.0),   # CF_V_b_%
     76: (8624,  "float",  1.0),   # CF_V_c_%
     77: (8626,  "float",  1.0),   # THD_I_a_%
     78: (8628,  "float",  1.0),   # THD_I_b_%
     79: (8630,  "float",  1.0),   # THD_I_c_%
     80: (8632,  "float",  1.0),   # THD_I_n_%
     81: (8634,  "float",  1.0),   # OTHD_I_a_%
     82: (8636,  "float",  1.0),   # OTHD_I_b_%
     83: (8638,  "float",  1.0),   # OTHD_I_c_%
     84: (8640,  "float",  1.0),   # OTHD_I_n_%
     85: (8642,  "float",  1.0),   # ETHD_I_a_%
     86: (8644,  "float",  1.0),   # ETHD_I_b_%
     87: (8646,  "float",  1.0),   # ETHD_I_c_%
     88: (8648,  "float",  1.0),   # ETHD_I_n_%
     89: (8650,  "float",  1.0),   # TDD_I_a_%
     90: (8652,  "float",  1.0),   # TDD_I_b_%
     91: (8654,  "float",  1.0),   # TDD_I_c_%
     92: (8656,  "float",  1.0),   # TDD_I_n_%
     93: (8658,  "float",  1.0),   # CF_I_a_%
     94: (8660,  "float",  1.0),   # CF_I_b_%
     95: (8662,  "float",  1.0),   # CF_I_c_%
     96: (8664,  "float",  1.0),   # CF_I_n_%
     97: (8666,  "float",  1.0),   # KF_I_a_%
     98: (8668,  "float",  1.0),   # KF_I_b_%
     99: (8670,  "float",  1.0),   # KF_I_c_%
    100: (8672,  "float",  1.0),   # KF_I_n_%
    101: (8674,  "float",  1.0),   # FLICK_V_a
    102: (8676,  "float",  1.0),   # FLICK_V_b
    103: (8678,  "float",  1.0),   # FLICK_V_c
    104: (24832, "float",  1.0),   # DMD_I_a_A
    105: (24834, "float",  1.0),   # DMD_I_b_A
    106: (24836, "float",  1.0),   # DMD_I_c_A
    107: (24838, "float",  1.0),   # DMD_I_avg_A
    108: (24984, "float",  1.0),   # DMD_P_a_kW
    109: (24986, "float",  1.0),   # DMD_P_b_kW
    110: (24988, "float",  1.0),   # DMD_P_c_kW
    111: (24990, "float",  1.0),   # DMD_P_kW
    112: (24992, "float",  1.0),   # DMD_Q_a_kvar
    113: (24994, "float",  1.0),   # DMD_Q_b_kvar
    114: (24996, "float",  1.0),   # DMD_Q_c_kvar
    115: (24998, "float",  1.0),   # DMD_Q_kvar
    116: (25000, "float",  1.0),   # DMD_S_a_kVA
    117: (25002, "float",  1.0),   # DMD_S_b_kVA
    118: (25004, "float",  1.0),   # DMD_S_c_kVA
    119: (25006, "float",  1.0),   # DMD_S_kVA
    120: (26048, "double", 1.0),   # EP_IMP_a_kWh
    121: (26052, "double", 1.0),   # EP_IMP_b_kWh
    122: (26056, "double", 1.0),   # EP_IMP_c_kWh
    123: (26060, "double", 1.0),   # EP_IMP_kWh
    124: (26064, "double", 1.0),   # EQ_IMP_a_kvarh
    125: (26068, "double", 1.0),   # EQ_IMP_b_kvarh
    126: (26072, "double", 1.0),   # EQ_IMP_c_kvarh
    127: (26076, "double", 1.0),   # EQ_IMP_kvarh
    128: (26080, "double", 1.0),   # EP_EXP_a_kWh
    129: (26084, "double", 1.0),   # EP_EXP_b_kWh
    130: (26088, "double", 1.0),   # EP_EXP_c_kWh
    131: (26092, "double", 1.0),   # EP_EXP_kWh
    132: (26096, "double", 1.0),   # EQ_EXP_a_kvarh
    133: (26100, "double", 1.0),   # EQ_EXP_b_kvarh
    134: (26104, "double", 1.0),   # EQ_EXP_c_kvarh
    135: (26108, "double", 1.0),   # EQ_EXP_kvarh
    136: (26112, "double", 1.0),   # EP_NET_a_kWh
    137: (26116, "double", 1.0),   # EP_NET_b_kWh
    138: (26120, "double", 1.0),   # EP_NET_c_kWh
    139: (26124, "double", 1.0),   # EP_NET_kWh
    140: (26128, "double", 1.0),   # EQ_NET_a_kvarh
    141: (26132, "double", 1.0),   # EQ_NET_b_kvarh
    142: (26136, "double", 1.0),   # EQ_NET_c_kvarh
    143: (26140, "double", 1.0),   # EQ_NET_kvarh
    144: (26144, "double", 1.0),   # EP_TOTAL_a_kWh
    145: (26148, "double", 1.0),   # EP_TOTAL_b_kWh
    146: (26152, "double", 1.0),   # EP_TOTAL_c_kWh
    147: (26156, "double", 1.0),   # EP_TOTAL_kWh
    148: (26160, "double", 1.0),   # EQ_TOTAL_a_kvarh
    149: (26164, "double", 1.0),   # EQ_TOTAL_b_kvarh
    150: (26168, "double", 1.0),   # EQ_TOTAL_c_kvarh
    151: (26172, "double", 1.0),   # EQ_TOTAL_kvarh
    152: (26176, "double", 1.0),   # ES_a_kVAh
    153: (26180, "double", 1.0),   # ES_b_kVAh
    154: (26184, "double", 1.0),   # ES_c_kVAh
    155: (26188, "double", 1.0),   # ES_kVAh
}


def acuvim3_snmp_to_modbus(param_idx: int) -> int | None:
    addr, _, _ = _excel_resolve("Acuvim3", _MIB_ACUVIM3_PARAM_NAMES, param_idx)
    if addr is not None:
        return addr
    entry = _ACUVIM3_MAP.get(param_idx)
    return entry[0] if entry else None

def acuvim3_param_type(param_idx: int) -> str | None:
    _, typ, _ = _excel_resolve("Acuvim3", _MIB_ACUVIM3_PARAM_NAMES, param_idx)
    if typ is not None:
        return typ
    entry = _ACUVIM3_MAP.get(param_idx)
    return entry[1] if entry else None

def acuvim3_modbus_scale(param_idx: int) -> float:
    _, _, scale = _excel_resolve("Acuvim3", _MIB_ACUVIM3_PARAM_NAMES, param_idx)
    if scale is not None:
        return scale
    entry = _ACUVIM3_MAP.get(param_idx)
    return entry[2] if entry else 1.0

def acuvim3_param_name(param_idx: int) -> str:
    if param_idx in _MIB_ACUVIM3_PARAM_NAMES:
        return _MIB_ACUVIM3_PARAM_NAMES[param_idx]
    return f"Param#{param_idx}"

def acurev2100_param_desc(param_idx: int) -> str:
    return _excel_desc("AcuRev2100", _MIB_2100_PARAM_NAMES, param_idx)

def acuvim3_param_desc(param_idx: int) -> str:
    return _excel_desc("Acuvim3", _MIB_ACUVIM3_PARAM_NAMES, param_idx)


# ─── AcuvimIIW 完整映射（param_idx → (modbus_addr, type, scale)）──────────────
# float32: idx 2-45；uint32/int32: idx 37-45（能量）；int16: idx 46-91（THD/SEQ）
# scale≠1 的情况：功率×0.001(W→kW)，不平衡度×100，THD/角度/序分量 × 对应系数
_ACUVIMIIIW_MAP: dict[int, tuple] = {
      2: (16384, "float",  1.0),   # FREQ_Hz
      3: (16386, "float",  1.0),   # VLN_a_V
      4: (16388, "float",  1.0),   # VLN_b_V
      5: (16390, "float",  1.0),   # VLN_c_V
      6: (16392, "float",  1.0),   # VLN_avg_V
      7: (16394, "float",  1.0),   # VLL_ab_V
      8: (16396, "float",  1.0),   # VLL_bc_V
      9: (16398, "float",  1.0),   # VLL_ca_V
     10: (16400, "float",  1.0),   # VLL_avg_V
     11: (16402, "float",  1.0),   # I_a_A
     12: (16404, "float",  1.0),   # I_b_A
     13: (16406, "float",  1.0),   # I_c_A
     14: (16408, "float",  1.0),   # I_avg_A
     15: (16410, "float",  1.0),   # I_n_A
     16: (16412, "float",  0.001), # P_a_kW  (Modbus=W, SNMP=kW)
     17: (16414, "float",  0.001), # P_b_kW
     18: (16416, "float",  0.001), # P_c_kW
     19: (16418, "float",  0.001), # P_kW
     20: (16420, "float",  0.001), # Q_a_kvar
     21: (16422, "float",  0.001), # Q_b_kvar
     22: (16424, "float",  0.001), # Q_c_kvar
     23: (16426, "float",  0.001), # Q_kvar
     24: (16428, "float",  0.001), # S_a_kVA
     25: (16430, "float",  0.001), # S_b_kVA
     26: (16432, "float",  0.001), # S_c_kVA
     27: (16434, "float",  0.001), # S_kVA
     28: (16436, "float",  1.0),   # PF_a
     29: (16438, "float",  1.0),   # PF_b
     30: (16440, "float",  1.0),   # PF_c
     31: (16442, "float",  1.0),   # PF
     32: (16444, "float",  100.0), # UNBL_V_%  (Modbus=ratio, SNMP=%)
     33: (16446, "float",  100.0), # UNBL_I_%
     34: (16450, "float",  0.001), # DMD_P_kW  (跳过 LC_avg 在 16448)
     35: (16452, "float",  0.001), # DMD_Q_kvar
     36: (16454, "float",  0.001), # DMD_S_kVA
     37: (16456, "uint32", 0.1),   # EP_IMP_kWh
     38: (16458, "uint32", 0.1),   # EP_EXP_kWh
     39: (16460, "uint32", 0.1),   # EQ_IMP_kvarh
     40: (16462, "uint32", 0.1),   # EQ_EXP_kvarh
     41: (16464, "uint32", 0.1),   # EP_TOTAL_kWh
     42: (16466, "int32",  0.1),   # EP_NET_kWh   (净值，可为负）
     43: (16468, "uint32", 0.1),   # EQ_TOTAL_kvarh
     44: (16470, "int32",  0.1),   # EQ_NET_kvarh （净值，可为负）
     45: (16472, "uint32", 0.1),   # ES_kVAh
     46: (16474, "word",   0.01),  # THD_V_a_%
     47: (16475, "word",   0.01),  # THD_V_b_%
     48: (16476, "word",   0.01),  # THD_V_c_%
     49: (16477, "word",   0.01),  # THD_V_avg_%
     50: (16478, "word",   0.01),  # THD_I_a_%
     51: (16479, "word",   0.01),  # THD_I_b_%
     52: (16480, "word",   0.01),  # THD_I_c_%
     53: (16481, "word",   0.01),  # THD_I_avg_%
     54: (16512, "word",   0.01),  # OTHD_V_a_%
     55: (16513, "word",   0.01),  # ETHD_V_a_%
     56: (16514, "word",   0.001), # CF_V_a_%
     57: (16515, "word",   0.01),  # THFF_V_a_%
     58: (16546, "word",   0.01),  # OTHD_V_b_%
     59: (16547, "word",   0.01),  # ETHD_V_b_%
     60: (16548, "word",   0.001), # CF_V_b_%
     61: (16549, "word",   0.01),  # THFF_V_b_%
     62: (16580, "word",   0.01),  # OTHD_V_c_%
     63: (16581, "word",   0.01),  # ETHD_V_c_%
     64: (16582, "word",   0.001), # CF_V_c_%
     65: (16583, "word",   0.01),  # THFF_V_c_%
     66: (16614, "word",   0.01),  # OTHD_I_a_%
     67: (16615, "word",   0.01),  # ETHD_I_a_%
     68: (16616, "word",   0.1),   # KF_I_a_%
     69: (16647, "word",   0.01),  # OTHD_I_b_%
     70: (16648, "word",   0.01),  # ETHD_I_b_%
     71: (16649, "word",   0.1),   # KF_I_b_%
     72: (16680, "word",   0.01),  # OTHD_I_c_%
     73: (16681, "word",   0.01),  # ETHD_I_c_%
     74: (16682, "word",   0.1),   # KF_I_c_%
     75: (17044, "word_signed", 0.1),   # SEQ_POS_REAL_V
     76: (17045, "word_signed", 0.1),   # SEQ_POS_IMG_V
     77: (17046, "word_signed", 0.1),   # SEQ_NEG_REAL_V
     78: (17047, "word_signed", 0.1),   # SEQ_NEG_IMG_V
     79: (17048, "word_signed", 0.1),   # SEQ_ZERO_REAL_V
     80: (17049, "word_signed", 0.1),   # SEQ_ZERO_IMG_V
     81: (17050, "word_signed", 0.001), # SEQ_POS_REAL_I
     82: (17051, "word_signed", 0.001), # SEQ_POS_IMG_I
     83: (17052, "word_signed", 0.001), # SEQ_NEG_REAL_I
     84: (17053, "word_signed", 0.001), # SEQ_NEG_IMG_I
     85: (17054, "word_signed", 0.001), # SEQ_ZERO_REAL_I
     86: (17055, "word_signed", 0.001), # SEQ_ZERO_IMG_I
     87: (17056, "word",   0.1),   # ANG_VLN_b
     88: (17057, "word",   0.1),   # ANG_VLN_c
     89: (17058, "word",   0.1),   # ANG_I_a
     90: (17059, "word",   0.1),   # ANG_I_b
     91: (17060, "word",   0.1),   # ANG_I_c
     92: (17061, "word",   0.1),   # ANG_VLN_a  (Excel scale=None → fallback to this entry)
     93: (17952, "uint32", 0.1),   # EP_IMP_a_kWh
     94: (17954, "uint32", 0.1),   # EP_EXP_a_kWh
     95: (17956, "uint32", 0.1),   # EP_IMP_b_kWh
     96: (17958, "uint32", 0.1),   # EP_EXP_b_kWh
     97: (17960, "uint32", 0.1),   # EP_IMP_c_kWh
     98: (17962, "uint32", 0.1),   # EP_EXP_c_kWh
     99: (17964, "uint32", 0.1),   # EQ_IMP_a_kvarh
    100: (17966, "uint32", 0.1),   # EQ_EXP_a_kvarh
    101: (17968, "uint32", 0.1),   # EQ_IMP_b_kvarh
    102: (17970, "uint32", 0.1),   # EQ_EXP_b_kvarh
    103: (17972, "uint32", 0.1),   # EQ_IMP_c_kvarh
    104: (17974, "uint32", 0.1),   # EQ_EXP_c_kvarh
    105: (17976, "uint32", 0.1),   # ES_a_kVAh
    106: (17978, "uint32", 0.1),   # ES_b_kVAh
    107: (17980, "uint32", 0.1),   # ES_c_kVAh
}


def acuvimiIW_snmp_to_modbus(param_idx: int) -> int | None:
    addr, _, _ = _excel_resolve("AcuvimIIW", _MIB_ACUVIMIIIW_PARAM_NAMES, param_idx)
    if addr is not None:
        return addr
    entry = _ACUVIMIIIW_MAP.get(param_idx)
    return entry[0] if entry else None

def acuvimiIW_param_type(param_idx: int) -> str | None:
    _, typ, _ = _excel_resolve("AcuvimIIW", _MIB_ACUVIMIIIW_PARAM_NAMES, param_idx)
    if typ is not None:
        return typ
    entry = _ACUVIMIIIW_MAP.get(param_idx)
    return entry[1] if entry else None

def acuvimiIW_modbus_scale(param_idx: int) -> float:
    _, _, scale = _excel_resolve("AcuvimIIW", _MIB_ACUVIMIIIW_PARAM_NAMES, param_idx)
    if scale is not None:
        return scale
    entry = _ACUVIMIIIW_MAP.get(param_idx)
    return entry[2] if entry else 1.0

def acuvimiIW_param_name(param_idx: int) -> str:
    if param_idx in _MIB_ACUVIMIIIW_PARAM_NAMES:
        return _MIB_ACUVIMIIIW_PARAM_NAMES[param_idx]
    return f"Param#{param_idx}"

def acuvimiIW_param_desc(param_idx: int) -> str:
    return _excel_desc("AcuvimIIW", _MIB_ACUVIMIIIW_PARAM_NAMES, param_idx)


# ─── AcuvimIIR 完整映射（param_idx → (modbus_addr, type, scale)）──────────────
# IIR 与 IIW 寄存器地址相同（已通过 Excel 地址表核实）。
# idx 93-107 在 Excel 中有完整记录，运行时走 Excel 路径；此处 fallback 与 IIW 一致。
_ACUVIMIIIR_MAP: dict[int, tuple] = {
    **_ACUVIMIIIW_MAP,
}


def acuvimiIR_snmp_to_modbus(param_idx: int) -> int | None:
    addr, _, _ = _excel_resolve("AcuvimIIR", _MIB_ACUVIMIIIR_PARAM_NAMES, param_idx)
    if addr is not None:
        return addr
    entry = _ACUVIMIIIR_MAP.get(param_idx)
    return entry[0] if entry else None

def acuvimiIR_param_type(param_idx: int) -> str | None:
    _, typ, _ = _excel_resolve("AcuvimIIR", _MIB_ACUVIMIIIR_PARAM_NAMES, param_idx)
    if typ is not None:
        return typ
    entry = _ACUVIMIIIR_MAP.get(param_idx)
    return entry[1] if entry else None

def acuvimiIR_modbus_scale(param_idx: int) -> float:
    _, _, scale = _excel_resolve("AcuvimIIR", _MIB_ACUVIMIIIR_PARAM_NAMES, param_idx)
    if scale is not None:
        return scale
    entry = _ACUVIMIIIR_MAP.get(param_idx)
    return entry[2] if entry else 1.0

def acuvimiIR_param_name(param_idx: int) -> str:
    if param_idx in _MIB_ACUVIMIIIR_PARAM_NAMES:
        return _MIB_ACUVIMIIIR_PARAM_NAMES[param_idx]
    return f"Param#{param_idx}"

def acuvimiIR_param_desc(param_idx: int) -> str:
    return _excel_desc("AcuvimIIR", _MIB_ACUVIMIIIR_PARAM_NAMES, param_idx)


# ─── AcuRev1300 映射（param_idx → Modbus，完全走 Excel 路径）────────────────
# 无硬编码地址表，100% 依赖 data/map/AcuRev-1300_v1.01_20260416.xlsx

def acurev1300_snmp_to_modbus(param_idx: int) -> int | None:
    addr, _, _ = _excel_resolve("AcuRev1300", _MIB_ACUREV1300_PARAM_NAMES, param_idx)
    return addr

def acurev1300_param_type(param_idx: int) -> str | None:
    _, typ, _ = _excel_resolve("AcuRev1300", _MIB_ACUREV1300_PARAM_NAMES, param_idx)
    return typ

def acurev1300_modbus_scale(param_idx: int) -> float:
    _, _, scale = _excel_resolve("AcuRev1300", _MIB_ACUREV1300_PARAM_NAMES, param_idx)
    return scale if scale is not None else 1.0

def acurev1300_param_name(param_idx: int) -> str:
    if param_idx in _MIB_ACUREV1300_PARAM_NAMES:
        return _MIB_ACUREV1300_PARAM_NAMES[param_idx]
    return f"Param#{param_idx}"

def acurev1300_param_desc(param_idx: int) -> str:
    return _excel_desc("AcuRev1300", _MIB_ACUREV1300_PARAM_NAMES, param_idx)


# ─── 设备映射分发器 ──────────────────────────────────────────────────────────
# (snmp_to_modbus_fn, param_type_fn, param_name_fn, modbus_scale_fn, param_desc_fn)
DEVICE_MAPPING_FNS = {
    "AcuRev4100": (acurev4100_snmp_to_modbus, acurev4100_param_type,
                   acurev4100_param_name,      acurev4100_modbus_scale,
                   acurev4100_param_desc),
    "AcuRev2100": (acurev2100_snmp_to_modbus, acurev2100_param_type,
                   acurev2100_param_name,      acurev2100_modbus_scale,
                   acurev2100_param_desc),
    "Acuvim3":    (acuvim3_snmp_to_modbus,    acuvim3_param_type,
                   acuvim3_param_name,         acuvim3_modbus_scale,
                   acuvim3_param_desc),
    "AcuvimIIW":  (acuvimiIW_snmp_to_modbus,  acuvimiIW_param_type,
                   acuvimiIW_param_name,       acuvimiIW_modbus_scale,
                   acuvimiIW_param_desc),
    "AcuvimIIR":  (acuvimiIR_snmp_to_modbus,  acuvimiIR_param_type,
                   acuvimiIR_param_name,       acuvimiIR_modbus_scale,
                   acuvimiIR_param_desc),
    "AcuRev1300": (acurev1300_snmp_to_modbus, acurev1300_param_type,
                   acurev1300_param_name,      acurev1300_modbus_scale,
                   acurev1300_param_desc),
}


def get_device_fns(device_name: str):
    """返回设备对应的 (to_modbus, param_type, param_name, scale, param_desc) 五元组。
    先按实例名查，找不到则按 mib_mapping.json 中的 model_type 回退。
    AcuRev4100b 等同型号多实例共用同一套映射函数。
    """
    fns = DEVICE_MAPPING_FNS.get(device_name)
    if fns is None:
        model_type = get_model_type(device_name)
        if model_type:
            fns = DEVICE_MAPPING_FNS.get(model_type)
    return fns
