"""项目配置读取（projects/AcuRev100/config.yaml 单一配置源）。

本工具原设计为"完全独立、不依赖 autotest 工程"，适配 ACmeter(AcuRev-100) 后需要与
项目已裁定的口径保持一致（精度容差、源安全门禁、台体 CT 换算基准、串口号），故改为
**优先读项目 config.yaml，读不到则回退内置默认值**——脱离工程单独拷走仍可运行。

单一配置源: projects/AcuRev100/config.yaml
  accuracy.quantities → 精度容差（pct=相对% / abs=绝对量）
  source             → CL3021 源参数 + 电流硬限幅 + A相供电护栏 + 档位钉死 + 台体CT换算
  transport.rtu      → 校验口串口参数（电表 USB 口）
  device             → 型号 / CT 类型（决定 GUI 可选项与测点 sheet）
"""
from __future__ import annotations

import logging
import os

try:
    import yaml
except ImportError:                                   # 独立拷走且未装 pyyaml 时退回默认值
    yaml = None

# tools/accuracy_quick/core/ → 上溯 3 级到 projects/AcuRev100/
_PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
CONFIG_PATH = os.path.join(_PROJECT_DIR, "config.yaml")

# ── 内置回退默认值（与 config.yaml 现值一致；改判据请改 config.yaml，勿改此处）──
DEFAULT_ACCURACY = {
    "voltage":        ("pct", 0.2),
    "current":        ("pct", 0.2),
    "active_power":   ("pct", 0.2),
    "reactive_power": ("pct", 0.5),
    "apparent_power": ("pct", 0.5),
    "phase_angle":    ("abs", 0.5),
    "frequency":      ("abs", 0.1),
}
DEFAULT_SOURCE = {
    "host": "192.168.0.50",
    "port": 10003,
    "local_port": 10005,
    "settle_s": 2.0,
    "current_injection": "via_ct",
    "max_current_a": {"via_ct": 25.0, "direct": 0.1},
    "max_current_a_phase": {"a": 20.0, "b": 20.0, "c": 15.0},
    "over_phase_cap": "skip",
    "supply_guard": {"phase_a_min_v": 100.0, "phase_a_keepalive_v": 100.0},
    "send_gear_frames": True,
    "gear_pin": {"voltage_v": 480.0, "current_a": 20.0},
    "precision_tool_gear_pin": False,
    "bench_ct_a": {"mA": 20.0, "mV": 5.0},
}
DEFAULT_RTU = {"port": "COM7", "baudrate": 19200, "slave_id": 1}
DEFAULT_DEVICE = {"model": "AcuRev-101-mA", "ct_type": "100mA"}

# CT 类型 → 台体 CT 键名（bench_ct_a / 测点 sheet 的 mA / mV 归类）
CT_FAMILY = {"100mA": "mA", "80mA": "mA", "333mV": "mV", "RCT": "mV"}
# 型号 → 该型号允许的 CT 类型（型号不可跨切，见 knowledge context 共存约束）。
# 80mA / RCT 的寄存器值固件侧未确认，暂不放开（addr_loader.CT_TYPE_MAP 亦只有两项）。
MODEL_CT_OPTIONS = {
    "AcuRev-101-mA": ["100mA"],
    "AcuRev-101-mV": ["333mV"],
}

_cache: dict | None = None


def load() -> dict:
    """读 config.yaml（带缓存）。文件缺失/解析失败返回空 dict，各 getter 自行回退默认。"""
    global _cache
    if _cache is not None:
        return _cache
    _cache = {}
    if yaml is None:
        logging.warning("未安装 pyyaml，项目配置不可读，全部使用内置默认值")
        return _cache
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fp:
            _cache = yaml.safe_load(fp) or {}
        logging.info("已读取项目配置: %s", CONFIG_PATH)
    except OSError as e:
        logging.warning("项目配置不可读(%s)，使用内置默认值: %s", CONFIG_PATH, e)
    except yaml.YAMLError as e:
        logging.error("项目配置解析失败，使用内置默认值: %s", e)
    return _cache


def source_cfg() -> dict:
    """source 段（缺项按 DEFAULT_SOURCE 补齐）。"""
    cfg = dict(DEFAULT_SOURCE)
    cfg.update(load().get("source") or {})
    for key in ("max_current_a", "max_current_a_phase", "supply_guard", "gear_pin", "bench_ct_a"):
        merged = dict(DEFAULT_SOURCE[key])
        merged.update(cfg.get(key) or {})
        cfg[key] = merged
    return cfg


def rtu_cfg() -> dict:
    """transport.rtu 段 = 电表 USB 校验口（RS-485 口被 Acuview2 占用）。"""
    cfg = dict(DEFAULT_RTU)
    cfg.update((load().get("transport") or {}).get("rtu") or {})
    return cfg


def device_cfg() -> dict:
    cfg = dict(DEFAULT_DEVICE)
    cfg.update(load().get("device") or {})
    return cfg


def max_current_a() -> float:
    """电流硬限幅：按 source.current_injection 选 via_ct / direct 上限（烧板事故后的硬门禁）。"""
    s = source_cfg()
    caps = s["max_current_a"]
    return float(caps.get(s.get("current_injection", "via_ct"), caps["direct"]))


def max_current_a_phase() -> dict:
    """逐相源侧电流上限（A）：台体各相回路承载不同，且不超过全局限幅。

    2026-07-27 实机：3E4WY 三相同出 20A 时源报 "Ic 过载" → C 相上限 15A。
    """
    cap = max_current_a()
    per = source_cfg()["max_current_a_phase"]
    return {ph: min(float(per.get(ph, cap)), cap) for ph in ("a", "b", "c")}


def gear_pin() -> dict:
    """本工具用的档位钉死参数：默认关闭（逐点档位以保源精度）。

    关闭理由（2026-07-27 裁定）：档位切换本身不掉源输出（掉输出的只有频率切换），
    且工具每个测点都先降到 0A 再切换 ⇒ 逐点档位无风险。
    而钉档会让低幅值点（480V 档打 120V、20A 档打 0.02A）的源输出误差达到 0.2% 量级，
    直接把电表的精度判定带偏——精度测试里这是不可接受的。
    """
    if not bool(source_cfg().get("precision_tool_gear_pin", False)):
        return {"voltage_v": 0.0, "current_a": 0.0}
    pin = source_cfg()["gear_pin"] or {}
    return {"voltage_v": float(pin.get("voltage_v") or 0.0),
            "current_a": float(pin.get("current_a") or 0.0)}


def over_phase_cap() -> str:
    """超逐相上限的测点处置策略：'skip'（跳过并留证，默认）/ 'abort'（中止本批）。"""
    mode = str(source_cfg().get("over_phase_cap", "skip")).strip().lower()
    return mode if mode in ("skip", "abort") else "skip"


def phase_cap_violations(src_point: dict) -> list[str]:
    """返回该源测点超逐相上限的说明列表（空=未超限）。

    逐相上限描述的是**当前这台源的个体能力**（本台 Ic 出不到 20A），不是测点/产品口径；
    因此判断放在运行时，测点矩阵始终按满载编写。
    """
    per = max_current_a_phase()
    out = []
    for ph in ("a", "b", "c"):
        val = abs(float(src_point.get(f"i{ph}", 0.0)))
        if val > per[ph]:
            out.append(f"I{ph.upper()}={val:g}A 超本源该相上限 {per[ph]:g}A")
    return out


def bench_ct_a(ct_type: str) -> float:
    """台体降流 CT 额定一次侧电流；未知 CT 类型按 mA 回路处理。"""
    return float(source_cfg()["bench_ct_a"].get(CT_FAMILY.get(ct_type, "mA"), 20.0))


def current_scale(ct_type: str, ct_primary: float) -> float:
    """源侧输出电流 → 电表期望读数的换算系数 = CT Primary ÷ 台体 CT 额定。"""
    base = bench_ct_a(ct_type)
    if base <= 0:
        return 1.0
    return float(ct_primary) / base


def supply_guard() -> tuple[float, float]:
    """(A相工作下限V, A相保活/强制V)。自供电表: A相电压=电表电源，低于下限即断电。"""
    g = source_cfg()["supply_guard"]
    return float(g.get("phase_a_min_v", 100.0)), float(g.get("phase_a_keepalive_v", 100.0))


def ct_options(model: str | None = None) -> list[str]:
    """按型号返回允许的 CT 类型；未知型号回退两型全给。"""
    model = model or device_cfg().get("model", "")
    return list(MODEL_CT_OPTIONS.get(model, ["100mA", "333mV"]))


def accuracy_thresholds() -> dict:
    """config.accuracy → 测点表列口径的阈值字典。

    返回 {v_acc, i_acc, angle_acc, p_acc, q_acc, s_acc}；
    pct 型转成比值（0.2% → 0.002），abs 型（相角）保持绝对值（°）。
    """
    quantities = dict(DEFAULT_ACCURACY)
    for name, item in ((load().get("accuracy") or {}).get("quantities") or {}).items():
        if isinstance(item, dict) and "value" in item:
            quantities[name] = (str(item.get("type", "pct")), float(item["value"]))

    def _ratio(name: str, fallback: float) -> float:
        kind, val = quantities.get(name, ("pct", fallback * 100))
        return val / 100.0 if kind == "pct" else val

    return {
        "v_acc":     _ratio("voltage", 0.002),
        "i_acc":     _ratio("current", 0.002),
        "p_acc":     _ratio("active_power", 0.002),
        "q_acc":     _ratio("reactive_power", 0.005),
        "s_acc":     _ratio("apparent_power", 0.005),
        # 相角是绝对量（°），不做百分比换算
        "angle_acc": quantities.get("phase_angle", ("abs", 0.5))[1],
    }


def cfg_restart_window_s() -> float:
    """写电表频率选择后"等延迟重启出现"的窗口秒数（config run.cfg_restart_window_s）。

    2026-07-13 记录写 0x1041 后 30~60s 才重启 → 原口径 90s；2026-07-27 两轮实测全程未见重启、
    纯等 91s，且电表在电压恢复后 <2s 即可工作，故默认 10s。
    安全网：重启若砸在采样期，test_engine 会自动重采本点。
    """
    return float((load().get("run") or {}).get("cfg_restart_window_s", 10.0))


def default_pct() -> float:
    """无精度要求点的定性口径（config accuracy.default_pct，默认 ±10%）→ 比值。"""
    return float((load().get("accuracy") or {}).get("default_pct", 10.0)) / 100.0
