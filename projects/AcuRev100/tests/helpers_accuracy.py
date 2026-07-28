"""AcuRev-100 精度类用例运行时（source_read / energy_accumulate 共用）。

职责: 读 case_map.yaml 中某条用例的 points → 逐点 CL3021 设源 → 稳定等待 →
Modbus(USB 校验口) 读寄存器 → 区间断言 → 汇总 Report。

设计说明(与 comm/ctl_acuview 引擎的关系):
- 引擎的 run_read_compare_case 面向"单控件 GUI OCR vs 单寄存器", 不含控源、
  也不支持多测点区间断言, 故精度类用例走本 helper(方案A: Modbus 断言为主)。
- GUI 显示比对(方案B)留了钩子 gui_check_page: 待真机补齐 Real-Time 表格行坐标
  与 Tesseract 后, 在 assert_point 里追加 OCR 比对即可, 用例文件无需改动。
- 还原: 用例结束把源归零(set_zero), 不改电表配置 → 天然可逆。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[1]          # projects/AcuRev100 (本文件在 tests/ 根)
REPO_ROOT = PROJECT_ROOT.parents[1]
if str(REPO_ROOT) not in sys.path:                           # 直跑本文件时兜底; pytest 由根 conftest 注入
    sys.path.insert(0, str(REPO_ROOT))

from comm.ctl_acuview.config import get_config  # noqa: E402
from comm.ctl_acuview.verify import Check, Report, Verifier, now  # noqa: E402
from projects.AcuRev100.tools.accuracy_quick.core.source_comm import SourceUdp  # noqa: E402
#   ↑ CL3021 UDP 控源(逐帧+ACK+档位钉死+同值角度帧跳发)。本类原在本文件内实现并在台面调通,
#     2026-07-27 下沉到 tools/accuracy_quick/core/source_comm.py, tests 与快速精度工具共用一份。

CASE_MAP = Path(__file__).resolve().parent / "case_map.yaml"
CONFIG = str(PROJECT_ROOT / "config.yaml")     # 项目单一配置源(engine + source 段都在这里)

# ⚠️ 自供电表: 电表从 CL3021 UaUbUcUn 取电, 电压归零=给电表断电。
# 所有还原/空闲动作保持 KEEPALIVE(220V/50Hz/0A); 源切换可能瞬时断电触发电表重启,
# 每次设源后必须 _wait_alive 等电表 Modbus 复活再读数。
# 🔴 A相供电保障(2026-07-14 用户指示, 最高优先级): A相电压取 config source.supply_guard,
#    点位 A相 < phase_a_min_v 时强制抬到 phase_a_keepalive_v(见 _guard_point)。
KEEPALIVE = dict(qua=0.0, qub=240.0, quc=120.0, qia=0.0, qib=240.0, qic=120.0,
                 ua=220.0, ub=220.0, uc=220.0, ia=0.0, ib=0.0, ic=0.0, freq=50.0)
#   ↑ 基线模板; 实际保活电压由 _keepalive() 按 config supply_guard + 当前接线方式重算,
#     不要直接用本字典下发(2026-07-28: 硬编码 220V 打进 B/C 是过载锁存的诱因, 见 _keepalive)。

# 各接线方式下"该带电的电压相"(与 accuracy_quick 引擎 VOLT_PHASES 同一口径)。
# 🔴 未接线相必须零输出: 台面未接线相的电压端子并接/接 N, 加压等于近似短路, 源会报
#    "Ub/Uc 过载"并锁存保护(2026-07-27 实机记录), 程序侧切屏重臂清不掉, 只能人工到面板复位。
VOLT_PHASES = {
    "1E2W":   ("a",),
    "2E3W1P": ("a", "c"),
    "3E4WY":  ("a", "b", "c"),
}
_WIRING_CURRENT: "str | None" = None      # 当前用例的接线方式(用例开头/写配置时登记)


def _set_wiring_context(wiring: "str | None") -> None:
    """登记当前用例的接线方式, 供 _keepalive() 决定哪几相该带电。"""
    global _WIRING_CURRENT
    if wiring:
        _WIRING_CURRENT = str(wiring)
BOOT_TIMEOUT_S = 240       # 电表冷启动到 Modbus 可用的最大等待(2026-07-13 实测约 3.5min)
WARM_PROBE_S = 20          # 暖启动探测: 源已在出力时电表应在线, 20s 读不通转冷初始化
DEAD_DETECT_S = 5          # 设点后无应答判定: 电表重启本身很快(2026-07-14 用户台面实证),
#                            读不通≈源输出掉0卡死(低压区幅值切换偶发), 必须立刻救源
#                            (重发本测点→切屏冷重臂), 傻等电表只会让源死更久——源不出电表起不来
#                            (12→8→5s: 2026-07-14 用户两次指示压救源响应; source_diag 实证
#                             set_point 瞬断均自恢复且电表在线, 5s 误救风险低)
REARM_BOOT_S = 45          # 救源(重发/冷重臂)后等电表复活的窗口: 源恢复出力后电表几秒即起, 留裕量
SAMPLE_N_DEFAULT = 10      # 每条判据的连采次数(2026-07-27 用户定): 快速连读 N 次, 全部落区间才 PASS。
#                            单次采样打在抖动量上是抛硬币(002 批实证: 频率 50.0/50.1 交替、
#                            电流在判据边界 3.004~3.008 浮动), 连采既抓抖动又给出极值证据。
#                            config run.sample_n 可覆盖; 置 1 回退旧单次口径。
MEASURE_READY_S = 7        # 设源后测量就绪等待: A相常驻有压 → 频率读值>10 才允许读判据(2026-07-16 换板后 20→7 提速, 用户指定)
#                            (2026-07-14 批跑教训: 档位切换后测量读零期读数 → 20V点误FAIL、
#                             期望0的判据被"测量死机的0"凑成空心PASS。45→20s: ADC损坏排查
#                             模式下读零是常态, 缩短等待加快批跑, 判据仍 FAIL 留证不误判)


def _supply_guard_cfg() -> "tuple[float, float]":
    """读 config source.supply_guard: (A相工作下限V, A相保活/强制V)。缺省 100/220。"""
    s = yaml.safe_load(Path(CONFIG).read_text(encoding="utf-8")).get("source", {}) or {}
    g = s.get("supply_guard", {}) or {}
    return float(g.get("phase_a_min_v", 100.0)), float(g.get("phase_a_keepalive_v", 220.0))


def _keepalive(src: "SourceUdp | None" = None) -> dict:
    """保活点: 带电相取 config supply_guard 的保活电压, 未接线相零输出, 电流全 0。

    🔴 2026-07-28 整改(过载锁存事故): 原来 KEEPALIVE 把 Ub/Uc 硬编码成 220V, 于是每次退场、
    每次救源都往 B/C 打 220V ——
      · 与用例电压无关: 100V 的用例退场反而把 B/C 抬到 220V(本次事故正是 100V 点退场时
        源报 "Ub/Uc 过载" 并锁存, 表断电, 程序侧拉不回, 只能人工到 CL3021 面板复位);
      · 与接线方式无关: 1E2W 的 B/C、2E3W1P 的 B 本该零输出, 加压等于近似短路(README
        「未接线相必须零输出」);
      · 而保活的唯一目的是给自供电表供 A 相电, B/C 加压纯属多余动作。
    现按 accuracy_quick 引擎 _keepalive_point() 的同一口径: 带电相 keep_v、未接线相 0V。

    🔴 传了 src 就沿用源当前角度(2026-07-28): 角度帧一发电表必重启一次, 而保活点只关心
    "A 相有压、电流为 0", 角度是什么无所谓 —— 把角度硬拉回标准三相纯属白赔一次重启。
    跨进程一致性由落盘状态(SourceUdp.state_path)保证, 不再依赖"台面收在标准三相"。
    不传 src 时退回标准三相角度(helpers_wiring/helpers_timing 的既有调用口径不变)。
    """
    keep_v = _supply_guard_cfg()[1]
    wired = VOLT_PHASES.get(_WIRING_CURRENT or "", ("a", "b", "c"))
    ka = dict(KEEPALIVE)
    ka["ua"] = keep_v
    ka["ub"] = keep_v if "b" in wired else 0.0
    ka["uc"] = keep_v if "c" in wired else 0.0
    if src is None:
        return ka
    if src.last_point:
        for k in ("qua", "qub", "quc", "qia", "qib", "qic", "freq"):
            ka[k] = src.last_point[k]
    elif src.current_angles():
        # 进程首次保活: last_point 还是 None, 但缓存已由落盘状态/assume_angles 预置。
        # 这里若漏看缓存就会回落到标准三相 → 开场白发一次角度帧 = 白赔一次重启
        # (2026-07-28 实测踩到: 落盘状态明明与用例角度一致, 却仍重启了两次)。
        for k, v in zip(("qua", "qub", "quc", "qia", "qib", "qic", "freq"),
                        src.current_angles()):
            ka[k] = v
    return ka


def _wire_guard(s: dict) -> dict:
    """未接线相强制零输出(与 accuracy_quick 引擎 _source_point 同规则, 2026-07-28 补到 tests 路径)。

    🔴 台面未接线相的电压端子是并接/接 N 的, 给它加压等于近似短路 → 源报 "Ub/Uc 过载" 并
    锁存保护, 程序侧清不掉, 只能人工到 CL3021 面板复位(2026-07-27 实机 + 2026-07-28 复现两次)。
    而 case_map 里 1E2W/2E3W1P 用例的测点普遍带着三相电压(生成时按 xlsx 原文填的), 其中
    未接线相纯属"陪跑"——009_01_case1 的 decisions 原话: "1E2W 仅A相计量: 断言 Va/Ia/P_A/P_SYS;
    B/C 源输出仅陪跑" ⇒ 清零不影响任何判据, 只去掉过载风险。
    扫描结果(2026-07-28): 全库 10 个测点命中(其中 009_01_case1、009_02_case1/2/4 可跑)。
    """
    wired = VOLT_PHASES.get(_WIRING_CURRENT or "", ("a", "b", "c"))
    off = [ph for ph in ("a", "b", "c")
           if ph not in wired and (s[f"u{ph}"] or s[f"i{ph}"])]
    if not off:
        return s
    g = dict(s)
    for ph in off:
        g[f"u{ph}"] = 0.0
        g[f"i{ph}"] = 0.0
    print(f"[wire] {_WIRING_CURRENT} 未接线相 {'/'.join(p.upper() for p in off)} 强制零输出 "
          f"(未接线相端子并接/接N, 加压=近似短路→源过载锁存)")
    return g


def _guard_point(s: dict) -> dict:
    """点位护栏: ① A相供电(最高优先级, 低于工作下限强制抬压) ② 未接线相强制零输出。"""
    min_v, keep_v = _supply_guard_cfg()
    if s["ua"] < min_v:
        g = dict(s)
        g["ua"] = keep_v
        print(f"[guard] A相 {s['ua']}V < 工作下限 {min_v}V → 强制 {keep_v}V "
              f"(config source.supply_guard, 电表供电最高优先级)")
        return _wire_guard(g)
    return _wire_guard(s)


def _wait_alive(config_path: str, timeout_s: float = BOOT_TIMEOUT_S) -> float:
    """轮询 Slave ID 寄存器直到电表 Modbus 复活; 返回耗时秒, 超时抛 RuntimeError。"""
    import yaml as _yaml
    from pymodbus.client import ModbusSerialClient
    rtu = _yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))["transport"]["rtu"]
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        c = ModbusSerialClient(port=rtu["port"], baudrate=rtu["baudrate"], parity=rtu["parity"],
                               bytesize=rtu["bytesize"], stopbits=rtu["stopbits"], timeout=1)
        try:
            if c.connect():
                r = c.read_holding_registers(4111, count=1, device_id=rtu["slave_id"])
                if r and not r.isError():
                    return time.time() - t0
        except Exception:
            pass
        finally:
            c.close()
        time.sleep(3)
    raise RuntimeError(f"电表 {rtu['port']} 在 {timeout_s}s 内未复活(自供电表: 检查源电压输出)")


def _wait_alive_via(vf: Verifier, timeout_s: float = BOOT_TIMEOUT_S) -> float:
    """复用 Verifier 已打开的串口轮询电表复活(串口被 Verifier 独占时用这个)。"""
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        try:
            vf.read_truth(4111)          # Modbus Slave ID, 读通即视为复活
            return time.time() - t0
        except Exception:
            time.sleep(3)
    raise RuntimeError(f"电表在 {timeout_s}s 内未复活(自供电表: 检查源电压输出)")


MEASURE_HOLD_S = 4         # 就绪判定保持时长: 频率连续有效满 4s 才放行(吃掉重启后 RMS 爬坡期,
#                            2026-07-14 用户要求: 读数必须不受板子瞬断重启影响)


def _exit_idle(src, config_path: str) -> None:
    """用例退场(2026-07-14 用户指示): 统一恢复保活(Ua=100V最低保活电压, 电流0)并探活校验。

    档位钉死后切电压幅值无切档瞬断, 退场态不再"保持最后测点电压"; 退场后探活,
    源输出停0(批跑结束后无人救的老毛病)就地冷重臂拉回, 拉不回置顶告警留给人工。

    🔴 退场不再把角度拉回标准三相(2026-07-28 改): 角度帧一发电表必重启一次, 而还原的唯一
    目的是保住 config assume_angles 的"台面收在标准角度"前提 —— 那个前提已由落盘状态
    (SourceUdp.state_path, 每次角度帧下发即更新)取代, 于是每条非标角度用例白赔的那次重启
    就此省掉。退场保活沿用当前角度, 只把电流归零、A 相保压。
    探活窗口按"是否真要切角度"取 REARM_BOOT_S / DEAD_DETECT_S: 真切了就必然重启一次, 别把
    它误判成源掉输出而去冷重臂(切屏帧=又一次重启, 正是 005/006 掉源被放大的一环)。
    """
    ka = _keepalive(src)
    restoring = src.angles_pending(ka)
    probe_s = REARM_BOOT_S if restoring else DEAD_DETECT_S
    base = _run_time_direct_edge(config_path)
    try:
        src.set_point(ka, settle_s=2)
        _wait_alive(config_path, timeout_s=probe_s)
        cur = _run_time_direct_edge(config_path)
        if base is not None and cur is not None:
            lost = (cur[1] - base[1]) - (cur[0] - base[0])
            print(f"[退场] 角度{'还原标准三相' if restoring else '未动'}, "
                  + (f"实测掉电 {lost:.1f}s(电表重启一次)" if lost > REBOOT_LOST_S
                     else f"run_time 连续(丢失 {lost:.1f}s), 未掉电"))
    except Exception:
        try:
            print("[退场] 电表不在线(源疑似输出停0) → 源冷重臂拉回保活")
            src.reinit_output()
            src.set_point(_keepalive(src), settle_s=3)
            _wait_alive(config_path)
        except Exception as exc:
            print(f"[退场] ⚠️⚠️ 源保活拉回失败, 电表处于断电状态, 需人工到台面开源: {exc}")


def _wait_measure_ready(vf: Verifier, timeout_s: float = MEASURE_READY_S) -> bool:
    """等测量块就绪: A相常驻有压 → SYSTEM_FREQUENCY>10 且连续保持 MEASURE_HOLD_S 秒。

    Modbus 应答≠测量可用: 电表只由 Va/Vn 供电(USB 口不供电), 通信起来后测量块可能仍在
    启动读零, 撞上"异常掉电重启后测量恒0"缺陷时更是长期读零(2026-07-27 用户澄清: 早前
    "USB 可维持 MCU 应答"的说法不成立, Modbus 能答即 Va 有电);
    读判据前必须过此门, 否则期望 0 的判据会被"测量死机的 0"凑成空心 PASS;
    保持时长防止"重启后首个有效样本"放行时 RMS 还在爬坡造成假 FAIL。
    """
    t0 = time.time()
    ok_since = None
    while time.time() - t0 < timeout_s:
        try:
            ok = float(vf.read_truth("SYSTEM_FREQUENCY")) > 10.0
        except Exception:
            ok = False
        now_ts = time.time()
        if ok:
            if ok_since is None:
                ok_since = now_ts
            if now_ts - ok_since >= MEASURE_HOLD_S:
                return True
        else:
            ok_since = None
        time.sleep(1.5)
    return False


RUN_TIME_REG = "DEVICE_RUN_TIME"   # 4121 uint32, 每秒 +1; 重启不清零, 但每次重启丢约 1s
REBOOT_LOST_S = 0.6                # "丢失秒数"判重启阈值, 由实测标定(2026-07-28, 对齐跳变取样后):
#                                    空窗本底 4 次 = -0.10 ~ +0.00s; 真角度切换 = +0.97/+1.08s
#                                    (另有一次 +1586.97s = 彻底冷启动, run_time 被清零)。
#                                    ⚠️ 两处教训: ① 早前拍脑袋定 2.0s, 把台面目击的 6 次真重启
#                                    全判成"未掉电"; ② 不对齐跳变时本底就有 ±0.5s, 与约 1s 的
#                                    重启信号同量级, 同一次切换能算出 0.3s 也能算出 1.9s。
#                                    改阈值前先跑空窗基线, 别凭感觉给裕量。


def _run_time(vf: Verifier) -> "float | None":
    """读电表运行秒表; 读不通返回 None。"""
    try:
        return float(vf.read_truth(RUN_TIME_REG))
    except Exception:
        return None


def _run_time_direct_edge(config_path: str,
                          timeout_s: float = 8.0) -> "tuple[float, float] | None":
    """不经 Verifier 取秒表跳变样本(退场阶段串口已释放时用), 口径同 _run_time_edge。

    退场也必须对齐跳变: 不对齐时 ±0.5s 的量化噪声会让"退场到底掉没掉电"变成掷硬币
    (2026-07-28 实测: 同样是"角度未动"的退场, 一次算 0.5s 一次算 0.7s, 跨在 0.6s 门限两边)。
    """
    import struct
    import yaml as _yaml
    from pymodbus.client import ModbusSerialClient
    rtu = _yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))["transport"]["rtu"]
    c = ModbusSerialClient(port=rtu["port"], baudrate=rtu["baudrate"], parity=rtu["parity"],
                           bytesize=rtu["bytesize"], stopbits=rtu["stopbits"], timeout=1)

    def _read() -> "float | None":
        try:
            r = c.read_holding_registers(4121, count=2, device_id=rtu["slave_id"])
            if not r or r.isError():
                return None
            return float(struct.unpack(">I", struct.pack(">HH", *r.registers))[0])
        except Exception:
            return None

    try:
        if not c.connect():
            return None
        t_end = time.time() + timeout_s
        first = _read()
        while time.time() < t_end:
            cur = _read()
            if cur is not None:
                if first is None:
                    first = cur
                elif cur != first:
                    return cur, time.time()
            time.sleep(0.05)
        return (first, time.time()) if first is not None else None
    finally:
        c.close()


def _run_time_edge(vf: Verifier, timeout_s: float = 2.5) -> "tuple[float, float] | None":
    """取样对齐到秒表跳变的那一刻, 返回 (跳变后的值, 观测到跳变的墙钟)。

    🔴 为什么必须对齐(2026-07-28 标定): run_time 是 1s 量化, 随手取样的误差就有 ±0.5s;
    而电表冷启动只要约 1s —— 信号和噪声同量级, 单次事件根本分不开(实测同一个角度切换,
    有时算出丢 0.3s 有时 1.9s)。等到计数跳变再记时刻, 两端都这么取, 量化误差压到一个轮询
    周期(≈50ms), 无重启窗口的"丢失秒数"就稳定在 0 附近, 一次重启即可分辨。
    """
    t_end = time.time() + timeout_s
    first = _run_time(vf)
    while time.time() < t_end:
        cur = _run_time(vf)
        if cur is not None:
            if first is None:
                first = cur
            elif cur != first:
                return cur, time.time()
        time.sleep(0.05)
    return (first, time.time()) if first is not None else None


def _reboot_probe(vf: Verifier) -> "tuple[float, float] | None":
    """开窗取基准 (run_time, 墙钟), 对齐到秒表跳变。None = 探针不可用(不影响主流程)。"""
    return _run_time_edge(vf)


def _downtime_since(vf: Verifier, base: "tuple[float, float] | None") -> "float | None":
    """闭窗算"丢失秒数" = 墙钟流逝 − run_time 增量; >REBOOT_LOST_S 即期间掉过电。

    ⚠️ 阈值必须靠本底标定, 不能拍脑袋(2026-07-28 教训): run_time 是 1s 量化, 空窗本底就有
    ±0.5s 抖动; 而电表冷启动只要 1s 级, 一次重启只丢 0.7~1.9s —— 早前把门限拍成 2.0s,
    结果 6 次台面目击的真重启被全判成"未掉电"。改判据前先跑空窗基线, 别再凭感觉给裕量。

    🔴 为什么不能靠探活判重启(2026-07-28 用户台面目击): 源瞬断后立刻恢复出力, 电表 3s 内
    就重启回来, 而设点后本就要 settle 2s 再探活 —— 探活第一拍就通, 重启被整个漏掉, 日志
    干净得像没事发生。run_time 是累计秒表(重启会丢掉上次落 NVM 之后的秒数), 掉一次电就对
    不上墙钟, 这才是能抓住"快到探不着的重启"的客观证据。
    """
    if base is None:
        return None
    # 闭窗同样对齐跳变(两端同口径); 预算放宽到 8s: 闭窗时电表可能刚重启完还在起串口,
    # 撞上就读不通, 2.5s 会直接把证据丢成 None。
    cur = _run_time_edge(vf, timeout_s=8.0)
    if cur is None:
        return None
    return (cur[1] - base[1]) - (cur[0] - base[0])


_SRC_PROBE: "bool | None" = None


def _src_probe_on(config_path: str = CONFIG) -> bool:
    """是否给每个源动作装秒表探针。环境变量 SRC_PROBE=1 优先, 否则读 config run.source_action_probe。

    默认关: 每个动作要等两次秒表跳变(约 +2s), 一条用例约 +10s, 批跑不该常态背这个开销。
    排查掉电时开: `$env:SRC_PROBE=1`(再配 SRC_TRACE=1 可打印帧级时刻)。
    """
    global _SRC_PROBE
    if _SRC_PROBE is None:
        if os.environ.get("SRC_PROBE"):
            _SRC_PROBE = True
        else:
            cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
            _SRC_PROBE = bool(cfg.get("source_action_probe", False))
    return bool(_SRC_PROBE)


def _lost_direct(base: "tuple[float, float] | None",
                 config_path: str) -> "float | None":
    """不经 Verifier 的丢秒计算(串口未被 Verifier 独占时用)。"""
    cur = _run_time_direct_edge(config_path)
    if base is None or cur is None:
        return None
    return (cur[1] - base[1]) - (cur[0] - base[0])


def _probe_action(label: str, fn, vf: "Verifier | None" = None,
                  config_path: str = CONFIG):
    """执行一个会碰源的动作, 并用 run_time 秒表判定这次动作有没有把电表打重启。

    🔴 补的是"未观测区间"(2026-07-28): 原来只有角度切换/读数窗口/退场三处有探针, 而
    **开场(联机+档位模式+钉档+首次保活)、测点下发、救源动作全是盲区** —— 用户台面目击
    批跑"中间掉两下", 日志却报"全程在线", 就是这些盲区造成的。装上后每个源动作都会自报
    丢秒, 掉电不再需要靠猜。关闭时本函数等价于直接调用 fn(), 零开销零行为差异。
    """
    if not _src_probe_on(config_path):
        return fn()
    base = _reboot_probe(vf) if vf is not None else _run_time_direct_edge(config_path)
    out = fn()
    lost = _downtime_since(vf, base) if vf is not None else _lost_direct(base, config_path)
    if lost is None:
        print(f"[probe] {label}: 秒表读不到(电表可能仍在重启/断电)")
    elif lost > REBOOT_LOST_S:
        print(f"[probe] {label}: 🔴 掉电重启 {lost:+.2f}s")
    else:
        print(f"[probe] {label}: 未重启(丢 {lost:+.2f}s)")
    return out


def _zero_current(s: dict) -> dict:
    """测点的零流版: 电压/角度/频率同目标, 三相电流 0(电表仍由 Va 供电, 无载)。"""
    z = dict(s)
    z["ia"] = z["ib"] = z["ic"] = 0.0
    return z


ANGLE_WITNESS_TOL_DEG = 5.0    # 反证容差: 只判"台面是不是我们记的那套角度", 不是精度判据
ANGLE_WITNESS_MIN_V = 20.0     # 该相指令电压低于此值视为未带电, 相角无意义, 不参与比对
ANGLE_WITNESS_SAMPLES = 3      # 每相连采次数(抗表侧相角抖动, 任一样本命中即确认)
_U_ANGLE_REGS = ("PHASE_A_VOLTAGE_PHASE_ANGLE", "PHASE_B_VOLTAGE_PHASE_ANGLE",
                 "PHASE_C_VOLTAGE_PHASE_ANGLE")


def _angle_diff(a: float, b: float) -> float:
    """圆周角差(0~180): 359° 与 1° 差 2°, 不是 358°。"""
    d = abs(a - b) % 360.0
    return min(d, 360.0 - d)


def _angles_confirmed_by_meter(vf: Verifier, src: "SourceUdp", i: int) -> bool:
    """电表反证: 跳发角度帧前, 拿表实测电压相角核对角度缓存是否还代表台面实况。

    🔴 为什么必须反证(2026-07-28): 角度缓存(落盘状态/config assume_angles)记的都是"我们发过
    什么", 源实际在出什么从来没读回来过。有人手动拨了 CL3021 面板, 缓存照样说"跟目标一致"
    → 跳发角度帧 → 整条用例在错角度上跑完, 不报错不告警, 数据静悄悄是错的。
    电表本身就在测相角, 拿它当证人即可闭环: 对得上就跳发(零重启), 对不上就作废缓存强制下发
    (多赔一次重启, 但数据是对的)。

    2026-07-28 实测三档对照(别当理论):
      · 阴性(谎称台面在 0/240/120, 实为 0/1/5) → 抓住, 强制下发 ✅
      · 正常分离角度(0/240/120) → 读数稳定, 确认放行, 该点 0 重启 ✅
      · 近同相角度(006_01_case2 的 B=1°) → 表读数本身在 -0.0/217/288 之间秒级慢抖,
        连采 3 次躲不开 ⇒ **反证会误报**, 后果是白发一次角度帧(多一次重启, 数据仍正确)。
        即这类病态角度用例享受不到跳发优化, 退回改动前水平, 不会更差。

    覆盖边界(免得误以为全覆盖):
      · 只能证电压角 —— 0A 时电流角无意义, "只有电流角被手动改过"仍会漏;
      · 读不通/测量未就绪返回 True 放行 —— 此时反证不可信, 强发只是白赔重启,
        真有问题会被后面的就绪门和判据拦下。
    """
    cached = src.angles_of(src.last_point) if src.last_point else src.assume_angles
    volts = None
    if src.last_point:
        volts = [src.last_point["ua"], src.last_point["ub"], src.last_point["uc"]]
    else:
        st = src.load_state()
        volts = (st or {}).get("volts")
        cached = tuple(st["angles"]) if st else cached
    if not cached or not volts:
        return True                       # 无账可对(会话首点且无落盘) → 交给下面按"要切"处理
    try:
        # 每相连采 ANGLE_WITNESS_SAMPLES 次: 表在近同相输入下相角读数会抖(006_01_case2 的
        # B 相设 1°, 实测在 -0.0/217/288 之间跳), 单次取样会把抖动误判成"台面被人改了"
        # (2026-07-28 实测踩到: 台面明明没动, 反证误报, 白赔一次重启)。
        samples = [[float(vf.read_truth(r)) for _ in range(ANGLE_WITNESS_SAMPLES)]
                   for r in _U_ANGLE_REGS]
    except Exception:
        return True
    if not any(m != 0.0 for grp in samples for m in grp):
        return True                       # 全 0 = 测量未就绪/ADC 异常, 反证不可信
    bad = []
    for ph, (exp, got, v) in enumerate(zip(cached[:3], samples, volts)):
        if float(v) < ANGLE_WITNESS_MIN_V:
            continue                      # 该相没带电, 相角读数无意义
        # 任一样本对得上即确认: 抖动型读数里只要出现过正确值, 台面就是对的;
        # 真错位时 3 个样本都会远离记账值, 照样抓得住。
        if not any(_angle_diff(float(exp), g) <= ANGLE_WITNESS_TOL_DEG for g in got):
            bad.append(f"{'ABC'[ph]}相 记账{exp:g}° vs 实测{[round(g, 1) for g in got]}")
    if bad:
        print(f"[点{i}] ⚠️ 台面角度与记账不符({'; '.join(bad)}) → 角度缓存作废, 强制下发角度帧")
        return False
    return True


def _prearm_angles(src: "SourceUdp", vf: Verifier, s: dict, i: int,
                   gate_s: float = MEASURE_READY_S) -> bool:
    """角度切换前置(2026-07-28, 005/006 频繁掉源整改): 在 0A 上先把角度切过去并等电表复活。

    🔴 成因已实测(2026-07-28): 100V/0A 上纯切角度(电压/频率/档位全程不动)连切 6 次, 台面目击
    **6 次表全部重启**, run_time 每次丢 0.7~1.9s(空窗本底仅 ±0.5s 量化噪声) ⇒ **角度帧一发
    必掉源、电表必重启**, 不是间歇性。005/006 是仅有的"角度≠标准三相"用例组(005 改电流角做
    PF, 006 改电压/电流角), 也正是仅有的频繁掉源组; 002/003/004 全部命中角度缓存跳发角度帧
    → 零重启。旧记录"只有频率变化才掉源、纯角度帧不掉"由此作废。
    ⇒ 每条用例的角度帧数量直接等于电表重启次数, 减帧就是减重启。

    整改前的时序: 目标幅值+新角度一把下发 → 带载瞬断 → 电表重启 → 5s 探活不过 → 重发/冷重臂
    救源(冷重臂的切屏帧又是一次瞬断) → 重启砸进读数窗口 → 哨兵判失效丢弃重读。
    整改后: 先只发"零流版"(角度在无载上切换, 不触发源过载保护) → 安静等电表重启复活+测量就绪
    (期间一帧不发, 不与重启抢) → 再发纯幅值帧加载电流(幅值帧不瞬断) → 读数窗口内无重启。

    返回是否真的做了角度切换(False = 角度命中缓存, 本点无瞬断风险, 调用方按原路径走)。
    """
    if not src.angles_pending(s):
        if _angles_confirmed_by_meter(vf, src, i):
            return False                 # 缓存与台面实况对得上 → 跳发角度帧, 本点零重启
        src.prime_angles(None)           # 缓存脏了 → 作废后按"要切角度"走完整流程
    print(f"[点{i}] 角度/频率与源当前值不同 → 先在 0A 上切换(会瞬断一次, 电表将重启)")
    probe = _reboot_probe(vf)                # 客观重启证据: 快到探不着的重启只有秒表能抓
    src.set_point(_zero_current(s), settle_s=2, force=True)
    try:
        alive = _wait_alive_via(vf, timeout_s=DEAD_DETECT_S)
    except RuntimeError:
        print(f"[点{i}] 角度切换瞬断 → 等电表重启复活(不发任何源帧, 最多 {REARM_BOOT_S}s)")
        try:
            alive = _wait_alive_via(vf, timeout_s=REARM_BOOT_S)
        except RuntimeError:
            print(f"[点{i}] 角度切换后电表未复活 → 交由后续三级救源处理")
            return True
    if alive > 5:
        print(f"[点{i}] 角度切换后电表重启, 等待 {alive:.0f}s 复活")
    if not _wait_measure_ready(vf, timeout_s=gate_s):
        print(f"[点{i}] 角度切换后测量块尚未就绪, 继续加载幅值(就绪门在读数前还会再判一次)")
    lost = _downtime_since(vf, probe)
    if lost is None:
        print(f"[点{i}] 角度切换完成(run_time 探针不可用, 无掉电证据)")
    elif lost > REBOOT_LOST_S:
        print(f"[点{i}] 角度切换实测掉电 {lost:.1f}s(电表已重启一次) → 读数在重启之后取, 不受影响")
    else:
        print(f"[点{i}] 角度切换完成, run_time 连续(丢失 {lost:.1f}s), 本次未掉电")
    return True


def _measure_still_ok(vf: Verifier) -> bool:
    """读后哨兵: 判据寄存器读完后频率仍有效 → 读数窗口内没发生瞬断重启, 数据可采信。"""
    try:
        return float(vf.read_truth("SYSTEM_FREQUENCY")) > 10.0
    except Exception:
        return False


def _reg_label(chk: dict) -> str:
    """报告用寄存器标识: 优先 case_map 的 register_name(重名寄存器按地址锁定时的可读名)。"""
    return str(chk.get("register_name") or chk["register"])


def _chk_expected(chk: dict, n: int = 1) -> str:
    """判据的期望描述: 单区间 [lo,hi] 或多区间并集 ranges(0°回绕类, 2026-07-14 用户定)。

    n>1 时标出连采判决口径(2026-07-27 用户定: 快速连采 n 个数, 全部落区间才算通过)。
    """
    if chk.get("ranges"):
        band = "∪".join(f"[{lo}, {hi}]" for lo, hi in chk["ranges"])
    else:
        band = f"[{chk['lo']}, {chk['hi']}]"
    return f"{band} × 连采{n}次全中" if n > 1 else band


def _chk_ok(chk: dict, val: float) -> bool:
    """判据判定: ranges 命中任一区间即过(0°相角回绕 [0,0.5]∪[359.5,360)); 否则单区间。"""
    if chk.get("ranges"):
        return any(lo <= val <= hi for lo, hi in chk["ranges"])
    return chk["lo"] <= val <= chk["hi"]


def _sample_n(config_path: str) -> int:
    """config run.sample_n: 每条判据的连采次数(默认 SAMPLE_N_DEFAULT, 1=旧单次口径)。"""
    cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
    try:
        return max(1, int(cfg.get("sample_n", SAMPLE_N_DEFAULT)))
    except (TypeError, ValueError):
        return SAMPLE_N_DEFAULT


def _read_samples(vf: "Verifier", register: str, n: int) -> "list[float]":
    """对同一寄存器快速连读 n 次(不加延时, 靠 Modbus 事务本身的节奏), 返回样本列表。"""
    return [float(vf.read_truth(register)) for _ in range(n)]


def _samples_verdict(chk: dict, vals: "list[float]") -> "tuple[bool, str]":
    """连采判决: 全部样本落区间才 PASS; 返回 (是否通过, 报告用 actual 串)。

    actual 串同时给出中位/极值/越界计数——单值看不出抖动, 失败时要能一眼判断是
    "整体偏移"还是"在判据边界抖动"(2026-07-27 用户要求的报告展示口径)。
    """
    bad = [v for v in vals if not _chk_ok(chk, v)]
    if len(vals) == 1:
        return not bad, f"{vals[0]:.4f}"
    ordered = sorted(vals)
    median = ordered[len(ordered) // 2] if len(ordered) % 2 else \
        (ordered[len(ordered) // 2 - 1] + ordered[len(ordered) // 2]) / 2
    return (not bad,
            f"中位{median:.4f} 范围[{ordered[0]:.4f}, {ordered[-1]:.4f}] "
            f"n={len(vals)} 越界{len(bad)}")


def _assert_instant(vf: "Verifier", pt: dict, i: int, report: Report,
                    record_only: bool = False, sample_n: int = 1) -> None:
    """能量类测点的瞬时判据 checks_instant(2026-07-27 补 SRS §3.5 频率精度)。

    energy_accumulate 的 checks 是"窗口首末差值"Δ 判据, 频率这类非累积量套上去恒等于 0 必失败,
    故另立 checks_instant: 在能量基线前连采现值按区间判定, 结构与 source_read 的 checks 一致
    (含连采 sample_n 次全中才 PASS 的口径)。
    """
    for param, chk in (pt.get("checks_instant") or {}).items():
        try:
            vals = _read_samples(vf, chk["register"], sample_n)
        except Exception as exc:                     # 读失败也要留证, 与 Δ 判据同口径
            report.add(Check(
                label=f"点{i} {param}",
                expected=_chk_expected(chk, sample_n),
                actual="(瞬时判据 Modbus 读取失败)",
                passed=False,
                detail=f"register={_reg_label(chk)}; {type(exc).__name__}: {exc}"))
            continue
        ok, shown = _samples_verdict(chk, vals)
        report.add(Check(
            label=f"点{i} {param}",
            expected=_chk_expected(chk, sample_n),
            actual=shown,
            passed=True if record_only else ok,
            detail=f"register={_reg_label(chk)}(瞬时判据, 能量窗起点采样)"
                   + (f"; 采样={[round(v, 4) for v in vals]}" if sample_n > 1 else "")
                   + (f"; 记录模式: 若判决={'PASS' if ok else 'FAIL'}" if record_only else "")))


def load_case(case_id: str) -> dict:
    data = yaml.safe_load(CASE_MAP.read_text(encoding="utf-8"))
    for e in data["cases"]:
        if e["case_id"] == case_id:
            return e
    raise KeyError(f"case_map.yaml 中无用例: {case_id}")


_SRC: "SourceUdp | None" = None


def _source_from_config() -> tuple["SourceUdp", float]:
    """按项目 config.yaml 的 source 段取 CL3021 UDP 控源(会话级单例, 联机/切屏只做一次)。"""
    global _SRC
    cfg = yaml.safe_load(Path(CONFIG).read_text(encoding="utf-8"))
    s = cfg.get("source", {})
    inj = str(s.get("current_injection", "direct"))
    caps = s.get("max_current_a", {}) or {}
    cap = float(caps.get(inj, 0.1))
    gear = bool(s.get("send_gear_frames", False))
    pin = s.get("gear_pin", {}) or {}
    pin_v = float(pin.get("voltage_v", 0.0))
    pin_i = float(pin.get("current_a", 0.0))
    assume = s.get("assume_angles") or None
    state_file = str(PROJECT_ROOT / str(s.get("state_file", ".source_state.json")))
    # 🔴 逐相限幅(2026-07-27 实机: 3E4WY 三相 20A 时源报 "Ic 过载"): 台面 C 相回路只到 15A。
    #    已跑用例的最大值 (20,20,15) 恰在限内, 本护栏只拦超限的新点。
    per_phase = {ph: min(float((s.get("max_current_a_phase", {}) or {}).get(ph, cap)), cap)
                 for ph in ("a", "b", "c")}
    if _SRC is None:
        _SRC = SourceUdp(host=s.get("host", "192.168.0.50"), port=int(s.get("port", 10003)),
                         local_port=int(s.get("local_port", 10005)), max_current_a=cap,
                         send_gear_frames=gear, pin_voltage_v=pin_v, pin_current_a=pin_i,
                         assume_angles=assume, max_current_a_phase=per_phase,
                         state_path=state_file)
        # 落盘状态优先于 config.assume_angles: 前者记的是上个进程真发过的角度(退场不再还原
        # 标准三相后, 台面就停在那个角度上), 后者只是"台面应该在标准三相"的静态假设。
        # 两者都只是记账, 台面被手动改过的情况由 _angles_confirmed_by_meter 反证兜底。
        st = _SRC.load_state()
        if st:
            _SRC.assume_angles = tuple(st["angles"])
            print(f"[src] 角度缓存取自落盘状态 {st['angles']} (updated {st.get('updated')})")
    else:
        _SRC.max_current_a = cap
        _SRC.max_current_a_phase = per_phase
        _SRC.send_gear_frames = gear
        _SRC.pin_voltage_v = pin_v
        _SRC.pin_current_a = pin_i
        _SRC.assume_angles = tuple(assume) if assume else None
        _SRC.state_path = Path(state_file)
    return _SRC, float(s.get("settle_s", 10))


# 电表侧用例前置配置: 接线方式 + 三通道 CT 类型/变比(寄存器值来自地址表与 accuracy_quick 映射)
WIRE_MODE_MAP = {"1E2W": 0, "2E3W1P": 1, "3E4WY": 2}
CT_TYPE_MAP = {"100mA": 0, "333mV": 2}      # AcuRev-100 无 RCT; 80mA 值待与固件确认
CT_PRIMARY_RANGE = (5, 2000)   # 4170/4174/4178 字面值写入, 5-2000A 连续(2026-07-14 真机实证: 写 1000/5
#                                回读一致, v1.05/地址表 v1.01 已落地; 07-09"枚举索引"口径废止)
_CH_BASE = {1: 4167, 2: 4171, 3: 4175}      # CHANNEL_x: +0 wiring_id, +1 direction, +2 ct_type, +3 ct_primary
FREQ_SEL_MAP = {"50Hz": 0, "60Hz": 1}       # FREQUENCY_SELECTION@4161(手工步骤"电表上设置50/60HZ")
FREQ_SEL_ADDR = 4161

# 电流方向(测点级, 004_01_case5 类): spec 枚举 0:positive 1:negative(4168/4172/4176)
DIRECTION_MAP = {"Positive": 0, "Negative": 1}
_DIR_ADDRS = (4168, 4172, 4176)             # CHANNEL_1/2/3_CURRENT_DIRECTION


def _set_direction(vf: Verifier, direction: str) -> None:
    """把三通道电流方向写为 direction(幂等: 读回比对只写差异), 写后留测量恢复时间。"""
    val = DIRECTION_MAP[direction]
    wrote = False
    for addr in _DIR_ADDRS:
        try:
            cur = int(vf.client.read_addr(addr))
        except Exception:
            cur = None
        if cur != val:
            vf.client.write(addr, val)
            wrote = True
    if wrote:
        print(f"[dir] 三通道电流方向 → {direction}({val})")
        time.sleep(2)


SEALING_STATUS_ADDR = 4160      # 0x0A=sealed(铅封锁定) / 0=unsealed; 锁定时测量配置全局拒写(exception 1)


CFG_STABILIZE_S = 360      # 配置写入后恢复总预算: 延迟重启窗口(≤90s)+冷启动(≤240s)+稳定保持
CFG_STABLE_HOLD_S = 15     # 判定恢复: 频率读值连续有效满 15s
CFG_RESTART_WINDOW_S = 90  # 写 4161 后延迟重启最迟出现时间(实测 30~60s, 留裕量)
SRC_REARM_OFFLINE_S = 12   # 失联持续该时长即认定源掉输出, 补拉保活重臂(2026-07-13 台面目击:
#                            重启窗口内源输出掉 0V 且不自恢复, 必须主动重发保活帧才能拉回)


def _apply_meter_cfg(vf: Verifier, src: "SourceUdp", mc: dict, report: Report,
                     adc_triage: bool = False):
    """按 case_map 的 meter_cfg 写接线方式/CT/频率选择(用例前置, 不还原——每条用例自带设定)。

    幂等写: 先读回比对, 只写有差异的寄存器; 全一致则整组跳过(免铅封检查、免测量块扰动)。
    真写了配置后必须吃掉电表的延迟重启窗口(实测写 4161 后约 30~60s 才重启), 且 2026-07-13
    台面目击: 重启窗口内 CL3021 源输出会掉到 0V 且不自恢复 → 失联持续时补拉保活重臂源输出,
    直到测量连续稳定才放行, 否则延迟重启会砸在首测点上(报告 20260713_163600 case2 点1 即此因)。
    """
    if not mc:
        return
    client = vf.client
    wiring = mc.get("wiring")
    _set_wiring_context(wiring)            # 保活点据此决定哪几相该带电(未接线相零输出)
    freq_sel = mc.get("freq_selection")
    ct_type = mc.get("ct_type")
    ct_primary = mc.get("ct_primary")
    desired: "dict[int, tuple[int, bool]]" = {}            # addr -> (目标值, 走range校验)
    if wiring:
        desired[4162] = (WIRE_MODE_MAP[wiring], True)      # SERVICE_CONFIGURATION
    if freq_sel:
        fval = FREQ_SEL_MAP.get(freq_sel)
        if fval is None:
            raise RuntimeError(f"频率选择 '{freq_sel}' 无寄存器映射(仅 50Hz/60Hz)")
        desired[FREQ_SEL_ADDR] = (fval, True)              # FREQUENCY_SELECTION
    if ct_type is not None:
        tval = CT_TYPE_MAP.get(ct_type)
        if tval is None:
            raise RuntimeError(f"CT 类型 '{ct_type}' 无寄存器映射(AcuRev-100 不支持?)")
        pval = int(ct_primary) if ct_primary else None
        if pval is not None and not (CT_PRIMARY_RANGE[0] <= pval <= CT_PRIMARY_RANGE[1]):
            raise RuntimeError(f"CT Primary {ct_primary}A 超出 5-2000A 连续范围")
        for base in _CH_BASE.values():
            desired[base + 2] = (tval, True)               # CHANNEL_x_INPUT_CT_TYPE
            if pval is not None:
                desired[base + 3] = (pval, False)          # CT_PRIMARY 字面值(spec range 未更新, 免 range 校验)
    diffs = {}
    for addr, (val, chk) in desired.items():
        try:
            cur = int(client.read_addr(addr))
        except Exception:
            cur = None                                     # 读不到按有差异处理, 走写入
        if cur != val:
            diffs[addr] = (val, chk)
    if diffs:
        seal = int(client.read_addr(SEALING_STATUS_ADDR))
        if seal != 0:
            raise RuntimeError(
                f"电表铅封锁定(SEALING_STATUS@4160={seal:#x}), 测量配置只读, 写 4161/4162/416x 会被拒(exception 1)。"
                "请拨开电表背面端子侧 Dip Switch 解锁(LCD 应弹 'Remote Configuration Mode')后重跑。")
        for addr, (val, chk) in diffs.items():
            if chk:
                client.write(addr, val)
            else:
                client.write(addr, val, check_range=False)
        time.sleep(1)
    report.add(Check(label="用例前置: 接线/CT/频率 配置",
                     expected=f"wiring={wiring}, ct={ct_type}({ct_primary}A), freq={freq_sel or '不设'}",
                     actual=f"写入 {len(diffs)}/{len(desired)} 项(其余已一致)", passed=True,
                     detail="寄存器 4161/4162/416x; 幂等写(读回比对只写差异)"))
    if diffs and adc_triage:
        # ADC 排查模式: 频率恒0, 按频率等稳定的循环必烧满 CFG_STABILIZE_S → 改定长短等待
        # (写4161有延迟重启风险等30s, 其余配置5s), 换板后 config 关掉 adc_triage 恢复完整逻辑
        wait_s = 30 if FREQ_SEL_ADDR in diffs else 5
        print(f"[cfg] adc_triage: 配置写入后定长等待 {wait_s}s(跳过按频率等稳定)")
        time.sleep(wait_s)
        return
    if diffs:
        # 吃掉延迟重启窗口: 写了 4161(频率选择)时, 即使当下读得到也要等满 CFG_RESTART_WINDOW_S,
        # 确认延迟重启已发生并恢复(或窗口内始终稳定=本次未触发重启)才放行; 其余配置项只要求
        # 连续稳定 CFG_STABLE_HOLD_S。失联持续超 SRC_REARM_OFFLINE_S 判源掉输出, 补拉保活重臂。
        expect_restart = FREQ_SEL_ADDR in diffs
        t0 = time.time()
        restarted = False
        offline_since = None
        stable_since = None
        while time.time() - t0 < CFG_STABILIZE_S:
            try:
                ok = float(vf.read_truth("SYSTEM_FREQUENCY")) > 10.0
            except Exception:
                ok = False
            now_ts = time.time()
            if ok:
                offline_since = None
                if stable_since is None:
                    stable_since = now_ts
                held = now_ts - stable_since >= CFG_STABLE_HOLD_S
                awaiting = (expect_restart and not restarted
                            and now_ts - t0 < CFG_RESTART_WINDOW_S)
                if held and not awaiting:
                    print(f"[cfg] 配置写入后测量已稳定(连续{CFG_STABLE_HOLD_S}s有效), "
                          f"共等待 {now_ts - t0:.0f}s"
                          + ("; 已跨过延迟重启" if restarted else ""))
                    return
            else:
                restarted = True                           # 中断/重启 → 重新计稳定时长
                stable_since = None
                if offline_since is None:
                    offline_since = now_ts
                elif now_ts - offline_since >= SRC_REARM_OFFLINE_S:
                    print("[cfg] 电表失联超时, 源疑似掉输出 → 补拉保活重臂")
                    src.set_point(_keepalive(src), settle_s=3, force=True)
                    offline_since = now_ts                 # 重臂后重新计失联时长
            time.sleep(3)
        print(f"[cfg] ⚠️ 配置写入后 {CFG_STABILIZE_S}s 内测量未达连续稳定, 继续执行并留证")


ENERGY_TRIAGE_WINDOW_S = 60    # ADC排查模式的能量积累窗: 全程读0已坐实, 60s 走通流程留证即可
ENERGY_POLL_CHUNK_S = 10       # 积累窗内分片睡眠粒度(便于日志观测, 不做中途判定)


def run_energy_case(case_meta: dict, config_path: str = CONFIG) -> Report:
    """能量累计用例(007): 设源→记能量基线→积累 duration_s→读末值→Δ区间断言→电流归零。

    与手工步骤"清零→读绝对值"的差异: 自动化用增量断言(E1-E0)等效清零, 不写
    CLEAR_ENERGY@4400——清零类操作按团队规则须人工确认(VERIFIER_BASELINE 红线5),
    增量法零破坏且无需还原。判据区间为 xlsx 原文(T=duration_s 的显示域 kWh/kvarh/kVAh)。
    adc_triage 模式: 积累窗压到 ENERGY_TRIAGE_WINDOW_S 只走流程留证(判据基于全时长,
    短窗读得非零也不判 PASS); 读数全零仍按已知 ADC 口径记 FAIL。
    测点另可带 checks_instant(现值区间判据, 如频率精度): 在记能量基线前采样判定, 见 _assert_instant。
    """
    get_config(config_path)
    _run_cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
    record_only = str(_run_cfg.get("assert_mode", "enforce")).lower() == "record"
    adc_triage = bool(_run_cfg.get("adc_triage", False))
    gate_s = 5.0 if adc_triage else MEASURE_READY_S
    sample_n = _sample_n(config_path)      # 每条判据连采次数(全中才 PASS)
    entry = load_case(case_meta["编号"])
    # 用例开头就登记接线方式: 会话开场的保活点(在 _apply_meter_cfg 之前)也必须遵守
    # "未接线相零输出", 否则 1E2W/2E3W1P 用例开场就会往 B/C 打压 → 源过载锁存。
    _set_wiring_context((entry.get("meter_cfg") or {}).get("wiring"))
    if entry.get("needs_review"):
        raise RuntimeError(
            f"{case_meta['编号']} 的映射仍带 needs_review 标记, 未确认不得执行: "
            f"{entry['needs_review']}")
    duration_s = float(entry.get("duration_s", 600))
    window_s = ENERGY_TRIAGE_WINDOW_S if adc_triage else duration_s
    report = Report(title=f"{case_meta['编号']} {case_meta['标题']}", started=now())
    src, settle_s = _source_from_config()
    try:
        if src.is_inited():
            try:
                boot = _wait_alive(config_path, timeout_s=DEAD_DETECT_S)
            except RuntimeError:
                print("[0] 用例开头电表不在线 → 源冷重臂(切屏重开输出)")
                _probe_action("开场·冷重臂(切屏)", src.reinit_output, config_path=config_path)
                _probe_action("开场·重臂后保活",
                              lambda: src.set_point(_keepalive(src), settle_s=3),
                              config_path=config_path)
                boot = _wait_alive(config_path)
        else:
            _probe_action("开场·warm_init(联机+档位模式+钉档)", src.warm_init,
                          config_path=config_path)
            _probe_action("开场·首次保活", lambda: src.set_point(_keepalive(src), settle_s=3),
                          config_path=config_path)
            try:
                boot = _wait_alive(config_path, timeout_s=WARM_PROBE_S)
            except RuntimeError:
                print("[0] 暖启动探测未通, 转源冷初始化(输出将瞬断, 电表冷重启)")
                _probe_action("开场·冷初始化(联机+切屏+档位模式+钉档)", src.init,
                              config_path=config_path)
                _probe_action("开场·冷初始化后保活",
                              lambda: src.set_point(_keepalive(src), settle_s=3),
                              config_path=config_path)
                boot = _wait_alive(config_path)
            else:
                src.mark_inited()
        print(f"[0] 电表 Modbus 就绪(等待 {boot:.0f}s), 源保持常驻输出")
        with Verifier() as vf:
            _apply_meter_cfg(vf, src, entry.get("meter_cfg"), report, adc_triage=adc_triage)
            for i, pt in enumerate(entry["points"], 1):
                s = _guard_point(pt["source"])
                _prearm_angles(src, vf, s, i, gate_s)   # 角度切换放到 0A 上做(见 _prearm_angles)
                _probe_action(f"点{i}·测点下发", lambda: src.set_point(s, settle_s=settle_s),
                              vf=vf)
                alive = None
                exc_last: "Exception | None" = None
                for rescue in ("probe", "resend", "reinit"):   # 三级救源(同 run_accuracy_case)
                    try:
                        if rescue == "resend":
                            print(f"[点{i}] 设源后 {DEAD_DETECT_S}s 无应答 → 重发本测点救源")
                            _probe_action(f"点{i}·救源重发",
                                          lambda: src.set_point(s, settle_s=3, force=True), vf=vf)
                        elif rescue == "reinit":
                            print(f"[点{i}] 重发无效 → 源冷重臂(切屏重开输出)后重发本测点")
                            _probe_action(f"点{i}·救源冷重臂(切屏)", src.reinit_output, vf=vf)
                            _probe_action(f"点{i}·冷重臂后重发",
                                          lambda: src.set_point(s, settle_s=3, force=True), vf=vf)
                        alive = _wait_alive_via(
                            vf,
                            timeout_s=DEAD_DETECT_S if rescue == "probe" else REARM_BOOT_S)
                        break
                    except RuntimeError as exc:
                        exc_last = exc
                if alive is None:
                    for param, chk in pt["checks"].items():
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=f"Δ∈[{chk['lo']}, {chk['hi']}] ({duration_s:.0f}s)",
                            actual="(设源后电表失联, 重发/冷重臂救源均无效)",
                            passed=False,
                            detail=f"register={_reg_label(chk)}; source={s}; {exc_last}"))
                    for param, chk in (pt.get("checks_instant") or {}).items():
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=_chk_expected(chk, sample_n),
                            actual="(设源后电表失联, 重发/冷重臂救源均无效)",
                            passed=False,
                            detail=f"register={_reg_label(chk)}; source={s}; {exc_last}"))
                    src.set_point(_keepalive(src), settle_s=3, force=True)
                    _wait_alive_via(vf)
                    continue
                if not _wait_measure_ready(vf, timeout_s=gate_s):
                    print(f"[点{i}] 测量未就绪(频率读0) → 能量不会累计, 走流程留证")
                _assert_instant(vf, pt, i, report, record_only, sample_n)
                # 能量基线(增量法起点); 读失败记 None, 末值同败则该判据 FAIL
                base: dict = {}
                for param, chk in pt["checks"].items():
                    try:
                        base[param] = float(vf.read_truth(chk["register"]))
                    except Exception:
                        base[param] = None
                print(f"[点{i}] 能量基线已记录, 积累窗 {window_s:.0f}s"
                      + (f"(adc_triage 短窗, 判据基于 {duration_s:.0f}s)" if window_s != duration_s else ""))
                t0 = time.time()
                while time.time() - t0 < window_s:
                    time.sleep(min(ENERGY_POLL_CHUNK_S, max(0.5, window_s - (time.time() - t0))))
                for param, chk in pt["checks"].items():
                    try:
                        end = float(vf.read_truth(chk["register"]))
                    except Exception as exc:
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=f"Δ∈[{chk['lo']}, {chk['hi']}] ({duration_s:.0f}s)",
                            actual="(能量末值 Modbus 读取失败)",
                            passed=False,
                            detail=f"register={_reg_label(chk)}; {type(exc).__name__}: {exc}"))
                        continue
                    if base[param] is None:
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=f"Δ∈[{chk['lo']}, {chk['hi']}] ({duration_s:.0f}s)",
                            actual=f"(基线读取失败, 末值={end})",
                            passed=False,
                            detail=f"register={_reg_label(chk)}"))
                        continue
                    delta = end - base[param]
                    if window_s != duration_s:
                        # 排查短窗: 判据基于全时长不可判 PASS; Δ=0 按已知 ADC 口径 FAIL 留证
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=f"Δ∈[{chk['lo']}, {chk['hi']}] ({duration_s:.0f}s)",
                            actual=("Δ=0(读数全零: 疑似板子 ADC 芯片损坏或源未出力)" if delta == 0
                                    else f"Δ={delta:.4f}({window_s:.0f}s 排查窗, 不判决仅留证)"),
                            passed=False,
                            detail=f"register={_reg_label(chk)}; base={base[param]}, end={end}; "
                                   f"adc_triage 短窗流程调试, 换板后 config 关 adc_triage 跑全时长"))
                        continue
                    ok = chk["lo"] <= delta <= chk["hi"]
                    report.add(Check(
                        label=f"点{i} {param}",
                        expected=f"Δ∈[{chk['lo']}, {chk['hi']}] ({duration_s:.0f}s)",
                        actual=f"Δ={delta:.4f} (base={base[param]:.4f}, end={end:.4f})",
                        passed=True if record_only else ok,
                        detail=f"register={_reg_label(chk)}" + (
                            f"; 记录模式: 若判决={'PASS' if ok else 'FAIL'}" if record_only else "")))
    finally:
        try:
            _exit_idle(src, config_path)             # 电流归零=能量停累; 保活+探活(2026-07-14)
        finally:
            report.save(f"auto_{case_meta['编号']}")
    return report


def ensure_source_keepalive(config_path: str = CONFIG) -> "SourceUdp":
    """把 CL3021 源电流降到 0(保留 Ua 供电 → 无功率累计), 返回 src。用于"清能量"类"断开输入源"。

    🔴 零重启要点: 源已在出力时**只降电流是无害的**(幅值帧不掉源), 会掉源的是角度帧
    (2026-07-28 实测: 一发必掉、电表必重启, 见 _prearm_angles)与切屏帧。故:
      - 源已在出力(电表在线) → *不做任何 init*, 保活点沿用源当前角度 → set_point 跳发角度帧,
        只发幅值(断流) → 电压不断、电表不重启。
      - 源未出力(电表断电) → 才完整初始化重开输出(冷启动瞬断属必要成本)。
    """
    get_config(config_path)
    src, _ = _source_from_config()
    ka = _keepalive(src)
    ka_angles = (ka["qua"], ka["qub"], ka["quc"], ka["qia"], ka["qib"], ka["qic"], ka["freq"])
    alive = False
    try:
        _wait_alive(config_path, timeout_s=DEAD_DETECT_S)
        alive = True
    except RuntimeError:
        alive = False
    if alive:
        # 源在出力: 预置角度缓存 → set_point 判定角度未变而跳发角度帧, 仅幅值(电流→0), 无瞬断。
        src.prime_angles(ka_angles)
        src.set_point(ka, settle_s=3)
    else:
        # 源没出力: 电表已断电, 必须完整初始化重开输出(会瞬断一次, 冷启动成本)。
        if not src.is_inited():
            src.warm_init()
        src.set_point(ka, settle_s=3)
        _wait_alive(config_path)
    return src


# ── 公开别名(供 wiring_check 等兄弟模块复用会话/自供电机制, 避免跨模块访问保护成员) ──
keepalive_point = _keepalive
guard_point = _guard_point
prearm_angles = _prearm_angles          # 角度切换前置(0A 上切+等复活); 010 接线检查亦可直接复用
zero_current_point = _zero_current
reboot_probe = _reboot_probe            # 重启探针(开窗/闭窗); 判据阈值见 REBOOT_LOST_S
downtime_since = _downtime_since
run_time_direct_edge = _run_time_direct_edge     # 不经 Verifier 的秒表取样(tools/source_diag 用)
angles_confirmed_by_meter = _angles_confirmed_by_meter
wait_alive = _wait_alive
wait_alive_via = _wait_alive_via
wait_measure_ready = _wait_measure_ready
apply_meter_cfg = _apply_meter_cfg
source_from_config = _source_from_config
exit_idle = _exit_idle


def run_accuracy_case(case_meta: dict, config_path: str = CONFIG) -> Report:
    """按 case_map.yaml 逐测点执行: 设源→稳定→Modbus 读→区间断言→源归零。"""
    get_config(config_path)                    # 切换引擎配置单例(报告目录/校验通道均依赖)
    # 断言模式(config run.assert_mode): enforce=区间判决 / record=只记录不判决(2026-07-13 ADC 损坏期间,
    # 用例文件断言不动, 引擎豁免"值在区间"判定但保留流程性失败; 报告记"若判决=PASS/FAIL"留证)
    _run_cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
    record_only = str(_run_cfg.get("assert_mode", "enforce")).lower() == "record"
    # ADC 排查模式(config run.adc_triage): 测量恒0已坐实期间, 就绪门压到5s且不重发测点,
    # 每测点省约50s(2026-07-14 用户反馈批跑过慢)。换板后 config 改回 false 恢复完整保护。
    adc_triage = bool(_run_cfg.get("adc_triage", False))
    gate_s = 5.0 if adc_triage else MEASURE_READY_S
    sample_n = _sample_n(config_path)      # 每条判据连采次数(全中才 PASS)
    entry = load_case(case_meta["编号"])
    # 用例开头就登记接线方式: 会话开场的保活点(在 _apply_meter_cfg 之前)也必须遵守
    # "未接线相零输出", 否则 1E2W/2E3W1P 用例开场就会往 B/C 打压 → 源过载锁存。
    _set_wiring_context((entry.get("meter_cfg") or {}).get("wiring"))
    if entry.get("needs_review"):
        raise RuntimeError(
            f"{case_meta['编号']} 的映射仍带 needs_review 标记, 未确认不得执行: "
            f"{entry['needs_review']}")
    report = Report(title=f"{case_meta['编号']} {case_meta['标题']}", started=now())
    src, settle_s = _source_from_config()
    try:
        # 自供电表 · 最小源交互原则(2026-07-13 频率扫描实证: 一次初始化+220V常驻+只切频率
        # 可连切7频点零断电; 反复联机/切屏/重发保活才是断电来源):
        #   - 本进程已初始化过 → 用例开头一帧源命令都不发(源停在上条用例 finally 的保活上), 只探活电表;
        #   - 进程首用例 → 暖启动: 联机+保活+探活 20s; 不通才冷初始化(切屏瞬断输出, 电表冷重启)。
        if src.is_inited():
            try:
                boot = _wait_alive(config_path, timeout_s=DEAD_DETECT_S)   # 不碰源, 只确认电表在线
            except RuntimeError:
                # 上条用例结束后源输出被关(保护/幅值帧停0) → 冷重臂重开输出(2026-07-14 实证)
                print("[0] 用例开头电表不在线 → 源冷重臂(切屏重开输出)")
                _probe_action("开场·冷重臂(切屏)", src.reinit_output, config_path=config_path)
                _probe_action("开场·重臂后保活",
                              lambda: src.set_point(_keepalive(src), settle_s=3),
                              config_path=config_path)
                boot = _wait_alive(config_path)
        else:
            _probe_action("开场·warm_init(联机+档位模式+钉档)", src.warm_init,
                          config_path=config_path)      # 联机+档位模式(不切屏, 输出不断)
            _probe_action("开场·首次保活", lambda: src.set_point(_keepalive(src), settle_s=3),
                          config_path=config_path)
            try:
                boot = _wait_alive(config_path, timeout_s=WARM_PROBE_S)
            except RuntimeError:
                print("[0] 暖启动探测未通, 转源冷初始化(输出将瞬断, 电表冷重启)")
                _probe_action("开场·冷初始化(联机+切屏+档位模式+钉档)", src.init,
                              config_path=config_path)
                _probe_action("开场·冷初始化后保活",
                              lambda: src.set_point(_keepalive(src), settle_s=3),
                              config_path=config_path)
                boot = _wait_alive(config_path)
            else:
                src.mark_inited()                      # 源已出力且档位模式已置, 免切屏
        print(f"[0] 电表 Modbus 就绪(等待 {boot:.0f}s), 源保持常驻输出")
        with Verifier() as vf:                       # 走 config 的 rtu 校验口
            _apply_meter_cfg(vf, src, entry.get("meter_cfg"), report, adc_triage=adc_triage)
            dir_touched = False                      # 本用例是否动过电流方向(结束必须还原 Positive)
            try:
                for i, pt in enumerate(entry["points"], 1):
                    if pt.get("direction"):          # 测点级电流方向(004_01_case5 类)
                        _set_direction(vf, pt["direction"])
                        dir_touched = True
                    s = _guard_point(pt["source"])   # 🔴 A相供电护栏(最高优先级前提)
                    # 角度/频率有变 → 先在 0A 上切换并等电表重启复活, 再加载幅值(见 _prearm_angles);
                    # 无变化时本调用一帧不发, 直接走原路径。
                    _prearm_angles(src, vf, s, i, gate_s)
                    _probe_action(f"点{i}·测点下发", lambda: src.set_point(s, settle_s=settle_s),
                                  vf=vf)
                    # 救源优先(2026-07-14 用户台面实证: 电表重启很快, 长时间无应答=源输出掉0
                    # 卡死——源不出电电表永远起不来, 傻等电表只会让源死更久):
                    # 12s 无应答 → 重发本测点救源 → 仍无 → 切屏冷重臂+重发本测点 → 全无效才判
                    # "源掉输出"记 FAIL 退保活。救活即继续正常测本点, 不白记失联。
                    alive = None
                    exc_last: "Exception | None" = None
                    for rescue in ("probe", "resend", "reinit"):
                        try:
                            if rescue == "resend":
                                print(f"[点{i}] 设源后 {DEAD_DETECT_S}s 无应答 → 重发本测点救源")
                                _probe_action(f"点{i}·救源重发",
                                              lambda: src.set_point(s, settle_s=3, force=True),
                                              vf=vf)
                            elif rescue == "reinit":
                                print(f"[点{i}] 重发无效 → 源冷重臂(切屏重开输出)后重发本测点")
                                _probe_action(f"点{i}·救源冷重臂(切屏)", src.reinit_output, vf=vf)
                                _probe_action(f"点{i}·冷重臂后重发",
                                              lambda: src.set_point(s, settle_s=3, force=True),
                                              vf=vf)
                            alive = _wait_alive_via(
                                vf,
                                timeout_s=DEAD_DETECT_S if rescue == "probe" else REARM_BOOT_S)
                            break
                        except RuntimeError as exc:
                            exc_last = exc
                    if alive is None:
                        # 三级救源均无效 → 该点全部判据记 FAIL 留证, 退保活后继续后续测点
                        for param, chk in pt["checks"].items():
                            report.add(Check(
                                label=f"点{i} {param}",
                                expected=_chk_expected(chk, sample_n),
                                actual="(设源后电表失联, 重发/冷重臂救源均无效)",
                                passed=False,
                                detail=f"register={_reg_label(chk)}; source={s}; {exc_last}"))
                        if not pt["checks"]:         # 无判据的测点也留一条失联记录
                            report.add(Check(label=f"点{i} (无判据)", expected="设源成功",
                                             actual="(设源后电表失联, 重发/冷重臂救源均无效)",
                                             passed=False, detail=f"source={s}; {exc_last}"))
                        src.set_point(_keepalive(src), settle_s=3, force=True)
                        try:
                            _wait_alive_via(vf)
                        except RuntimeError:
                            raise RuntimeError(
                                f"点{i} 设源后电表失联, 本测点重臂/冷重臂/保活均无效——源输出"
                                "疑似被保护性关闭且远程恢复失败, 需人工到台面开启源输出后重跑"
                            ) from exc_last
                        continue
                    if alive > 5:
                        print(f"[点{i}] 源切换后电表重启, 等待 {alive:.0f}s 复活")
                    # 测量就绪门(2026-07-14): Modbus 通≠测量可用。频率读不到有效值 → 先重发本测点
                    # 一次(档位切换后幅值帧疑被吞的兜底), 仍不就绪 → 全判据 FAIL 留证, 不读零凑数。
                    # adc_triage 模式: 门压到5s且跳过重发(源侧正常已确认, 重发是纯浪费)
                    if not _wait_measure_ready(vf, timeout_s=gate_s):
                        if not adc_triage:
                            print(f"[点{i}] 测量未就绪(频率读0/读不通) → 重发本测点一次")
                            try:
                                _probe_action(f"点{i}·就绪门重发",
                                              lambda: src.set_point(s, settle_s=settle_s,
                                                                    force=True), vf=vf)
                                _wait_alive_via(vf, timeout_s=DEAD_DETECT_S)
                            except Exception as exc:
                                print(f"[点{i}] 重发后电表仍失联: {exc}")
                        if adc_triage or not _wait_measure_ready(vf, timeout_s=gate_s):
                            for param, chk in pt["checks"].items():
                                report.add(Check(
                                    label=f"点{i} {param}",
                                    expected=_chk_expected(chk, sample_n),
                                    actual="(读数全零: 疑似板子 ADC 芯片损坏或源未出力)",
                                    passed=False,
                                    detail=f"register={_reg_label(chk)}; source={s}; "
                                           f"读数为0类失败=已知 ADC 问题口径(2026-07-14 用户指示), "
                                           f"流程继续不中断"
                                           + ("; adc_triage模式(免重发)" if adc_triage else "; 重发一次后频率仍读0")))
                            continue
                    # 读判据: 先整组读值→读后哨兵验证读数窗口内无瞬断重启→才入报告;
                    # 哨兵失效则丢弃整组重读一次(2026-07-14 用户要求: 读数与板子重启解耦,
                    # 绝不采信"半程重启"的混合数据)
                    vals: dict = {}
                    errs: dict = {}
                    for attempt in (1, 2):
                        vals, errs = {}, {}
                        probe = _reboot_probe(vf)      # 读窗开窗: 秒表基准
                        for param, chk in pt["checks"].items():
                            try:
                                vals[param] = _read_samples(vf, chk["register"], sample_n)
                            except Exception as exc:
                                errs[param] = exc
                        # 双哨兵: 频率仍有效(测量块没死) + run_time 跟得上墙钟(窗口内没掉过电)。
                        # 只看频率会漏掉 3s 级重启——恢复后频率立刻正常, 但样本已是重启后爬坡值。
                        lost = _downtime_since(vf, probe)
                        rebooted = lost is not None and lost > REBOOT_LOST_S
                        if _measure_still_ok(vf) and not rebooted:
                            break
                        if attempt == 1:
                            why = (f"窗口内掉电 {lost:.1f}s" if rebooted else "测量失效(频率读0/读不通)")
                            print(f"[点{i}] 读数窗口内{why}, 丢弃本组读数重读")
                            try:
                                _wait_alive_via(vf, timeout_s=DEAD_DETECT_S)
                            except RuntimeError:
                                pass
                            if not _wait_measure_ready(vf):
                                break              # 测量拉不回 → 按已读到的值留证(零值判 FAIL)
                    if not pt["checks"]:
                        # 手工用例明写"本测点无精度要求"(如 004_01_case2/case11 的第二点):
                        # 仍记一条流程性 PASS, 否则报告里连"这个点跑过"都看不出来。
                        report.add(Check(
                            label=f"点{i} (无精度要求, 仅走流程)",
                            expected="(手工用例注明本测点不断言)",
                            actual="已下发测点且电表在线",
                            passed=True,
                            detail=f"source={s}"))
                    for param, chk in pt["checks"].items():
                        if param in errs:          # 读失败记 FAIL, 报告完整落盘
                            exc = errs[param]
                            report.add(Check(
                                label=f"点{i} {param}",
                                expected=_chk_expected(chk, sample_n),
                                actual="(Modbus 读取失败)",
                                passed=False,
                                detail=f"register={_reg_label(chk)}; {type(exc).__name__}: {exc}"))
                            continue
                        ok, shown = _samples_verdict(chk, vals[param])
                        report.add(Check(
                            label=f"点{i} {param}",
                            expected=_chk_expected(chk, sample_n),
                            actual=shown,
                            passed=True if record_only else ok,
                            detail=f"register={_reg_label(chk)}"
                                   + (f"; 采样={[round(v, 4) for v in vals[param]]}" if sample_n > 1 else "")
                                   + (f"; 记录模式(判决暂缓): 若判决={'PASS' if ok else 'FAIL'}" if record_only else "")))
                    # TODO(真机/方案B): 此处追加 GUI OCR 比对(需 Real-Time 表格行坐标 + Tesseract)
            finally:
                # 🔴 方向还原(2026-07-14 用户指示): 动过电流方向的用例, 结束必须还原 Positive,
                # 失败重试一次, 仍失败 → Check FAIL(整例判失败)并需在汇总里置顶告警
                if dir_touched:
                    readback = None
                    ok = False
                    for _attempt in (1, 2):
                        try:
                            _set_direction(vf, "Positive")
                            readback = [int(vf.client.read_addr(a)) for a in _DIR_ADDRS]
                            ok = all(v == DIRECTION_MAP["Positive"] for v in readback)
                        except Exception as exc:
                            readback = f"还原异常: {exc}"
                            ok = False
                        if ok:
                            break
                    report.add(Check(label="用例还原: 电流方向=Positive",
                                     expected="4168/4172/4176 = 0(positive)",
                                     actual=str(readback), passed=ok,
                                     detail="2026-07-14 用户指示: 方向类用例结束必须还原 Positive; "
                                            "还原失败属设备遗留非默认状态, 需人工恢复"))
    finally:
        # 退场: 统一恢复保活(Ua=100V)+探活校验(2026-07-14 用户指示, 见 _exit_idle)。
        # 报告落盘放最内层 finally: 无论用例中途怎么失败, 证据必须留下(2026-07-13 监督整改)
        try:
            _exit_idle(src, config_path)
        finally:
            report.save(f"auto_{case_meta['编号']}")
    return report
