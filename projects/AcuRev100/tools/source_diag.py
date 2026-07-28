"""CL3021 控源动作"输出掉0"排查工具（手动执行, 人眼观察源输出）。

背景(2026-07-14): wiring_check 批跑期间工程师目击源输出多次掉 0, 但脚本侧
"电表失联"探测无感知——电表仅从 Ua/Vn 取电, Ub/Uc 掉 0 探测不到; Ua 短暂瞬断
在 8s 探测窗内恢复也探测不到; ADC 损坏期无法读 RMS 交叉验证。
本工具把控源拆成单一原语动作, 每个函数只重复执行一种动作, 由工程师盯着源
输出面板逐动作排查: 哪个动作一执行输出就掉 0, 即为元凶。

用法(仓库根目录, 手动交互):
    python projects/AcuRev100/tools/source_diag.py
    # 可选: 先 chcp 65001 避免控制台中文乱码
菜单选动作编号 + 重复次数; 每次迭代打印时间戳、执行动作、并给出**重启判定**。
q 退出(自动恢复保活)。

判定口径(2026-07-28 升级): 原来每次迭代只做一次 Modbus 探活, 打印"电表在线"——这个判据
是瞎的: 电表 1 秒级就重启回来, 探活必然显示在线, 目击到的重启全被漏掉。现改用
DEVICE_RUN_TIME(0x1019) 秒表: 动作前后各取一次"跳变对齐"样本, 墙钟流逝 − 秒表增量 >0.6s
即掉电重启(空窗本底 ±0.00s, 一次重启丢 0.97~1.9s, 已实测标定)。
⚠️ 仍以人眼为准的部分: 秒表只能抓"表真重启"; 源瞬掉而表靠电容扛过去时秒表是连续的,
   以及 Ub/Uc 掉 0(电表只从 Ua 取电)——这两类只有盯源输出面板才看得见。

2026-07-28 已实测确认的因果(菜单里对应动作可复验):
  · 角度帧(动作5): **一发必掉源, 电表必重启** —— 6/6 目击 + 秒表每次丢 0.7~1.9s;
  · 同值频率帧(动作6): 不掉源(002/004 每点发两轮从不掉);
  · 幅值帧(动作2/3/16/17): 不掉源, 含 0→20A 带载加载、100↔270V 跳变;
  · 跨档切换(动作18/19): **未验证** —— 旧记录说"档位切换不掉源"但同一批记录里"角度帧不掉源"
    已被证伪, 故不再采信, 这正是本次要测的。

安全: 电流固定 ≤5A(远低于 config via_ct 25A 上限); 档位沿用 config gear_pin
钉死(480V/20A); 全程 Ua≥100V 保供电(动作本身的瞬断除外——那正是要观察的对象)。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]           # projects/AcuRev100
REPO_ROOT = PROJECT_ROOT.parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# 控制台编码兜底(2026-07-27 修): Windows 控制台默认 GBK, 编不了 U+2194 双向箭头、U+26A0
# 警告号、U+274C 叉号 → 打印菜单时抛 UnicodeEncodeError, 表现为"脚本启动不了"
# (源其实已初始化完, 只是菜单没打出来)。菜单文案已改用 ASCII 记号, 这里再兜一层:
# errors=replace 保留控制台原编码(中文正常显示), 编不了的字符退化成 ?, 永不因打印崩溃。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

from projects.AcuRev100.tests import helpers_accuracy as ha  # noqa: E402
from projects.AcuRev100.tools.accuracy_quick.core import source_comm as _sc  # noqa: E402

CONFIG = str(PROJECT_ROOT / "config.yaml")

# 基准点(与 wiring_check 批跑的 3E4WY-ABC 基准一致)
BASE = dict(ua=220.0, ub=220.0, uc=220.0, qua=0.0, qub=240.0, quc=120.0,
            ia=5.0, ib=5.0, ic=5.0, qia=0.0, qib=240.0, qic=120.0, freq=50.0)
FAULT_UB5 = dict(BASE, ub=5.0)          # wiring 批跑典型故障点(Vb 缺失)
KA_AMP = dict(BASE, ua=100.0, ia=0.0, ib=0.0, ic=0.0)   # 保活点幅值(角度与 BASE 相同)


class DiagSource(ha.SourceUdp):
    """暴露单帧原语(供逐动作排查); 复用父类的 ACK 校验/重试/档位钉死。"""

    def frame_angle(self, s: dict) -> None:
        self._cmd("角度帧", _sc._build_angle_update(s["quc"], s["qub"], s["qua"],
                                                    s["qic"], s["qib"], s["qia"], s["freq"]))

    def frame_amplitude(self, s: dict) -> None:
        self._cmd("幅值帧", _sc._build_amplitude_update(s["uc"], s["ub"], s["ua"],
                                                        s["ic"], s["ib"], s["ia"]))

    def frame_freq(self, freq: float) -> None:
        self._cmd("频率帧", _sc._build_freq_update(freq))

    def frame_voltage_gear(self, v: float) -> None:
        self._cmd("电压档位帧", _sc._build_voltage_gear(v, v, v))

    # ⛔ 不要再加"逐相档位帧"原语(2026-07-28 事故): 原有代码只发过三相同档, 帧里
    #    gear(uc)/gear(ub)/gear(ua) 的字节顺序与设备实际相位的对应关系**从未验证**;
    #    直接拿来做逐相跨档, 结果把 A 相钉进了 ≤240V 档, 随后回基准点灌 220V(92% 满刻度)
    #    → 源报 Ua+Ub 过载并锁存, 表断电、程序拉不回, 只能人工到面板复位。
    #    跨档一律用下面的三相同档形式(已验证), 并且改过档位后必须先恢复钉死档再回基准点。

    def frame_current_gear(self, i: float) -> None:
        self._cmd("电流档位帧", _sc._build_current_gear(i, i, i))

    def frame_gear_mode(self) -> None:
        self._cmd("档位模式帧", _sc._build_gear_mode("00000000"))

    def frame_connect(self) -> None:
        self._cmd("联机帧", _sc._build_connect())

    def frame_switch_screen(self) -> None:
        self._cmd("切AC屏帧", _sc._build_switch_screen_cmd(0x01))


GEAR_CROSS_BASE = dict(BASE, ua=100.0, ub=100.0, uc=100.0, ia=0.0, ib=0.0, ic=0.0)
#   ↑ 跨档测试的安全基点: 三相 100V + 0A。100V 在 <=240V 档占 42%、在 <=480V 档占 21%,
#     两档都远离满刻度; 且 A 相有压 → 表全程在线, run_time 秒表可用。


def _gear_cross_v(src: "DiagSource", i: int) -> None:
    """电压档位跨档轮切(三相同档 <=240V <-> <=480V, 即档2<->档1), 输出恒 100V/0A。

    只在这两档之间切: 100V 在两档里分别是 42%/21% 量程, 不会压满刻度; 更小的档(<=120V)
    对 100V 就是 83% 量程 —— 那正是已记录的 "Ub/Uc 过载" 诱因, 不碰。
    """
    if i == 1:
        src.frame_amplitude(GEAR_CROSS_BASE)
        time.sleep(1.0)
    src.frame_voltage_gear(240.0 if i % 2 else 480.0)


def _gear_cross_i(src: "DiagSource", i: int) -> None:
    """电流档位跨档轮切(输出恒 0A, 无过载风险); 第1次迭代先降到安全基点。"""
    if i == 1:
        src.frame_amplitude(GEAR_CROSS_BASE)
        time.sleep(1.0)
    src.frame_current_gear((20.0, 5.0, 0.5)[i % 3])   # 档2 / 档4 / 档7


def _restore_pinned_gear(src: "DiagSource") -> None:
    """把档位恢复成会话钉死值(config source.gear_pin), 供动作组收尾用。

    改过档位的动作组(18/19, 或手动发过 7/8 的小档)结束后, 若直接回 220V/5A 的基准点, 就是
    把大幅值灌进小量程 → 源过载锁存(2026-07-28 事故)。恢复必须"先放大量程, 再升幅值"。
    """
    if src.pin_voltage_v:
        src.frame_voltage_gear(src.pin_voltage_v)
        time.sleep(1.0)
    if src.pin_current_a:
        src.frame_current_gear(src.pin_current_a)
        time.sleep(0.5)


def _probe_meter() -> str:
    """单次 Modbus 探活(1s): 仅能反映 Ua 供电; Ub/Uc 掉 0 探测不到, 以人眼为准。"""
    import yaml
    from pymodbus.client import ModbusSerialClient
    rtu = yaml.safe_load(Path(CONFIG).read_text(encoding="utf-8"))["transport"]["rtu"]
    c = ModbusSerialClient(port=rtu["port"], baudrate=rtu["baudrate"], parity=rtu["parity"],
                           bytesize=rtu["bytesize"], stopbits=rtu["stopbits"], timeout=1)
    try:
        if c.connect():
            r = c.read_holding_registers(4111, count=1, device_id=rtu["slave_id"])
            if r and not r.isError():
                return "电表在线"
        return "[!]电表失联(Ua疑似掉0)"
    except Exception:
        return "[!]电表失联(Ua疑似掉0)"
    finally:
        c.close()


def _verdict(base) -> str:
    """动作后的重启判定: run_time 丢秒 > 阈值即掉电重启(口径见模块 docstring)。"""
    cur = ha.run_time_direct_edge(CONFIG)
    if base is None or cur is None:
        return f"[!]秒表读不到({_probe_meter()})"
    lost = (cur[1] - base[1]) - (cur[0] - base[0])
    if lost > ha.REBOOT_LOST_S:
        return f"[!!]掉电重启 (run_time 丢 {lost:+.2f}s)"
    return f"未重启 (run_time 丢 {lost:+.2f}s)"


def _run(name: str, n: int, interval_s: float, action) -> None:
    """重复执行单一动作 n 次: 秒表开窗→执行→秒表闭窗判重启, 工程师同步盯源输出。"""
    input(f"\n即将执行【{name}】×{n}(间隔{interval_s}s)。请盯住源输出面板, 回车开始...")
    hits = 0
    for i in range(1, n + 1):
        base = ha.run_time_direct_edge(CONFIG)      # 开窗(对齐秒表跳变)
        ts = time.strftime("%H:%M:%S")
        try:
            action(i)
            note = _verdict(base)
        except Exception as exc:
            note = f"[X]动作异常: {exc}"
        hits += "[!!]" in note
        print(f"  [{ts}] 第{i}/{n}次 → {note}")
        time.sleep(interval_s)
    print(f"【{name}】完成: {n} 次里秒表判定掉电重启 {hits} 次。")
    print("  ⚠️ 秒表只抓'表真重启'; 源瞬掉但表扛过去了、或 Ub/Uc 掉0, 只有你盯面板看得见 —— "
          "肉眼次数与上面对不上时以肉眼为准, 并把差异告诉我。")


def main() -> None:
    import yaml
    s_cfg = yaml.safe_load(Path(CONFIG).read_text(encoding="utf-8")).get("source", {}) or {}
    pin = s_cfg.get("gear_pin", {}) or {}
    src = DiagSource(host=s_cfg.get("host", "192.168.0.50"), port=int(s_cfg.get("port", 10003)),
                     local_port=int(s_cfg.get("local_port", 10005)), max_current_a=25.0,
                     pin_voltage_v=float(pin.get("voltage_v", 0.0)),
                     pin_current_a=float(pin.get("current_a", 0.0)))
    print("初始化: 暖启动(联机+档位模式+钉档) → 输出基准点(与批跑相同的会话开场)")
    src.warm_init()
    src.set_point(dict(BASE), settle_s=2)
    print(f"基准已输出: Ua/Ub/Uc=220V, Ia/Ib/Ic=5A。探活: {_probe_meter()}")

    actions = {
        "1": ("保活重发(完整 set_point KEEPALIVE, 角度+幅值+频率×2轮)", 3.0,
              lambda i: src.set_point(ha.keepalive_point(), settle_s=0)),
        "2": ("同值幅值帧(220V/5A 原值重发, 不该有任何扰动)", 3.0,
              lambda i: src.frame_amplitude(BASE)),
        "3": ("幅值切换(全相 220V <-> 120V, 同 480V 钉死档)", 3.0,
              lambda i: src.frame_amplitude(dict(BASE, ua=120.0, ub=120.0, uc=120.0)
                                            if i % 2 else BASE)),
        "4": ("单相低压幅值切换(ub 220V <-> 5V, wiring 故障注入同款)", 3.0,
              lambda i: src.frame_amplitude(FAULT_UB5 if i % 2 else BASE)),
        "5": ("角度帧切换(qub 240° <-> 200°)——2026-07-28 已确认: 一发必掉源+表必重启", 3.0,
              lambda i: src.frame_angle(dict(BASE, qub=200.0) if i % 2 else BASE)),
        "6": ("同值频率帧(50Hz 原值重发)——已确认不掉源", 3.0,
              lambda i: src.frame_freq(50.0)),
        "7": ("电压档位帧(480V 同档重发, 不跨档)", 4.0,
              lambda i: src.frame_voltage_gear(480.0)),
        "8": ("电流档位帧(20A 同档重发, 不跨档)", 4.0,
              lambda i: src.frame_current_gear(20.0)),
        "9": ("档位模式帧(warm_init 每会话发一次的那帧)", 4.0,
              lambda i: src.frame_gear_mode()),
        "10": ("联机帧(每条用例开场可能重发)", 3.0,
               lambda i: src.frame_connect()),
        "11": ("切AC屏帧(冷重臂用——历史结论: 必瞬断且重开输出)", 5.0,
               lambda i: src.frame_switch_screen()),
        "12": ("完整测点下发(set_point 基准, 角度+幅值+频率×2轮, 批跑每点动作)", 4.0,
               lambda i: src.set_point(dict(BASE), settle_s=0)),
        "13": ("wiring 批跑序列复现(基准→ub=5V 故障→基准, set_point×3)", 4.0,
               lambda i: [src.set_point(dict(BASE), settle_s=1),
                          src.set_point(dict(FAULT_UB5), settle_s=1),
                          src.set_point(dict(BASE), settle_s=1)]),
        # ── 14~17: 第二元凶隔离(2026-07-14 动作1同值保活仍掉0, 疑在 0A 电流或低压幅值帧) ──
        "14": ("同值幅值帧·保活值(ua=100V/ub=uc=220V/0A 重发, 首次含一次转换)", 3.0,
               lambda i: src.frame_amplitude(KA_AMP)),
        "15": ("幅值帧交替 220V/5A  <->  保活(100V/0A)(完整保活转换, 无角度帧)", 3.0,
               lambda i: src.frame_amplitude(KA_AMP if i % 2 else BASE)),
        "16": ("幅值帧交替 220V/5A  <->  100V/5A(只降压, 电流保持——隔离低压因素)", 3.0,
               lambda i: src.frame_amplitude(dict(BASE, ua=100.0) if i % 2 else BASE)),
        "17": ("幅值帧交替 220V/5A  <->  220V/0A(只断流, 电压保持——隔离0A因素)", 3.0,
               lambda i: src.frame_amplitude(dict(BASE, ia=0.0, ib=0.0, ic=0.0)
                                             if i % 2 else BASE)),
        # ── 18~19: 跨档切换(2026-07-28 新增) —— 决定 pytest 路径能否关掉钉档改逐点档位 ──
        #    安全设计: 先把输出降到 0A + B/C 30V(动作内第1次迭代自动做), A 相恒 100V 保供电;
        #    跨的档全部 ≥ 实际输出幅值且留足余量, 不会压在满刻度上(那是已记录的 Ub/Uc 过载诱因)。
        "18": ("电压档位跨档(B/C 在 <=60V/<=120V/<=240V 三档间轮切, A相档位不动)", 4.0,
               lambda i: _gear_cross_v(src, i)),
        "19": ("电流档位跨档(0A 输出上, 在 20A/5A/0.5A 三档间轮切)", 4.0,
               lambda i: _gear_cross_i(src, i)),
    }
    while True:
        print("\n──── 控源动作排查菜单(盯源输出, 哪个动作掉0哪个是元凶) ────")
        for k, (name, _iv, _fn) in actions.items():
            print(f"  {k:>2}. {name}")
        print("   q. 退出(恢复保活)")
        choice = input("选动作编号: ").strip().lower()
        if choice == "q":
            break
        if choice not in actions:
            print("无效编号")
            continue
        name, interval, fn = actions[choice]
        try:
            n = int(input("重复次数(默认 10): ").strip() or "10")
        except ValueError:
            n = 10
        _run(name, n, interval, fn)
        # 🔴 必须先恢复钉死档再回基准点(2026-07-28 事故整改): 基准点是 220V/5A, 而跨档动作
        #    会把档位留在小量程上, 直接回基准点 = 220V 灌进 <=60V/<=120V 档 → 源报过载并
        #    锁存, 表断电、程序拉不回, 只能人工到面板复位。恢复顺序: 先放大量程, 再升幅值。
        print("动作组结束 → 先恢复钉死档位(480V/20A) → 再回基准点")
        _restore_pinned_gear(src)
        src.set_point(dict(BASE), settle_s=2)
    print("退出: 恢复保活(Ua=100V, 电流0)")
    src.set_point(ha.keepalive_point(), settle_s=2)
    print(f"最终探活: {_probe_meter()}")
    src.close()


if __name__ == "__main__":
    main()
