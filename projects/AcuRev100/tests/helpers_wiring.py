"""AcuRev-100 接线检查判据用例运行时（wiring_check 模块, 010_07~010_13）。

职责: 读 case_map_wiring.yaml 中某条用例的 points → 逐点 CL3021 设源(故障注入) →
关→开 WIRE_CHECK_SWITCH 触发立即检测 → Modbus 读错误码/相序 → 位级精确断言 → Report。

与 helpers_accuracy 的关系(完全复用其控源与自供电机制, 不重复实现):
- 控源会话/档位钉死/三级救源/A相护栏/保活退场: 直接 import 使用;
- 接线方式前置(meter_cfg → SERVICE_CONFIGURATION@4162): 复用 _apply_meter_cfg;
- 新增三件事: ①相序配置前置(PHASE_ORDER@4195, 铅封预检) ②开关重开触发立即检测
  (WIRE_CHECK_SWITCH@12288, 0=ON/1=OFF, 用户裁定: 开关开启瞬间检查一次, 常开 1 次/min)
  ③错误码位断言(12290/12291 精确值比对, 短路逻辑=被跳过的位必须为 0)。

断言口径(2026-07-14 工程师确认):
- v_code/i_code 按期望精确值比对(uint16 全字段), 未置位=不告警一并验证;
- 电压缺失类用例的 i_code 只记录不断言(该相电压≈0 时 pf 无定义, 固件电流侧行为未约定);
- voltagePhaseOrder=2 不参与条件13 比较(bit8 不置位, 用户裁定 2026-07-14)。
"""
from __future__ import annotations

import sys
import time
from collections.abc import Callable
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[1]          # projects/AcuRev100 (本文件在 tests/ 根)
REPO_ROOT = PROJECT_ROOT.parents[1]
if str(REPO_ROOT) not in sys.path:                           # 直跑本文件时兜底; pytest 由根 conftest 注入
    sys.path.insert(0, str(REPO_ROOT))

from comm.ctl_acuview.config import get_config  # noqa: E402
from comm.ctl_acuview.verify import Check, Report, Verifier, now  # noqa: E402
from projects.AcuRev100.tests import helpers_accuracy as ha  # noqa: E402

CASE_MAP = Path(__file__).resolve().parent / "case_map_wiring.yaml"
CONFIG = str(PROJECT_ROOT / "config.yaml")

# ── 接线检查寄存器(spec: RACG Modbus Address Table / Wiring Check sheet) ──
WIRE_CHECK_SWITCH = 12288    # 0x3000 uint16 R/W  0:ON 1:OFF(注意语义反直觉)
WIRE_CHECK_STATUS = 12289    # 0x3001 uint16 R    0:Not Started 1:In Progress 2:Completed
V_ERROR_CODE = 12290         # 0x3002 uint16 R    电压错误码(位定义见 V_BITS)
I_ERROR_CODE = 12291         # 0x3003 uint16 R    电流错误码(位定义见 I_BITS)
VOLTAGE_PHASE_ORDER = 12304  # 0x3010 uint32 R    0:ABC 1:ACB 2:未判定
PHASE_ORDER_CFG = 4195       # 0x1063 uint16 R/W  ABC=0/ACB=1(测量配置, 铅封锁定拒写)

V_BITS = {0: "Va缺失", 1: "Vb缺失", 2: "Vc缺失", 3: "Va-Vn反接", 4: "Vb-Vn反接",
          5: "Vc-Vn反接", 6: "Vb相位错误", 7: "Vc相位错误", 8: "相序配置错误"}
I_BITS = {0: "Ia缺失", 1: "Ib缺失", 2: "Ic缺失", 3: "Ia反向", 4: "Ib反向",
          5: "Ic反向", 6: "Ia-Ib互错", 7: "Ia-Ic互错", 8: "Ib-Ic互错"}

TRIGGER_COMPLETED_S = 20     # 开关重开后等 WIRE_CHECK_STATUS=Completed 的窗口
TRIGGER_FALLBACK_S = 70      # 开关拒写且电表确在线时退化等周期检测(1 次/min, 留 10s 裕量)
CODE_SETTLE_S = 1.5          # status=Completed 后读码前留隙(结果寄存器刷新)
RIDE_THROUGH_CONFIRM_S = 2.0  # 探活复确认间隔: 角度帧关死输出后电表电容余电可撑数秒,
                              # 首探可能假活(2026-07-16 台面实证: 有源电表 3s 内启动,
                              # ≥30s 离线空窗一律是源输出关死而非电表开机慢)


def _fmt(val: int, names: dict) -> str:
    """错误码可读化: 0x0008 → '0x0008[Va-Vn反接]'; 0 → '0x0000[无告警]'。"""
    hits = [n for b, n in names.items() if val >> b & 1]
    return f"{val:#06x}[{'|'.join(hits) if hits else '无告警'}]"


def load_case(case_id: str) -> dict:
    data = yaml.safe_load(CASE_MAP.read_text(encoding="utf-8"))
    for e in data["cases"]:
        if e["case_id"] == case_id:
            return e
    raise KeyError(f"case_map_wiring.yaml 中无用例: {case_id}")


def _apply_phase_order(vf: Verifier, po: "int | None", report: Report) -> None:
    """相序配置前置(仅 3E4WY 用例携带): 幂等写 PHASE_ORDER@4195, 写前铅封预检。"""
    if po is None:
        return
    try:
        cur = int(vf.client.read_addr(PHASE_ORDER_CFG))
    except Exception:
        cur = None
    if cur != po:
        seal = int(vf.client.read_addr(ha.SEALING_STATUS_ADDR))
        if seal != 0:
            raise RuntimeError(
                f"电表铅封锁定(SEALING_STATUS@4160={seal:#x}), PHASE_ORDER@4195 拒写。"
                "请拨开电表背面 Dip Switch 解锁后重跑。")
        vf.client.write(PHASE_ORDER_CFG, po)
        time.sleep(2)
    report.add(Check(label="用例前置: 相序配置",
                     expected=f"PHASE_ORDER@4195={po}({'ABC' if po == 0 else 'ACB'})",
                     actual="已一致(免写)" if cur == po else f"已写入(原值 {cur})",
                     passed=True, detail="寄存器 0x1063; 幂等写"))


def _trigger_check(vf: Verifier, completed_s: float = TRIGGER_COMPLETED_S,
                   recover: "Callable[[], None] | None" = None) -> str:
    """关→开接线检查开关触发立即检测(用户裁定机制); 开关拒写按离线/在线分路处理。

    completed_s: 等 status=Completed 的窗口。ADC 损坏期检查任务不执行、status 恒 0,
    每点等满 20s 纯浪费(2026-07-14 批跑 333s/3条, 触发等待占近 2/3)——adc_triage 模式
    由调用方压到 5s; 换板后恢复默认 20s 并真机核实 Completed 语义。
    recover: 源输出关死救回回调。开关写失败最常见原因是角度帧关死输出→电表断电
    (2026-07-16 台面实证), 此时先救源再重试一次; 仅当电表确在线(疑寄存器拒写)才
    退化等周期检测, 离线且救源无效则放弃触发(读码阶段还有一次救源机会)。
    返回触发过程描述(写入 Check.detail 留证)。
    """
    note = ""
    for attempt in (1, 2):
        try:
            vf.client.write(WIRE_CHECK_SWITCH, 1, check_range=False)   # OFF
            time.sleep(1.0)
            vf.client.write(WIRE_CHECK_SWITCH, 0, check_range=False)   # ON → 立即检查一次
            break
        except Exception as exc:
            if attempt == 1 and recover is not None:
                try:
                    recover()
                    note = "开关首写失败→救源后重试; "
                    continue
                except Exception as rexc:
                    return (f"开关写入失败({exc}), 救源亦失败({type(rexc).__name__}: {rexc})"
                            " → 放弃触发, 读码留证")
            try:
                ha.wait_alive_via(vf, timeout_s=3.0)
            except RuntimeError:
                return f"{note}开关写入失败({exc}) 且电表离线 → 放弃触发, 读码留证"
            time.sleep(TRIGGER_FALLBACK_S)
            return (f"{note}开关写入失败({exc}), 电表在线(疑拒写)"
                    f" → 退化等待周期检测 {TRIGGER_FALLBACK_S}s")
    t0 = time.time()
    status = None
    while time.time() - t0 < completed_s:
        try:
            status = int(vf.client.read_addr(WIRE_CHECK_STATUS))
            if status == 2:
                time.sleep(CODE_SETTLE_S)
                return f"{note}开关重开触发, status=Completed({time.time() - t0:.0f}s)"
        except Exception:
            pass
        time.sleep(1.0)
    time.sleep(CODE_SETTLE_S)
    return (f"{note}开关重开触发, 但 {completed_s:.0f}s 内 status={status}"
            f"(未达 Completed), 继续读码留证")


def _assert_point(vf: Verifier, report: Report, idx: int, expect: dict,
                  trigger_note: str, measure_ready: bool, record_only: bool,
                  recover: "Callable[[], None] | None" = None,
                  retrigger: "Callable[[], str] | None" = None) -> None:
    """读 12290/12291(/12304) 并按期望精确比对; i_assert=false 的点电流码只记录。

    recover/retrigger: 读码失败(源输出关死电表断电)时一次性救源+重触发+重读,
    避免整点判据以"Modbus 读取失败"作废(2026-07-16 3E4WY 批 19 条读失败的整改)。
    """
    ready_note = "" if measure_ready else "; ⚠️测量未就绪(频率读0, ADC口径), 读数存疑"
    recover_used = False

    def _read(reader: "Callable[[], object]"):
        """读寄存器; 首次异常且有救源回调时救源+重触发后重读一次。"""
        nonlocal recover_used, trigger_note
        try:
            return reader()
        except Exception:
            if recover_used or recover is None:
                raise
            recover_used = True
            recover()
            if retrigger is not None:
                trigger_note = f"{retrigger()}; 前序读码失败已救源重触发"
            return reader()
    checks: "list[tuple[str, int, int | None, dict, bool]]" = [
        ("电压错误码", V_ERROR_CODE, expect.get("v_code"), V_BITS, True),
        ("电流错误码", I_ERROR_CODE, expect.get("i_code"), I_BITS,
         bool(expect.get("i_assert", True))),
    ]
    for label, addr, exp, names, do_assert in checks:
        if exp is None:
            continue
        try:
            val = int(_read(lambda a=addr: vf.client.read_addr(a)))
        except Exception as exc:
            report.add(Check(label=f"点{idx} {label}", expected=_fmt(exp, names),
                             actual="(Modbus 读取失败)", passed=False,
                             detail=f"addr={addr}; {type(exc).__name__}: {exc}"))
            continue
        ok = val == exp
        if not do_assert:
            report.add(Check(label=f"点{idx} {label}(记录)", expected="不断言(仅记录)",
                             actual=_fmt(val, names), passed=True,
                             detail=f"addr={addr}; 电压缺失类用例该相 pf 无定义, 电流码不断言"
                                    f"(2026-07-14 确认); 参考期望 {_fmt(exp, names)}"
                                    f"{ready_note}; {trigger_note}"))
            continue
        report.add(Check(label=f"点{idx} {label}", expected=_fmt(exp, names),
                         actual=_fmt(val, names),
                         passed=True if record_only else ok,
                         detail=f"addr={addr}{ready_note}; {trigger_note}" + (
                             f"; 记录模式: 若判决={'PASS' if ok else 'FAIL'}" if record_only else "")))
    if expect.get("phase_order") is not None:
        exp_po = int(expect["phase_order"])
        try:
            po = int(float(_read(lambda: vf.read_truth(VOLTAGE_PHASE_ORDER))))
            ok = po == exp_po
            report.add(Check(label=f"点{idx} voltagePhaseOrder", expected=str(exp_po),
                             actual=str(po), passed=True if record_only else ok,
                             detail=f"addr={VOLTAGE_PHASE_ORDER}(0x3010){ready_note}" + (
                                 f"; 记录模式: 若判决={'PASS' if ok else 'FAIL'}" if record_only else "")))
        except Exception as exc:
            report.add(Check(label=f"点{idx} voltagePhaseOrder", expected=str(exp_po),
                             actual="(Modbus 读取失败)", passed=False,
                             detail=f"addr={VOLTAGE_PHASE_ORDER}; {type(exc).__name__}: {exc}"))


def _point_expect_labels(expect: dict) -> "list[str]":
    """失联时用于逐条记 FAIL 的判据标签清单。"""
    labels = []
    if expect.get("v_code") is not None:
        labels.append("电压错误码")
    if expect.get("i_code") is not None:
        labels.append("电流错误码")
    if expect.get("phase_order") is not None:
        labels.append("voltagePhaseOrder")
    return labels


def run_wiring_case(case_meta: dict, config_path: str = CONFIG) -> Report:
    """按 case_map_wiring.yaml 逐测点执行: 设源→触发检测→读错误码→位断言→保活退场。

    点结构约定: 每条故障用例 points = [正常基准(期望全0), 故障点(期望置位), 还原点(期望全0)],
    与用例表步骤 2/3(基准)→4/5(故障)→6(还原清除) 一一对应; 基准/相序类用例为单点。
    """
    get_config(config_path)
    run_cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
    record_only = str(run_cfg.get("assert_mode", "enforce")).lower() == "record"
    adc_triage = bool(run_cfg.get("adc_triage", False))
    gate_s = 5.0 if adc_triage else ha.MEASURE_READY_S
    entry = load_case(case_meta["编号"])
    if entry.get("needs_review"):
        raise RuntimeError(f"{case_meta['编号']} 的映射仍带 needs_review 标记, 未确认不得执行: "
                           f"{entry['needs_review']}")
    report = Report(title=f"{case_meta['编号']} {case_meta['标题']}", started=now())
    src, settle_s = ha.source_from_config()
    try:
        # 自供电表 · 最小源交互(与 run_accuracy_case 相同的会话开场)
        if src.is_inited():
            try:
                boot = ha.wait_alive(config_path, timeout_s=ha.DEAD_DETECT_S)
            except RuntimeError:
                print("[0] 用例开头电表不在线 → 源冷重臂(切屏重开输出)")
                src.reinit_output()
                src.set_point(ha.keepalive_point(), settle_s=3)
                boot = ha.wait_alive(config_path)
        else:
            src.warm_init()
            src.set_point(ha.keepalive_point(), settle_s=3)
            try:
                boot = ha.wait_alive(config_path, timeout_s=ha.WARM_PROBE_S)
            except RuntimeError:
                print("[0] 暖启动探测未通, 转源冷初始化(输出将瞬断, 电表冷重启)")
                src.init()
                src.set_point(ha.keepalive_point(), settle_s=3)
                boot = ha.wait_alive(config_path)
            else:
                src.mark_inited()
        print(f"[0] 电表 Modbus 就绪(等待 {boot:.0f}s), 源保持常驻输出")
        with Verifier() as vf:
            ha.apply_meter_cfg(vf, src, entry.get("meter_cfg"), report, adc_triage=adc_triage)
            _apply_phase_order(vf, entry.get("phase_order"), report)
            for i, pt in enumerate(entry["points"], 1):
                s = ha.guard_point(pt["source"])
                src.set_point(s, settle_s=settle_s)
                alive = None
                exc_last: "Exception | None" = None
                for rescue in ("probe", "resend", "reinit"):      # 三级救源(同 accuracy)
                    try:
                        if rescue == "resend":
                            print(f"[点{i}] 设源后 {ha.DEAD_DETECT_S}s 无应答 → 重发本测点救源")
                            src.set_point(s, settle_s=3, force=True)
                        elif rescue == "reinit":
                            print(f"[点{i}] 重发无效 → 源冷重臂(切屏重开输出)后重发本测点")
                            src.reinit_output()
                            src.set_point(s, settle_s=3, force=True)
                        alive = ha.wait_alive_via(
                            vf, timeout_s=ha.DEAD_DETECT_S if rescue == "probe" else ha.REARM_BOOT_S)
                        time.sleep(RIDE_THROUGH_CONFIRM_S)       # 防电容余电假活: 隔 2s 复确认
                        ha.wait_alive_via(vf, timeout_s=3.0)     # 复确认不过则进下一级救源
                        break
                    except RuntimeError as exc:
                        exc_last = exc
                if alive is None:
                    for label in _point_expect_labels(pt["expect"]):
                        report.add(Check(label=f"点{i} {label}", expected="(见 case_map 期望)",
                                         actual="(设源后电表失联, 重发/冷重臂救源均无效)",
                                         passed=False, detail=f"source={s}; {exc_last}"))
                    src.set_point(ha.keepalive_point(), settle_s=3, force=True)
                    ha.wait_alive_via(vf)
                    continue
                if alive > 5:
                    print(f"[点{i}] 源切换后电表重启, 等待 {alive:.0f}s 复活")
                measure_ready = ha.wait_measure_ready(vf, timeout_s=gate_s)
                if not measure_ready:
                    print(f"[点{i}] 测量未就绪(频率读0/读不通, ADC口径) → 仍触发检测读码留证")

                def _recover(point=s, idx=i) -> None:
                    """源输出关死救回: force 重发本测点, 有源电表 3s 内上线(2026-07-16 实证)。"""
                    print(f"[点{idx}] 触发/读码阶段电表离线 → force 重发本测点救源")
                    src.set_point(point, settle_s=3, force=True)
                    ha.wait_alive_via(vf, timeout_s=ha.REARM_BOOT_S)

                trig_s = 5.0 if adc_triage else TRIGGER_COMPLETED_S
                trigger_note = _trigger_check(vf, completed_s=trig_s, recover=_recover)
                _assert_point(vf, report, i, pt["expect"], trigger_note, measure_ready,
                              record_only, recover=_recover,
                              retrigger=lambda ts=trig_s: _trigger_check(vf, completed_s=ts))
    finally:
        try:
            _exit_switch_on()
        finally:
            try:
                ha.exit_idle(src, config_path)      # 保活(Ua=100V)+探活, 杜绝源留 0V
            finally:
                report.save(f"wiring_{case_meta['编号']}")
    return report


def _exit_switch_on() -> None:
    """退场兜底: 确保接线检查开关恢复 ON(0, 出厂默认启用); 失败仅告警不阻塞退场。"""
    try:
        with Verifier() as vf:
            if int(vf.client.read_addr(WIRE_CHECK_SWITCH)) != 0:
                vf.client.write(WIRE_CHECK_SWITCH, 0, check_range=False)
                print("[退场] WIRE_CHECK_SWITCH 已恢复 ON(0)")
    except Exception as exc:
        print(f"[退场] ⚠️ WIRE_CHECK_SWITCH 状态确认失败, 需人工核对开关为 ON: {exc}")
