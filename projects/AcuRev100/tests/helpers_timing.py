"""AcuRev-100 计时类用例运行时（011: 时钟设置/PC对时/运行·负载时间累计/计时器重置）。

职责: 按 case_map.yaml 的 type 分派——
- clock_set:          遍历写时钟(4128~4133)→读回→总秒差 [0, tol_s] 判定→结束还原 PC 时间
- counter_accumulate: 读计数器→(可选设源)→等待→再读→Δ≈期望(等待时长或0)判定
- op_clear:           写清除寄存器(4404/4405)→计数器读值≈0 判定

设计说明:
- 计时链路与 ADC 无关(RTC/秒计数), ADC 损坏期间 run time 类判据有机会真 PASS;
  负载时间依赖测量(电流>Ist 判负载), ADC 读零期间不会累计, FAIL 按已知口径留证。
- "掉电重启后时间/计数保持"类步骤会给自供电表断电, 由 config run.allow_power_cycle
  把关(默认 false → 该步骤记 MANUAL 不计判据; 用户确认后置 true 才自动执行)。
- counter 判据是相对量(Δ≈实际等待时长), adc_triage 缩短等待窗后判据依然成立
  (与能量的固定 kWh 判据不同), 故短窗照常判决。
"""
from __future__ import annotations

import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[1]          # projects/AcuRev100 (本文件在 tests/ 根)
REPO_ROOT = PROJECT_ROOT.parents[1]
if str(REPO_ROOT) not in sys.path:                           # 直跑本文件时兜底
    sys.path.insert(0, str(REPO_ROOT))

from comm.ctl_acuview.config import get_config  # noqa: E402
from comm.ctl_acuview.verify import Check, Report, Verifier, now  # noqa: E402
from projects.AcuRev100.tests.helpers_accuracy import (  # noqa: E402
    CONFIG, DEAD_DETECT_S, REARM_BOOT_S, WARM_PROBE_S, _apply_meter_cfg, _exit_idle,
    _guard_point, _keepalive, _source_from_config, _wait_alive, _wait_alive_via, load_case)

CLOCK_ADDRS = (("year", 4128), ("month", 4129), ("day", 4130),
               ("hour", 4131), ("minute", 4132), ("second", 4133))
CLOCK_SETTLE_S = 2         # 写完时钟到读回的间隔(覆盖固件应用延迟)
COUNTER_TRIAGE_WAIT_S = 90  # adc_triage 时计数器段的等待窗(判据随实际等待缩放, 依然有效)
OP_CLEAR_TOL_S = 10        # 清除后计数器允许的残留读值(清除瞬间已重新计数)
POWER_CYCLE_OFF_S = 4      # 掉电重启: 三相归零保持时长(自供电表必死)


def _read_clock(vf: Verifier) -> datetime:
    v = {k: int(vf.client.read_addr(a)) for k, a in CLOCK_ADDRS}
    return datetime(v["year"], v["month"], v["day"], v["hour"], v["minute"], v["second"])


def _write_clock(vf: Verifier, dt: datetime) -> None:
    """写时钟, 规避固件的"组合日期逐字段校验"(2026-07-14 实测: 当前日=31 时写 MONTH=11
    会被拒 exception 4, 因瞬时组合 11-31 非法): 先把 DAY 置 1(任何年月都合法), 再写
    年→月→真实日→时→分→秒。"""
    order = [("day", 4130, 1), ("year", 4128, dt.year), ("month", 4129, dt.month),
             ("day", 4130, dt.day), ("hour", 4131, dt.hour),
             ("minute", 4132, dt.minute), ("second", 4133, dt.second)]
    for _k, addr, val in order:
        vf.client.write(addr, val)


def _session_boot(src, config_path: str) -> float:
    """会话/用例开头确保电表在线(与 run_accuracy_case 同一套暖启动+冷重臂逻辑)。"""
    if src.is_inited():
        try:
            return _wait_alive(config_path, timeout_s=DEAD_DETECT_S)
        except RuntimeError:
            print("[0] 用例开头电表不在线 → 源冷重臂(切屏重开输出)")
            src.reinit_output()
            src.set_point(_keepalive(), settle_s=3)
            return _wait_alive(config_path)
    src.warm_init()
    src.set_point(_keepalive(), settle_s=3)
    try:
        boot = _wait_alive(config_path, timeout_s=WARM_PROBE_S)
    except RuntimeError:
        print("[0] 暖启动探测未通, 转源冷初始化(输出将瞬断, 电表冷重启)")
        src.init()
        src.set_point(_keepalive(), settle_s=3)
        return _wait_alive(config_path)
    src.mark_inited()
    return boot


def _set_point_rescued(src, vf: Verifier, s: dict, settle_s: float, label: str) -> bool:
    """设源+三级救源(同 run_accuracy_case); 返回是否在线。"""
    src.set_point(s, settle_s=settle_s)
    for rescue in ("probe", "resend", "reinit"):
        try:
            if rescue == "resend":
                print(f"[{label}] 设源后 {DEAD_DETECT_S}s 无应答 → 重发本测点救源")
                src.set_point(s, settle_s=3, force=True)
            elif rescue == "reinit":
                print(f"[{label}] 重发无效 → 源冷重臂(切屏重开输出)后重发本测点")
                src.reinit_output()
                src.set_point(s, settle_s=3, force=True)
            _wait_alive_via(vf, timeout_s=DEAD_DETECT_S if rescue == "probe" else REARM_BOOT_S)
            return True
        except RuntimeError:
            continue
    src.set_point(_keepalive(), settle_s=3, force=True)
    _wait_alive_via(vf)
    return False


def _power_cycle(src, vf: Verifier, config_path: str) -> float:
    """掉电重启(仅 run.allow_power_cycle=true 时被调用): 三相归零→保持→保活拉回→等复活。

    ⚠️ 电表断电时 USB 虚拟串口(COM6 由电表供电)整个消失, vf 的旧串口句柄永久失效
    (2026-07-14 实测: 复用旧句柄等复活烧满 240s 超时)——必须先关句柄, 用独立短连
    客户端轮询复活(每轮新开串口, 设备重枚举后 COM 号不变), 复活后再重开 vf 句柄。
    """
    dead = dict(_keepalive())
    dead.update(ua=0.0, ub=0.0, uc=0.0, ia=0.0, ib=0.0, ic=0.0)
    src.set_point(dead, settle_s=POWER_CYCLE_OFF_S)      # 故意断电, 不走 _guard_point
    src.set_point(_keepalive(), settle_s=3)
    try:
        vf.client.close()
    except Exception:
        pass
    try:
        boot = _wait_alive(config_path, timeout_s=DEAD_DETECT_S * 2)
    except RuntimeError:
        # 全零→有压切换高发"输出停0卡死"(2026-07-14 case2/5 实测: 只发保活帧拉不回,
        # 傻等240s超时且批跑结束源留在0) → 冷重臂(切屏重开输出)再拉一次
        print("[cycle] 掉电后保活未拉回(源输出疑停0) → 源冷重臂重开输出")
        src.reinit_output()
        src.set_point(_keepalive(), settle_s=3)
        boot = _wait_alive(config_path)
    vf.client.connect()
    return boot


def run_timing_case(case_meta: dict, config_path: str = CONFIG) -> Report:
    get_config(config_path)
    run_cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")).get("run") or {}
    adc_triage = bool(run_cfg.get("adc_triage", False))
    allow_cycle = bool(run_cfg.get("allow_power_cycle", False))
    entry = load_case(case_meta["编号"])
    if entry.get("needs_review"):
        raise RuntimeError(
            f"{case_meta['编号']} 的映射仍带 needs_review 标记, 未确认不得执行: "
            f"{entry['needs_review']}")
    report = Report(title=f"{case_meta['编号']} {case_meta['标题']}", started=now())
    src, settle_s = _source_from_config()
    try:
        boot = _session_boot(src, config_path)
        print(f"[0] 电表 Modbus 就绪(等待 {boot:.0f}s), 源保持常驻输出")
        with Verifier() as vf:
            _apply_meter_cfg(vf, src, entry.get("meter_cfg"), report, adc_triage=adc_triage)
            kind = entry["type"]
            if kind == "clock_set":
                _run_clock_set(entry, src, vf, report, allow_cycle, config_path)
            elif kind == "counter_accumulate":
                _run_counter(entry, src, vf, report, adc_triage, settle_s, allow_cycle,
                             config_path)
            elif kind == "op_clear":
                _run_op_clear(entry, vf, report)
            else:
                raise RuntimeError(f"未知计时用例类型: {kind}")
    finally:
        try:
            _exit_idle(src, config_path)             # 保活(Ua=100V)+探活校验(2026-07-14)
        finally:
            report.save(f"auto_{case_meta['编号']}")
    return report


def _run_clock_set(entry: dict, src, vf: Verifier, report: Report, allow_cycle: bool,
                   config_path: str) -> None:
    tol_s = float(entry.get("tol_s", 8))
    for i, spec in enumerate(entry["clock_points"], 1):
        target = datetime.now().replace(microsecond=0) if spec == "PC" \
            else datetime.strptime(spec, "%Y-%m-%d %H:%M:%S")
        label = "PC时间" if spec == "PC" else spec
        try:
            _write_clock(vf, target)
            time.sleep(CLOCK_SETTLE_S)
            got = _read_clock(vf)
            delta = (got - target).total_seconds()
            ok = 0 <= delta <= tol_s + CLOCK_SETTLE_S
            report.add(Check(label=f"点{i} 时钟设置 {label}",
                             expected=f"读回-设定 ∈ [0, {tol_s + CLOCK_SETTLE_S:.0f}]s(含读写延迟)",
                             actual=f"读回 {got} (Δ={delta:.0f}s)", passed=ok,
                             detail="寄存器 4128~4133 逐个 FC6 写入; 总秒差比较覆盖分钟滚动/闰年"))
        except Exception as exc:
            report.add(Check(label=f"点{i} 时钟设置 {label}", expected="写入并读回成功",
                             actual=f"(异常: {type(exc).__name__}: {exc})", passed=False,
                             detail="寄存器 4128~4133"))
    if allow_cycle:
        try:
            before = _read_clock(vf)
            t0 = time.time()
            alive = _power_cycle(src, vf, config_path)
            got = _read_clock(vf)
            expect_lo = before + timedelta(seconds=0)
            expect_hi = before + timedelta(seconds=time.time() - t0 + 5)
            ok = expect_lo <= got <= expect_hi
            report.add(Check(label="掉电重启后时间保持", expected=f"{expect_lo} ~ {expect_hi}",
                             actual=f"{got} (复活{alive:.0f}s)", passed=ok,
                             detail="run.allow_power_cycle=true; 时钟应由 RTC 电池/超级电容维持"))
        except Exception as exc:
            report.add(Check(label="掉电重启后时间保持", expected="断电→复活→时钟连续",
                             actual=f"(异常: {type(exc).__name__}: {exc})", passed=False,
                             detail="run.allow_power_cycle=true; 掉电流程异常, 留证不中断"))
    else:
        print("[clock] 掉电重启验证步骤: run.allow_power_cycle=false → MANUAL(不计判据)")
    # 还原: 时钟回 PC 时间(遍历残留 2099 等值会影响后续日志/时标)
    restore = datetime.now().replace(microsecond=0)
    try:
        _write_clock(vf, restore)
        time.sleep(CLOCK_SETTLE_S)
        got = _read_clock(vf)
        ok = 0 <= (got - restore).total_seconds() <= tol_s + CLOCK_SETTLE_S
        report.add(Check(label="用例还原: 时钟回 PC 时间", expected=f"≈{restore}",
                         actual=str(got), passed=ok,
                         detail="遍历残留值必须还原, 还原失败需人工恢复(置顶告警)"))
    except Exception as exc:
        report.add(Check(label="用例还原: 时钟回 PC 时间", expected=f"≈{restore}",
                         actual=f"(还原异常: {exc})", passed=False, detail="需人工恢复电表时钟"))


def _run_counter(entry: dict, src, vf: Verifier, report: Report,
                 adc_triage: bool, settle_s: float, allow_cycle: bool,
                 config_path: str) -> None:
    counter = entry["counter"]
    for i, seg in enumerate(entry["segments"], 1):
        wait_s = float(seg.get("wait_s", 300))
        if adc_triage:
            wait_s = min(wait_s, COUNTER_TRIAGE_WAIT_S)   # 判据随实际等待缩放, 短窗判决有效
        tol_s = float(seg.get("tol_s", 10))
        label = seg.get("label", f"段{i}")
        if seg.get("source"):
            if not _set_point_rescued(src, vf, _guard_point(dict(seg["source"])),
                                      settle_s, f"段{i}"):
                report.add(Check(label=f"段{i} {label}", expected="设源成功",
                                 actual="(设源后电表失联, 救源均无效)", passed=False,
                                 detail=f"counter={counter}"))
                continue
        try:
            t1 = float(vf.read_truth(counter))
            t_start = time.time()
            time.sleep(wait_s)
            waited = time.time() - t_start
            t2 = float(vf.read_truth(counter))
        except Exception as exc:
            report.add(Check(label=f"段{i} {label}", expected="计数器可读",
                             actual=f"(读取失败: {type(exc).__name__}: {exc})", passed=False,
                             detail=f"counter={counter}"))
            continue
        delta = t2 - t1
        if seg["expect"] == "wait":
            ok = abs(delta - waited) <= tol_s
            exp = f"Δ≈{waited:.0f}s(实际等待)±{tol_s:.0f}s"
        else:                                            # expect == "zero"
            ok = 0 <= delta <= tol_s
            exp = f"Δ∈[0, {tol_s:.0f}]s(不累计)"
        report.add(Check(label=f"段{i} {label}", expected=exp,
                         actual=f"Δ={delta:.0f}s (t1={t1:.0f}, t2={t2:.0f})", passed=ok,
                         detail=f"counter={counter}; 等待{waited:.0f}s"
                               + ("; adc_triage 短窗(判据随实际等待缩放, 判决有效)"
                                  if adc_triage else "")))
    if allow_cycle and entry.get("cycle_keep_check"):
        try:
            before = float(vf.read_truth(counter))
            alive = _power_cycle(src, vf, config_path)
            after = float(vf.read_truth(counter))
            ok = after >= before - 5                      # 掉电保持(重启后继续计数只增不减)
            report.add(Check(label="掉电重启后计数保持", expected=f"≥{before:.0f}-5",
                             actual=f"{after:.0f} (复活{alive:.0f}s)", passed=ok,
                             detail=f"counter={counter}; run.allow_power_cycle=true"))
        except Exception as exc:
            report.add(Check(label="掉电重启后计数保持", expected="断电→复活→计数保持",
                             actual=f"(异常: {type(exc).__name__}: {exc})", passed=False,
                             detail=f"counter={counter}; 掉电流程异常, 留证不中断"))
    elif entry.get("cycle_keep_check"):
        print("[counter] 掉电重启保持验证: run.allow_power_cycle=false → MANUAL(不计判据)")


def _run_op_clear(entry: dict, vf: Verifier, report: Report) -> None:
    op = entry["op"]
    counter = op["counter"]
    before = float(vf.read_truth(counter))
    vf.client.write(int(op["clear_addr"]), 1, check_range=False)
    time.sleep(3)
    after = float(vf.read_truth(counter))
    ok = 0 <= after <= OP_CLEAR_TOL_S
    report.add(Check(label=f"清除 {counter}",
                     expected=f"清除后读值 ∈ [0, {OP_CLEAR_TOL_S}]s",
                     actual=f"清除前 {before:.0f}s → 清除后 {after:.0f}s", passed=ok,
                     detail=f"写 {op['clear_addr']}=1(Clearing); 清除即刻重新计数故留容差"))
