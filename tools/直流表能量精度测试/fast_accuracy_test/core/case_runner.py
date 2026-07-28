import logging
import time
from dataclasses import dataclass

from .accuracy import evaluate, evaluate_direct, combine, Metric

SETTLE_S = 5.0  # 控源输出后等待源稳定再记录(秒)


@dataclass
class CaseResult:
    case: object
    metrics: list
    overall: str
    pulse_samples: list = None      # 能量脉冲误差的逐次采样(仅测脉冲的用例有；写进结果表)


def _sample_names(is_dual):
    if is_dual:
        return ["voltage", "current_1", "current_2", "power_1", "power_2", "power_sum"]
    return ["voltage", "current", "power"]


def _read_into(reader, name, buf, failed):
    try:
        buf[name].append(reader.read(name))
    except IOError:
        failed.add(name)


# 源稳定判定参数
STABILIZE_MAX_EXTRA = 15.0   # settle_s 之后额外轮询等待源稳定的最长秒数
STABILIZE_STEP = 0.5         # 轮询间隔(秒)
STABILIZE_TOL = 0.01         # 连续两次读数相对变化阈值(1%)，小于则认为已稳定


def _stabilize(reader, settle_s, sleep_fn, probe="voltage", target=None):
    """等源真正稳定后再开始记录：先等固定 settle_s，再轮询 probe(默认电压)读数，
    直到连续两次读数非零且相对变化 < STABILIZE_TOL，认为已稳定；最多再等
    STABILIZE_MAX_EXTRA 秒。读不到/读到~0 则重置继续等。
    target 为该项设定值；若设定值≈0(如纯电流点电压目标为0)则跳过轮询。"""
    sleep_fn(settle_s)
    if target is not None and abs(target) < 1e-9:
        return
    prev = None
    waited = 0.0
    while waited < STABILIZE_MAX_EXTRA:
        try:
            v = reader.read(probe)
        except IOError:
            v = None
        if v is not None and abs(v) > 1e-9:
            if prev is not None and abs(v - prev) <= STABILIZE_TOL * abs(v):
                return
            prev = v
        else:
            prev = None
        sleep_fn(STABILIZE_STEP)
        waited += STABILIZE_STEP


def _drop_zero_samples(buf, names):
    """丢弃 V/I/P 采样里疑似源未稳定的 0 读数(仅当该项还有非零样本时)，
    防止一个 0 把最小值/平均值带偏。若某项全为 0(源真没输出)则保留，
    让其在精度判定中暴露为失败而非被掩盖。返回丢弃总数。"""
    dropped = 0
    for name in names:
        vals = buf.get(name) or []
        nz = [v for v in vals if abs(v) > 1e-9]
        if nz and len(nz) != len(vals):
            dropped += len(vals) - len(nz)
            buf[name] = nz
    return dropped


def _collect_fixed(reader, names, n, interval, sleep_fn):
    buf = {name: [] for name in names}
    failed = set()
    for _ in range(max(1, n)):
        sleep_fn(interval)
        for name in names:
            _read_into(reader, name, buf, failed)
    return buf, failed


def _collect_duration(reader, names, duration_s, interval, sleep_fn, time_fn):
    buf = {name: [] for name in names}
    failed = set()
    end = time_fn() + duration_s
    while time_fn() <= end:
        for name in names:
            _read_into(reader, name, buf, failed)
        sleep_fn(interval)
    if not any(buf[n] for n in names) and not failed:
        for name in names:
            _read_into(reader, name, buf, failed)
    return buf, failed


def _energy_specs(is_dual, current_1, current_2):
    if is_dual:
        return [
            ("Energy_1", current_1, "import_energy_1" if current_1 >= 0 else "export_energy_1"),
            ("Energy_2", current_2, "import_energy_2" if current_2 >= 0 else "export_energy_2"),
        ]
    return [("Energy", current_1, "import_energy" if current_1 >= 0 else "export_energy")]


def _measure(label, unit, ref, name, buf, failed, threshold):
    if name in failed:
        return Metric(label=label, unit=unit, ref=ref, threshold_pct=threshold, result="Failed")
    return evaluate(label, unit, ref, buf.get(name, []), threshold)


def _has_pulse(case):
    # 填了脉冲常数即测能量脉冲误差；能量脉冲误差列(阈值)可空：空则只记录不判定(N/A)
    return getattr(case, "pulse_const", None) is not None


def run_case(case, reader, source, *, is_dual, settle_s=SETTLE_S, include_pulse=False,
             sleep_fn=time.sleep, time_fn=time.time):
    names = _sample_names(is_dual)
    # 测脉冲的用例：把脉冲常数同时设到“两个设备”——被检表(Modbus 写寄存器)和源(如 XL9600 参数配置)
    if include_pulse and _has_pulse(case):
        wfn = getattr(reader, "write_pulse_const", None)
        if wfn is not None:
            try:
                wfn(case.pulse_const)
            except Exception as e:
                logging.warning("用例 %s 写电表脉冲常数失败: %s", case.test_case, e)
        try:
            source.set_pulse_const(case.pulse_const)
        except Exception as e:
            logging.warning("用例 %s 设源脉冲常数失败: %s", case.test_case, e)
    source.output(case.voltage, case.current_1)
    # 等源真正稳定后再记录：固定 settle_s + 主动轮询电压稳定
    _stabilize(reader, settle_s, sleep_fn, target=case.voltage)

    measure_energy = case.wait_h > 0
    energy_specs = _energy_specs(is_dual, case.current_1, case.current_2)
    e_start, e_end, e_failed = {}, {}, set()

    if measure_energy:
        for _, _, reg in energy_specs:
            try:
                e_start[reg] = reader.read(reg)
            except IOError:
                e_failed.add(reg)
        buf, failed = _collect_duration(reader, names, case.wait_h * 3600,
                                        case.sample_interval, sleep_fn, time_fn)
        for _, _, reg in energy_specs:          # 能量结束值在停源前读(能量窗口=等待时间)
            try:
                e_end[reg] = reader.read(reg)
            except IOError:
                e_failed.add(reg)
    else:
        buf, failed = _collect_fixed(reader, names, case.sample_cnt, case.sample_interval, sleep_fn)

    # 剔除疑似源未稳定的 0 读数(否则最小值会被钉死成 0)
    dropped = _drop_zero_samples(buf, names)
    if dropped:
        logging.warning("用例 %s 丢弃 %d 个疑似未稳定的 0 读数", case.test_case, dropped)

    # 脉冲误差检测：源还在输出时进行（仅当本 case 两列都填且源支持）
    pulse_measured = None
    pulse_samples = None
    if include_pulse and _has_pulse(case):
        fn = getattr(source, "measure_pulse_error", None)
        if fn is not None:
            try:
                pulse_measured = fn(case.pulse_const)
                sfn = getattr(source, "last_pulse_samples", None)   # 取逐次采样(写进结果表)
                if sfn is not None:
                    pulse_samples = sfn()
            except Exception as e:
                # 脉冲误差测不到(如超时)只影响该项(N/A)，不炸掉整条用例的 V/I/P 结果
                logging.warning("用例 %s 脉冲误差测量失败: %s", case.test_case, e)
        if pulse_measured is None:
            logging.warning("用例 %s 需要脉冲误差，但当前控源不支持或未测到", case.test_case)

    source.stop()

    va, ca, pa = (case.voltage_accuracy * 100, case.current_accuracy * 100,
                  case.power_accuracy * 100)
    metrics = []

    if is_dual:
        p1_ref = case.voltage * case.current_1 / 1000
        p2_ref = case.voltage * case.current_2 / 1000
        metrics.append(_measure("Voltage", "V", case.voltage, "voltage", buf, failed, va))
        metrics.append(_measure("Current_1", "A", case.current_1, "current_1", buf, failed, ca))
        metrics.append(_measure("Current_2", "A", case.current_2, "current_2", buf, failed, ca))
        metrics.append(_measure("Power_1", "kW", p1_ref, "power_1", buf, failed, pa))
        metrics.append(_measure("Power_2", "kW", p2_ref, "power_2", buf, failed, pa))
        metrics.append(_measure("Power_Sum", "kW", p1_ref + p2_ref, "power_sum", buf, failed, pa))
    else:
        p_ref = case.voltage * case.current_1 / 1000
        metrics.append(_measure("Voltage", "V", case.voltage, "voltage", buf, failed, va))
        metrics.append(_measure("Current", "A", case.current_1, "current", buf, failed, ca))
        metrics.append(_measure("Power", "kW", p_ref, "power", buf, failed, pa))

    # 能量：首尾差值；读失败则该能量项标 Failed
    for label, cur, reg in energy_specs:
        e_ref = abs(case.voltage * cur * case.wait_h / 1000)
        if not measure_energy:
            metrics.append(Metric(label=label, unit="kWh", ref=0.0, threshold_pct=ca))
        elif reg in e_failed:
            metrics.append(Metric(label=label, unit="kWh", ref=e_ref, threshold_pct=ca, result="Failed"))
        else:
            e_meas = e_end[reg] - e_start[reg]
            metrics.append(evaluate(label, "kWh", e_ref, [e_meas], ca))

    # 能量脉冲误差(可选)：本轮有任一 case 用脉冲时，每行都带此列(不适用则 N/A)
    if include_pulse:
        if _has_pulse(case):
            thr = case.pulse_error_acc * 100 if case.pulse_error_acc is not None else None
            metrics.append(evaluate_direct("Pulse_Error", "%", pulse_measured, thr))
        else:
            metrics.append(Metric(label="Pulse_Error", unit="%", ref=0.0))

    return CaseResult(case=case, metrics=metrics, overall=combine(metrics), pulse_samples=pulse_samples)
