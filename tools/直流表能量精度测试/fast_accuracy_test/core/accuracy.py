from dataclasses import dataclass
from statistics import mean

EPS = 1e-9


@dataclass
class Metric:
    label: str
    unit: str
    ref: float
    min: float = None
    max: float = None
    avg: float = None
    err_min: float = None      # 最小值的误差
    err_max: float = None      # 最大值的误差（注意：是“最大值的误差”，不是“最大的误差”）
    err_avg: float = None      # 平均值的误差
    err_worst: float = None    # 最大误差 = 离真值最远端点的误差 = max(err_min, err_max)
    threshold_pct: float = None
    result: str = "N/A"


def _err(x, ref):
    return abs(x - ref) / abs(ref) * 100


def evaluate(label, unit, ref, samples, threshold_pct):
    m = Metric(label=label, unit=unit, ref=ref, threshold_pct=threshold_pct)
    if abs(ref) < EPS or not samples:
        m.result = "N/A"
        return m
    m.min, m.max, m.avg = min(samples), max(samples), mean(samples)
    m.err_min = _err(m.min, ref)
    m.err_max = _err(m.max, ref)
    m.err_avg = _err(m.avg, ref)
    m.err_worst = max(m.err_min, m.err_max)   # 真正的“最大误差”：离真值最远端点的误差
    within = all(e <= threshold_pct for e in (m.err_min, m.err_max, m.err_avg))
    m.result = "Passed" if within else "Failed"
    return m


def evaluate_direct(label, unit, measured, threshold_pct):
    """measured 本身就是误差值(如设备给的脉冲误差%)，直接和阈值比绝对值。
    measured 为 None → N/A(没测到)。"""
    m = Metric(label=label, unit=unit, ref=0.0, threshold_pct=threshold_pct)
    if measured is None:
        m.result = "N/A"
        return m
    e = abs(measured)
    m.min = m.max = m.avg = measured
    m.err_min = m.err_max = m.err_avg = m.err_worst = e
    if threshold_pct is None:
        m.result = "N/A"          # 没填阈值：仅记录测量值，不判定
    else:
        m.result = "Passed" if e <= threshold_pct else "Failed"
    return m


def combine(metrics):
    return "Failed" if any(m.result == "Failed" for m in metrics) else "Passed"
