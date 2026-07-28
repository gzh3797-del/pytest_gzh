import statistics
import time

NUMERIC_FIELDS = ["period", "pw", "duty", "frqdev", "frq"]

def _stat_block(vals):
    if not vals:
        return {"value": None, "mean": None, "min": None, "max": None, "sdev": None}
    return {
        "value": vals[-1],
        "mean": statistics.fmean(vals),
        "min": min(vals),
        "max": max(vals),
        "sdev": statistics.stdev(vals) if len(vals) > 1 else 0.0,
    }

def compute_stats(samples):
    out = {}
    for f in NUMERIC_FIELDS:
        vals = [s[f] for s in samples if s.get(f) is not None]
        out[f] = _stat_block(vals)
    out["num"] = len(samples)
    return out

class NoPulseError(Exception):
    pass

def acquire(client, n_periods=10, min_duration=2.0, now=time.monotonic, sleep=time.sleep):
    first = client.query_fcnt()
    period = first.get("period")
    if not period or period <= 0:
        raise NoPulseError("未检测到有效脉冲，无法计算采集时长")
    duration = max(n_periods * period, min_duration)
    samples = []
    start = now()
    while now() - start < duration:
        try:
            s = client.query_fcnt()
            if s.get("frq"):
                samples.append(s)
            sleep(0)
        except Exception:
            sleep(0.05)
    if not samples:
        raise NoPulseError("采集窗口内未取到有效样本")
    return {
        "n_periods": n_periods,
        "duration": duration,
        "stats": compute_stats(samples),
        "samples": samples,
    }
