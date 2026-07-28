import os
import csv

STAT_FIELDS = ["period", "pw", "duty", "frqdev", "frq"]
_STAT_KEYS = ["value", "mean", "min", "max", "sdev"]
_FIELD_CN = {"period": "周期", "pw": "正脉宽", "duty": "占空比",
             "frqdev": "频率偏差", "frq": "频率"}
_STAT_CN = {"value": "Value", "mean": "Mean", "min": "Min", "max": "Max", "sdev": "Sdev"}

def _build_columns():
    cols = ["时间戳", "周期数N", "采集时长s", "样本数Num"]
    for f in STAT_FIELDS:
        for k in _STAT_KEYS:
            cols.append(f"{_FIELD_CN[f]}_{_STAT_CN[k]}")
    cols += ["模式", "高频抑制", "参考频率", "触发电平"]
    return cols

COLUMNS = _build_columns()

def row_from_result(timestamp, result, config):
    stats = result["stats"]
    row = [timestamp, result["n_periods"], round(result["duration"], 3), stats["num"]]
    for f in STAT_FIELDS:
        for k in _STAT_KEYS:
            v = stats[f][k]
            row.append("" if v is None else v)
    row += [config.get("mode", ""), config.get("hfr", ""),
            config.get("refq", ""), config.get("trg", "")]
    return row

def append_csv(path, row):
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(COLUMNS)
        w.writerow(row)

def row_to_tsv(row):
    return "\t".join("" if c is None else str(c) for c in row)
