import os
import time
import openpyxl

# 被测项名 -> 中文(报告表头用)。能量按电流符号读 import(+)/export(-)，故统称“能量”。
_LABEL_CN = {
    "Voltage": "电压",
    "Current": "电流", "Current_1": "电流1", "Current_2": "电流2",
    "Power": "功率", "Power_1": "功率1", "Power_2": "功率2", "Power_Sum": "功率合",
    "Energy": "能量", "Energy_1": "能量1", "Energy_2": "能量2",
    "Pulse_Error": "能量脉冲误差",
}


def _cn(label):
    return _LABEL_CN.get(label, label)


def _metric_columns(metric):
    L = _cn(metric.label)
    u = metric.unit
    if metric.label == "Pulse_Error":
        # 能量脉冲误差：真值列放"允许误差/限值"(= Input 能量脉冲误差 ×100，与实测同为 %)；
        # 加实测误差值 + 判定。min/max/avg/最大误差 省略，逐次原始值见"采样1..N"列。
        return [f"{L}真值(%)", f"{L}(%)", f"{L}判定"]
    return [f"{L}真值({u})", f"{L}最小值", f"{L}最大值", f"{L}平均值",
            f"{L}平均误差(%)", f"{L}最大误差(%)", f"{L}判定"]


def _metric_values(metric):
    if metric.label == "Pulse_Error":
        # 真值列=允许误差限值(threshold_pct = Input 能量脉冲误差×100)；再给实测误差值 + 判定
        return [metric.threshold_pct, metric.err_avg, metric.result]
    # “最大误差”列取 err_worst = max(最小值误差, 最大值误差)，即离真值最远那个点的误差；
    # 不能用 err_max（那是“最大值的误差”，当测量值低于真值时会偏小、甚至小于平均误差）。
    return [metric.ref, metric.min, metric.max, metric.avg,
            metric.err_avg, metric.err_worst, metric.result]


def write_report(results, result_dir, device_model):
    if not results:
        raise ValueError("results must not be empty")
    os.makedirs(result_dir, exist_ok=True)
    fname = f"precision_measure_{device_model}_{time.strftime('%Y%m%d%H%M%S')}.xlsx"
    path = os.path.join(result_dir, fname)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "precision"

    template = results[0].metrics
    header = ["用例"]
    for m in template:
        header += _metric_columns(m)
    # 能量脉冲误差的逐次采样(若有)：追加列，放在“总判定”之前；没有脉冲测试则不加列(向后兼容)
    max_pulse = max((len(getattr(r, "pulse_samples", None) or []) for r in results), default=0)
    for k in range(max_pulse):
        header.append("能量脉冲误差采样%d(%%)" % (k + 1))
    header.append("总判定")
    ws.append(header)

    for res in results:
        row = [res.case.test_case]
        for m in res.metrics:
            row += _metric_values(m)
        ps = getattr(res, "pulse_samples", None) or []
        for k in range(max_pulse):
            row.append(ps[k] if k < len(ps) else None)
        row.append(res.overall)
        ws.append(row)

    wb.save(path)
    return path
