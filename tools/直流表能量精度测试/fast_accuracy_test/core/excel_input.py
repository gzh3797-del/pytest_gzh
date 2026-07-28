from dataclasses import dataclass
import openpyxl


@dataclass
class Case:
    test_case: str
    voltage: float
    current_1: float
    current_2: float
    wait_h: float
    voltage_accuracy: float
    current_accuracy: float
    power_accuracy: float
    sample_cnt: int
    sample_interval: float
    # 可选：脉冲检测(TE3100)。两者都填才测能量脉冲误差；任一为空则不测，按原流程跑。
    pulse_const: float = None       # 能量脉冲常数 (imp/kWh)
    pulse_error_acc: float = None   # 能量脉冲误差阈值(比例值，同其它精度列 ×100 得%)


# 必填列
_COLS = {
    "test_case": "test_case", "voltage": "Voltage",
    "current_1": "Current_1", "current_2": "Current_2", "wait_h": "等待时间(h)",
    "voltage_accuracy": "voltage_accuracy", "current_accuracy": "current_accuracy",
    "power_accuracy": "power_accuracy", "sample_cnt": "采样次数", "sample_interval": "采样间隔",
}
# 可选列（不存在或留空都按 None 处理，向后兼容旧表）
_OPTIONAL_COLS = {
    "pulse_const": "能量脉冲常数",
    "pulse_error_acc": "能量脉冲误差",
}


def _opt_float(row, header, col):
    """读可选数值列：列不存在或单元格为空 → None。"""
    if col not in header:
        return None
    v = row[header[col]]
    if v in (None, ""):
        return None
    return float(v)


def read_cases(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = {name: i for i, name in enumerate(rows[0])}
    missing = [col for col in _COLS.values() if col not in header]
    if missing:
        raise ValueError(f"{path}: 缺少列 {missing}")
    idx = {field: header[col] for field, col in _COLS.items()}

    cases = []
    for row in rows[1:]:
        if row[idx["test_case"]] in (None, ""):
            continue
        cases.append(Case(
            test_case=str(row[idx["test_case"]]),
            voltage=float(row[idx["voltage"]]),
            current_1=float(row[idx["current_1"]]),
            current_2=float(row[idx["current_2"]]),
            wait_h=float(row[idx["wait_h"]]),
            voltage_accuracy=float(row[idx["voltage_accuracy"]]),
            current_accuracy=float(row[idx["current_accuracy"]]),
            power_accuracy=float(row[idx["power_accuracy"]]),
            sample_cnt=int(row[idx["sample_cnt"]]),
            sample_interval=float(row[idx["sample_interval"]]),
            pulse_const=_opt_float(row, header, _OPTIONAL_COLS["pulse_const"]),
            pulse_error_acc=_opt_float(row, header, _OPTIONAL_COLS["pulse_error_acc"]),
        ))
    return cases
