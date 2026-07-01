# -*- coding: utf-8 -*-
"""
datalog_comparator.py — Datalog 快照 vs 实时 Modbus 比对

支持两种文件格式：
  CSV  → 两段式（范围检查 + 数值比对），容差 max(±0.05, ±5%)
  JSON → 三段式（范围检查 + 单位检查 + 数值比对），容差 max(±0.05, ±5%)
"""
from __future__ import annotations

import asyncio
import csv
import html
import json
import logging
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent.parent.parent.parent  # AcuHMI-1-7/
sys.path.insert(0, str(_PROJECT_ROOT))

import config
from template_reader import find_template_file, get_datalog_params, natural_sort_key
from utils.modbus_reader import ModbusResult, read_device_params

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 设备文件名匹配关键字
# ─────────────────────────────────────────────────────────────────────────────

_DEVICE_FILE_KEYWORDS: dict[str, str] = {
    "acurev4100": "acurev4100",
    "acurev2100": "acurev2100",
    "acuvimiiw":  "iiw",
    "acuvimiir":  "iir",
    "acuvim3":    "acuvim3",
    "pxm350":     "pxm350",
    "acuiom01":   "acuiom01",
    "acuiom02":   "acuiom02",
}


# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class DatalogScopeReport:
    template_count:   int
    file_count:       int
    matched_keys:     list[str]
    missing_from_file: list[str]
    extra_in_file:    list[str]

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_file and not self.extra_in_file


@dataclass
class DatalogUnitResult:
    param_key:  str
    file_unit:  str
    tmpl_unit:  str
    status:     str = ""


@dataclass
class DatalogCompareResult:
    param_key:    str
    file_value:   Optional[float] = None
    modbus_value: Optional[float] = None
    file_error:   str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    tol_basis:    str = ""
    status:       str = ""

    @property
    def ok(self) -> bool:
        return self.status == "PASS"


# ─────────────────────────────────────────────────────────────────────────────
# 文件查找
# ─────────────────────────────────────────────────────────────────────────────

def find_datalog_file(data_dir: str, device_key: str) -> str:
    keyword = _DEVICE_FILE_KEYWORDS.get(device_key.lower(), device_key.lower())
    matched: list[Path] = []
    for ext in ("*.json", "*.csv"):
        matched.extend(
            p for p in Path(data_dir).glob(ext)
            if keyword.lower() in p.name.lower()
        )
    if not matched:
        raise FileNotFoundError(
            f"在 {data_dir} 中未找到包含 '{keyword}' 的 Datalog 文件（json/csv）"
        )
    json_files = [p for p in matched if p.suffix.lower() == ".json"]
    chosen = sorted(json_files or matched, key=lambda p: p.stat().st_mtime)
    return str(chosen[-1])


# ─────────────────────────────────────────────────────────────────────────────
# CSV 加载
# ─────────────────────────────────────────────────────────────────────────────

def load_datalog_csv(
    csv_path: str,
    row_index: int = -1,
) -> tuple[dict[str, Optional[float]], str, list[str]]:
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        headers = next(reader)
        data_rows = list(reader)

    if not data_rows:
        raise ValueError(f"CSV 文件无数据行：{csv_path}")

    row = data_rows[-1] if row_index == -1 else data_rows[row_index]
    timestamp_str = str(row[0]) if row else "未知"
    file_columns  = [h.strip() for h in headers[1:] if h.strip()]

    value_map: dict[str, Optional[float]] = {}
    for h, v in zip(headers[1:], row[1:]):
        h = h.strip()
        if not h:
            continue
        try:
            value_map[h] = float(v) if v not in (None, "") else None
        except (ValueError, TypeError):
            value_map[h] = None

    return value_map, timestamp_str, file_columns


# ─────────────────────────────────────────────────────────────────────────────
# JSON 加载
# ─────────────────────────────────────────────────────────────────────────────

def load_datalog_json(
    json_path: str,
    row_index: int = -1,
) -> tuple[dict[str, Optional[float]], dict[str, str], str, list[str]]:
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    timestamps = data.get("timestamp", [])
    if timestamps:
        idx = -1 if row_index == -1 else row_index
        ts_val = timestamps[idx]
        timestamp_str = str(ts_val)
        val_idx = len(timestamps) - 1 if row_index == -1 else row_index
    else:
        timestamp_str = "未知"
        val_idx = 0

    readings = data["device"]["readings"]

    value_map: dict[str, Optional[float]] = {}
    unit_map:  dict[str, str] = {}
    file_columns: list[str] = []

    for r in readings:
        pkey = str(r["param"]).strip()
        unit = str(r.get("unit", "")).strip()
        vals = r.get("value", [])
        file_columns.append(pkey)
        unit_map[pkey] = unit
        if vals and val_idx < len(vals):
            raw = vals[val_idx]
            try:
                value_map[pkey] = float(raw)
            except (ValueError, TypeError):
                value_map[pkey] = None
        else:
            value_map[pkey] = None

    return value_map, unit_map, timestamp_str, file_columns


# ─────────────────────────────────────────────────────────────────────────────
# 单位比对（JSON 专用）
# ─────────────────────────────────────────────────────────────────────────────

def _norm_unit(u: str) -> str:
    return u.strip().lower().replace("°", "deg").replace("℃", "degc").replace("deg c", "degc")


def _check_unit(param_key: str, file_unit: str, tmpl_unit: str) -> DatalogUnitResult:
    r = DatalogUnitResult(param_key=param_key, file_unit=file_unit, tmpl_unit=tmpl_unit)
    if not tmpl_unit and not file_unit:
        r.status = "SKIP"
        return r
    r.status = "PASS" if _norm_unit(file_unit) == _norm_unit(tmpl_unit) else "FAIL"
    return r


# ─────────────────────────────────────────────────────────────────────────────
# 数值比对
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one(
    param_key: str,
    file_value: Optional[float],
    mr: Optional[ModbusResult],
) -> DatalogCompareResult:
    cr = DatalogCompareResult(param_key=param_key)

    if file_value is None:
        cr.file_error = "文件无数据"
    if mr is None or not mr.ok:
        cr.modbus_error = (mr.error if mr else "未读取到")

    if cr.file_error and cr.modbus_error:
        cr.status = "BOTH_ERR"; return cr
    if cr.file_error:
        cr.status = "FILE_ERR"; return cr
    if cr.modbus_error:
        cr.status = "MODBUS_ERR"; return cr

    cv, mv = file_value, mr.value
    cr.file_value   = cv
    cr.modbus_value = mv

    diff = abs(cv - mv)
    ref  = max(abs(cv), abs(mv))
    cr.diff_abs = diff
    if ref <= 1e-9:
        cr.diff_pct = 0.0
        cr.tol_basis = "zero"
        cr.status = "PASS"
    else:
        cr.diff_pct = diff / ref * 100
        rel_tol = ref * config.DATALOG_TOLERANCE_PERCENT / 100
        abs_tol = config.DATALOG_TOLERANCE_ABSOLUTE
        if rel_tol >= abs_tol:
            cr.tol_basis = f"相对 ≤{config.DATALOG_TOLERANCE_PERCENT}%"
            cr.status = "PASS" if diff <= rel_tol else "FAIL"
        else:
            cr.tol_basis = f"绝对 ≤{abs_tol}"
            cr.status = "PASS" if diff <= abs_tol else "FAIL"
    return cr


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

async def run_datalog_comparison(
    file_path: str,
    row_index: int = -1,
    param_keys: Optional[list[str]] = None,
) -> tuple[DatalogScopeReport, list[DatalogUnitResult], list[DatalogCompareResult], str]:
    t0 = time.time()
    ext = Path(file_path).suffix.lower()

    unit_map: dict[str, str] = {}
    if ext == ".json":
        value_map, unit_map, timestamp_str, file_columns = load_datalog_json(file_path, row_index)
    else:
        value_map, timestamp_str, file_columns = load_datalog_csv(file_path, row_index)

    file_keys = set(file_columns)

    tmpl_unit_map: dict[str, str] = {}
    try:
        tmpl_path   = find_template_file(config.TEMPLATE_DIR, config.DEVICE_NAME)
        tmpl_params = get_datalog_params(tmpl_path)
        tmpl_keys   = {p.param_key for p in tmpl_params}
        tmpl_unit_map = {p.param_key: p.unit for p in tmpl_params}
    except Exception as exc:
        log.warning("无法加载模板，范围检查将跳过：%s", exc)
        tmpl_keys = set()

    matched_keys_set = tmpl_keys & file_keys if tmpl_keys else file_keys
    scope_report = DatalogScopeReport(
        template_count    = len(tmpl_keys),
        file_count        = len(file_keys),
        matched_keys      = sorted(matched_keys_set, key=natural_sort_key),
        missing_from_file = sorted(tmpl_keys - file_keys, key=natural_sort_key),
        extra_in_file     = sorted(file_keys - tmpl_keys, key=natural_sort_key),
    )

    unit_results: list[DatalogUnitResult] = []
    if ext == ".json" and tmpl_keys:
        for key in sorted(matched_keys_set, key=natural_sort_key):
            if param_keys is None or key in param_keys:
                unit_results.append(
                    _check_unit(key, unit_map.get(key, ""), tmpl_unit_map.get(key, ""))
                )

    compare_keys = sorted(matched_keys_set, key=natural_sort_key)
    if param_keys:
        compare_keys = [k for k in compare_keys if k in param_keys]

    log.info("读取实时 Modbus 寄存器（%d 项）…", len(compare_keys))
    # 优先使用 verify_files 设置的覆盖键（处理多台同型设备场景）
    _modbus_key = getattr(config, "MODBUS_OVERRIDE_KEY", None) or config.DEVICE_NAME
    host_info = config.MODBUS_DEVICE_MAP.get(_modbus_key)
    if isinstance(host_info, dict):
        host = host_info["ip"]
        port = host_info["port"]
        unit = host_info["unit"]
    elif host_info:
        host, port, unit = host_info  # 兼容旧 tuple 格式
    else:
        host = getattr(config, "MODBUS_HOST", "127.0.0.1")
        port = getattr(config, "MODBUS_PORT", 502)
        unit = getattr(config, "MODBUS_UNIT", 1)
    modbus_map = await read_device_params(config.DEVICE_NAME, compare_keys, host, port, unit)

    results: list[DatalogCompareResult] = []
    for key in compare_keys:
        mr = modbus_map.get(key)
        results.append(_compare_one(key, value_map.get(key), mr))

    log.info("比对完成，耗时 %.1f 秒，共 %d 项", time.time() - t0, len(results))
    return scope_report, unit_results, results, timestamp_str


# ─────────────────────────────────────────────────────────────────────────────
# 摘要
# ─────────────────────────────────────────────────────────────────────────────

def summary(results: list[DatalogCompareResult]) -> dict:
    by_status: dict[str, int] = {}
    for r in results:
        by_status[r.status] = by_status.get(r.status, 0) + 1
    total  = len(results)
    passed = by_status.get("PASS", 0)
    failed = by_status.get("FAIL", 0)
    errors = total - passed - failed
    return {
        "total": total, "pass": passed, "fail": failed, "error": errors,
        "pass_rate": f"{passed/total*100:.1f}%" if total else "N/A",
        "by_status": by_status,
        "worst_fails": sorted(
            [r for r in results if r.status == "FAIL"],
            key=lambda r: r.diff_pct or 0, reverse=True,
        )[:10],
    }


def print_summary(
    scope: DatalogScopeReport,
    unit_results: list[DatalogUnitResult],
    results: list[DatalogCompareResult],
) -> None:
    s = summary(results)
    print("\n" + "=" * 70)
    print("  Datalog 快照 vs 实时 Modbus 比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"  【参数范围】模板={scope.template_count}  文件={scope.file_count}  "
          f"匹配={len(scope.matched_keys)}  "
          f"缺失={len(scope.missing_from_file)}  多余={len(scope.extra_in_file)}")
    if unit_results:
        uf = [r for r in unit_results if r.status == "FAIL"]
        print(f"  【单位检查】{len(unit_results)} 项，FAIL={len(uf)}")
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    print("=" * 70)


_DEVICE_MAP: dict[str, tuple[str, str]] = {
    "acurev4100": ("AcuRev4100", "devices.acurev4100"),
    "acurev2100": ("AcuRev2100", "devices.acurev2100"),
    "acuvimiiw":  ("AcuvimIIW",  "devices.acuvimiiw"),
    "acuvimiir":  ("AcuvimIIR",  "devices.acuvimiir"),
    "acuvim3":    ("AcuVIM3",    "devices.acuvim3"),
    "pxm350":     ("PXM350",     "devices.pxm350"),
    "acuiom01":   ("AcuIOM01",   "devices.acuiom01"),
    "acuiom02":   ("AcuIOM02",   "devices.acuiom02"),
}
