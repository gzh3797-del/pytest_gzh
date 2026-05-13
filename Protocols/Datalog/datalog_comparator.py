# -*- coding: utf-8 -*-
"""
datalog_comparator.py — Datalog 快照 vs 实时 Modbus 比对

支持两种文件格式：
  CSV  → 两段式（范围检查 + 数值比对），容差 ±5% / ±0.05
  JSON → 三段式（范围检查 + 单位检查 + 数值比对），容差 ±5% / ±0.05

用法：
  python Protocols/Datalog/datalog_comparator.py --device acurev4100
  python Protocols/Datalog/datalog_comparator.py --device acurev4100 --file <路径>
  python Protocols/Datalog/datalog_comparator.py --device acurev4100 --row 1
  python Protocols/Datalog/datalog_comparator.py --device acurev4100 --keys FREQ_Hz VLN_a_V
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

sys.path.insert(0, str(Path(__file__).parent.parent))

import config
from modbus_reader import ModbusResult, get_reader
from template_reader import find_template_file, get_datalog_params

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
    missing_from_file: list[str]   # 模板有但文件无
    extra_in_file:    list[str]    # 文件有但模板无

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_file and not self.extra_in_file


@dataclass
class DatalogUnitResult:
    param_key:  str
    file_unit:  str    # 文件中的单位（JSON 字段）
    tmpl_unit:  str    # 模板单位
    status:     str    # PASS | FAIL | SKIP


@dataclass
class DatalogCompareResult:
    param_key:    str
    file_value:   Optional[float] = None
    modbus_value: Optional[float] = None
    file_error:   str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | FILE_ERR | MODBUS_ERR | BOTH_ERR

    @property
    def ok(self) -> bool:
        return self.status == "PASS"


# ─────────────────────────────────────────────────────────────────────────────
# 文件查找
# ─────────────────────────────────────────────────────────────────────────────

def find_datalog_file(data_dir: str, device_key: str) -> str:
    """
    在 data_dir 中查找与设备匹配的 Datalog 文件（JSON 优先，次选 CSV；取修改时间最新）。
    """
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
    # JSON 优先：若有 JSON 文件则只取 JSON，否则取 CSV
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
    """
    读取 Datalog CSV 一行数据。

    Returns:
        (value_map, timestamp_str, file_columns)
    """
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        headers = next(reader)
        data_rows = list(reader)

    if not data_rows:
        raise ValueError(f"CSV 文件无数据行：{csv_path}")

    row = data_rows[-1] if row_index == -1 else data_rows[row_index]
    log.info("CSV 使用第 %d 数据行，时间戳：%s",
             len(data_rows) if row_index == -1 else row_index + 1, row[0])

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
    """
    读取 Datalog JSON 文件。

    JSON 结构：
      { "timestamp": [...], "device": { "readings": [{"param","unit","value":[...]}, ...] } }

    Returns:
        (value_map, unit_map, timestamp_str, file_columns)
        value_map:   dict[param_key → float | None]
        unit_map:    dict[param_key → unit_str]
        file_columns: 全部 param 列表（不含 TimeTag）
    """
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
    log.info("JSON 设备：%s  readings=%d  时间戳索引=%d（%s）",
             data["device"].get("name", ""), len(readings), val_idx, timestamp_str)

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
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0

    tol = max(
        config.DATALOG_TOLERANCE_ABSOLUTE,
        ref * config.DATALOG_TOLERANCE_PERCENT / 100,
    )
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

async def run_datalog_comparison(
    file_path: str,
    row_index: int = -1,
    param_keys: Optional[list[str]] = None,
) -> tuple[DatalogScopeReport, list[DatalogUnitResult], list[DatalogCompareResult], str]:
    """
    执行 Datalog 比对。CSV 返回空 unit_results；JSON 返回单位检查结果。
    """
    t0 = time.time()
    ext = Path(file_path).suffix.lower()

    # ── 加载文件 ───────────────────────────────────────────────────────────────
    unit_map: dict[str, str] = {}
    if ext == ".json":
        log.info("加载 JSON：%s", file_path)
        value_map, unit_map, timestamp_str, file_columns = load_datalog_json(file_path, row_index)
    else:
        log.info("加载 CSV：%s", file_path)
        value_map, timestamp_str, file_columns = load_datalog_csv(file_path, row_index)

    file_keys = set(file_columns)

    # ── 模板范围 ───────────────────────────────────────────────────────────────
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
        matched_keys      = sorted(matched_keys_set),
        missing_from_file = sorted(tmpl_keys - file_keys),
        extra_in_file     = sorted(file_keys - tmpl_keys),
    )
    log.info("范围检查：模板=%d  文件=%d  匹配=%d  缺失=%d  多余=%d",
             len(tmpl_keys), len(file_keys), len(matched_keys_set),
             len(scope_report.missing_from_file), len(scope_report.extra_in_file))

    # ── 单位检查（仅 JSON） ────────────────────────────────────────────────────
    unit_results: list[DatalogUnitResult] = []
    if ext == ".json" and tmpl_keys:
        for key in sorted(matched_keys_set):
            if param_keys is None or key in param_keys:
                unit_results.append(
                    _check_unit(key, unit_map.get(key, ""), tmpl_unit_map.get(key, ""))
                )
        unit_fail = sum(1 for r in unit_results if r.status == "FAIL")
        log.info("单位检查：%d 项，FAIL=%d", len(unit_results), unit_fail)

    # ── 实时 Modbus ────────────────────────────────────────────────────────────
    compare_keys = sorted(matched_keys_set)
    if param_keys:
        compare_keys = [k for k in compare_keys if k in param_keys]

    log.info("读取实时 Modbus 寄存器（%d 项）…", len(compare_keys))
    async with get_reader() as modbus:
        modbus_results = await modbus.read_params(compare_keys)
    modbus_map = {r.param_key: r for r in modbus_results}

    # ── 数值比对 ───────────────────────────────────────────────────────────────
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
    if scope.missing_from_file:
        print(f"  缺失参数（前10）: {scope.missing_from_file[:10]}")
    if scope.extra_in_file:
        print(f"  多余参数（前10）: {scope.extra_in_file[:10]}")
    if unit_results:
        uf = [r for r in unit_results if r.status == "FAIL"]
        print(f"  【单位检查】{len(unit_results)} 项，FAIL={len(uf)}")
        for r in uf[:5]:
            print(f"    {r.param_key}: 文件={r.file_unit}  模板={r.tmpl_unit}")
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    if s["worst_fails"]:
        print("\n  差异最大的失败参数（Top 10）：")
        for r in s["worst_fails"]:
            print(f"    {r.param_key:40s}  文件={r.file_value:.6g}  "
                  f"Modbus={r.modbus_value:.6g}  Δ%={r.diff_pct:.2f}%")
    print("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "PASS":       "#d4edda",
    "FAIL":       "#f8d7da",
    "FILE_ERR":   "#fff3cd",
    "MODBUS_ERR": "#fff3cd",
    "BOTH_ERR":   "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS":       "通过",
    "FAIL":       "失败",
    "FILE_ERR":   "文件异常",
    "MODBUS_ERR": "Modbus异常",
    "BOTH_ERR":   "双路异常",
}


def _fmt(v: Optional[float], digits: int = 6) -> str:
    return "—" if v is None else f"{v:.{digits}g}"


def generate_html_report(
    scope: DatalogScopeReport,
    unit_results: list[DatalogUnitResult],
    results: list[DatalogCompareResult],
    timestamp_str: str = "",
    file_path: str = "",
    output_path: Optional[str] = None,
) -> str:
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"datalog_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_basename = Path(file_path).name if file_path else "—"
    file_type    = "JSON" if file_path.lower().endswith(".json") else "CSV"

    # ── Section 1: 范围检查 ────────────────────────────────────────────────────
    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    scope_color = "#d4edda" if scope.scope_ok else "#f8d7da"
    scope_label = "一致" if scope.scope_ok else "不一致"
    missing_html = (
        f'<h3 style="color:#721c24;margin-top:12px">模板有但文件缺失（{len(scope.missing_from_file)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.missing_from_file, "#f8d7da")}</tbody></table>'
    ) if scope.missing_from_file else ""
    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">文件有但模板未列入（{len(scope.extra_in_file)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.extra_in_file, "#fff3cd")}</tbody></table>'
    ) if scope.extra_in_file else ""
    scope_badge = (
        (f'<span class="badge err-badge">缺失 {len(scope.missing_from_file)}</span>' if scope.missing_from_file else "") +
        (f'<span class="badge warn-badge">多余 {len(scope.extra_in_file)}</span>' if scope.extra_in_file else "") +
        ('<span class="badge ok-badge">一致</span>' if scope.scope_ok else "")
    )
    scope_html = f"""
<details open class="section">
<summary>一、参数范围检查
  <span class="sum-info">模板 {scope.template_count} / {file_type} {scope.file_count} / 匹配 {len(scope.matched_keys)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板 DataLog 参数</div></div>
  <div class="card total"><div class="num">{scope.file_count}</div><div class="lbl">{file_type} 实际列数</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_file)}</div><div class="lbl">模板有/文件缺</div></div>
  <div class="card err">  <div class="num">{len(scope.extra_in_file)}</div><div class="lbl">文件多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 单位检查（仅 JSON） ────────────────────────────────────────
    unit_html = ""
    val_section_num = "二"
    if unit_results:
        val_section_num = "三"
        unit_fail = [r for r in unit_results if r.status == "FAIL"]
        unit_badge = (
            (f'<span class="badge err-badge">不匹配 {len(unit_fail)}</span>' if unit_fail else "") +
            ('<span class="badge ok-badge">全部一致</span>' if not unit_fail else "")
        )
        unit_rows = "".join(
            f'<tr style="background:{"#f8d7da" if r.status == "FAIL" else "#d4edda" if r.status == "PASS" else "#f8f9fa"}">'
            f'<td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(r.param_key)}</td>'
            f'<td class="val">{html.escape(r.file_unit)}</td>'
            f'<td class="val">{html.escape(r.tmpl_unit)}</td>'
            f'<td class="stat">{"失败" if r.status == "FAIL" else "通过" if r.status == "PASS" else "跳过"}</td>'
            f'</tr>'
            for i, r in enumerate(unit_results)
        )
        unit_html = f"""
<details class="section">
<summary>二、单位检查
  <span class="sum-info">{len(unit_results)} 项 · FAIL={len(unit_fail)}</span>
  {unit_badge}
</summary>
<div class="section-body">
<table>
<colgroup><col class="c-idx"><col class="c-key"><col class="c-val"><col class="c-val"><col class="c-stat"></colgroup>
<thead><tr><th>#</th><th>param_key</th><th>文件单位</th><th>模板单位</th><th>结果</th></tr></thead>
<tbody>{unit_rows}</tbody>
</table>
</div>
</details>"""

    # ── Section 2/3: 数值比对 ──────────────────────────────────────────────────
    rows_html = []
    for i, r in enumerate(results):
        bg   = _STATUS_COLOR.get(r.status, "#ffffff")
        stat = _STATUS_LABEL.get(r.status, r.status)
        err_hint = ""
        if r.file_error:
            err_hint += f"文件: {html.escape(r.file_error)}"
        if r.modbus_error:
            if err_hint: err_hint += "<br>"
            err_hint += f"Modbus: {html.escape(r.modbus_error)}"
        rows_html.append(f"""
        <tr style="background:{bg}">
          <td class="num">{i+1}</td>
          <td class="key">{html.escape(r.param_key)}</td>
          <td class="val">{_fmt(r.file_value)}</td>
          <td class="val">{_fmt(r.modbus_value)}</td>
          <td class="val">{_fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"}</td>
          <td class="val">{f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"}</td>
          <td class="stat">{stat}</td>
          <td class="err">{err_hint}</td>
        </tr>""")

    val_badge = (
        (f'<span class="badge err-badge">失败 {s["fail"]}</span>' if s["fail"] else "") +
        (f'<span class="badge warn-badge">异常 {s["error"]}</span>' if s["error"] else "") +
        ('<span class="badge ok-badge">全部通过</span>' if not s["fail"] and not s["error"] else "")
    )
    val_html = f"""
<details open class="section">
<summary>{val_section_num}、数值比对（Datalog 快照 vs 实时 Modbus）
  <span class="sum-info">共 {s['total']} 项 · 通过率 {s['pass_rate']}</span>
  {val_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{s['total']}</div><div class="lbl">总参数</div></div>
  <div class="card pass"> <div class="num">{s['pass']}</div> <div class="lbl">通过 PASS ({s['pass_rate']})</div></div>
  <div class="card fail"> <div class="num">{s['fail']}</div> <div class="lbl">失败 FAIL</div></div>
  <div class="card err">  <div class="num">{s['error']}</div><div class="lbl">异常 ERROR</div></div>
</div>
<table>
<colgroup>
  <col class="c-idx"><col class="c-key"><col class="c-val"><col class="c-val">
  <col class="c-diff"><col class="c-pct"><col class="c-stat"><col class="c-err">
</colgroup>
<thead>
  <tr>
    <th>#</th><th>参数名 (param_key)</th><th>{file_type} 快照值</th><th>Modbus 实时值</th>
    <th>绝对差值</th><th>相对差值</th><th>结果</th><th>错误信息</th>
  </tr>
</thead>
<tbody>{"".join(rows_html)}</tbody>
</table>
</div>
</details>"""

    modbus_info = (
        f"RTU {config.MODBUS_RTU_PORT}" if getattr(config, "MODBUS_MODE", "tcp") == "rtu"
        else f"{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT}"
    )
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Datalog vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>
  body {{ font-family: "Microsoft YaHei", Arial, sans-serif; font-size: 13px;
          margin: 20px; background: #f5f5f5; color: #333; }}
  h1   {{ font-size: 20px; margin-bottom: 4px; }}
  .device-name {{ font-size: 15px; font-weight: bold; color: #0056b3; margin-bottom: 4px; }}
  .meta {{ color: #666; font-size: 12px; margin-bottom: 16px; }}
  .cards {{ display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }}
  .card {{ background: #fff; border-radius: 6px; padding: 14px 20px;
           box-shadow: 0 1px 4px rgba(0,0,0,.1); min-width: 110px; text-align: center; }}
  .card .num {{ font-size: 28px; font-weight: bold; }}
  .card .lbl {{ font-size: 11px; color: #888; margin-top: 2px; }}
  .card.pass  .num {{ color: #28a745; }}
  .card.fail  .num {{ color: #dc3545; }}
  .card.err   .num {{ color: #ffc107; }}
  .card.total .num {{ color: #007bff; }}
  table {{ border-collapse: collapse; width: 100%; background: #fff;
           box-shadow: 0 1px 4px rgba(0,0,0,.1); border-radius: 6px;
           overflow: hidden; table-layout: fixed; margin-bottom: 24px; }}
  colgroup col.c-idx  {{ width: 48px; }}
  colgroup col.c-key  {{ width: 280px; }}
  colgroup col.c-val  {{ width: 110px; }}
  colgroup col.c-diff {{ width: 90px; }}
  colgroup col.c-pct  {{ width: 80px; }}
  colgroup col.c-stat {{ width: 90px; }}
  colgroup col.c-err  {{ width: auto; }}
  thead tr {{ background: #343a40; color: #fff; }}
  th   {{ padding: 9px 8px; text-align: center; font-size: 12px; white-space: nowrap; }}
  td   {{ padding: 5px 8px; border-bottom: 1px solid #eee; vertical-align: middle; }}
  td.num  {{ text-align: right; color: #999; font-size: 11px; }}
  td.key  {{ font-family: monospace; font-size: 12px; word-break: break-all; }}
  td.val  {{ text-align: right; font-family: monospace; font-size: 12px; }}
  td.stat {{ text-align: center; font-weight: bold; font-size: 12px; }}
  td.err  {{ font-size: 11px; color: #666; word-break: break-all; }}
  tr:hover td {{ filter: brightness(0.96); }}
  thead th {{ position: sticky; top: 0; z-index: 1; }}
  details.section {{ background: #fff; border-radius: 8px; margin-bottom: 20px;
                    box-shadow: 0 1px 4px rgba(0,0,0,.1); overflow: hidden; }}
  details.section > summary {{ list-style: none; cursor: pointer; padding: 12px 16px;
    background: #f0f4f8; border-left: 4px solid #0056b3;
    font-size: 15px; font-weight: bold; color: #0056b3;
    display: flex; align-items: center; gap: 8px; user-select: none; }}
  details.section > summary::-webkit-details-marker {{ display: none; }}
  details.section > summary::before {{ content: "▶"; font-size: 10px; margin-right: 4px; }}
  details[open].section > summary::before {{ content: "▼"; }}
  .sum-info {{ font-size: 12px; font-weight: normal; color: #666; margin-left: 4px; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px;
           font-size: 11px; font-weight: bold; margin-left: 4px; }}
  .ok-badge   {{ background: #d4edda; color: #155724; }}
  .err-badge  {{ background: #f8d7da; color: #721c24; }}
  .warn-badge {{ background: #fff3cd; color: #856404; }}
  .section-body {{ padding: 16px; }}
</style>
</head>
<body>
<h1>Datalog 快照 vs 实时 Modbus 比对报告</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  数据文件：{html.escape(file_basename)}（{file_type}） &nbsp;|&nbsp;
  快照时间戳：{html.escape(timestamp_str)} &nbsp;|&nbsp;
  Modbus：{modbus_info} &nbsp;|&nbsp;
  容差：±{config.DATALOG_TOLERANCE_PERCENT}% / ±{config.DATALOG_TOLERANCE_ABSOLUTE}
</div>
{scope_html}
{unit_html}
{val_html}
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("HTML 报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 设备映射
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────────────────

async def _main(
    file_path: str,
    row_index: int,
    param_keys: Optional[list[str]],
) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    scope, unit_results, results, timestamp_str = await run_datalog_comparison(
        file_path, row_index, param_keys,
    )
    print_summary(scope, unit_results, results)
    report_path = generate_html_report(
        scope, unit_results, results,
        timestamp_str=timestamp_str,
        file_path=file_path,
    )
    print(f"\n  HTML 报告：{report_path}\n")


if __name__ == "__main__":
    # ── --device ──────────────────────────────────────────────────────────────
    _dev_value: Optional[str] = None
    if "--device" in sys.argv:
        idx = sys.argv.index("--device")
        _dev_value = sys.argv[idx + 1].lower()
        if _dev_value not in _DEVICE_MAP:
            print(f"[ERROR] 未知设备 '{_dev_value}'，可选：{list(_DEVICE_MAP)}")
            sys.exit(1)
        config.DEVICE_NAME, config.DEVICE_MODULE = _DEVICE_MAP[_dev_value]

    if config.DEVICE_NAME in config.MODBUS_DEVICE_MAP:
        config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
            config.MODBUS_DEVICE_MAP[config.DEVICE_NAME]

    # ── --file ────────────────────────────────────────────────────────────────
    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        _file_path = sys.argv[idx + 1]
    else:
        _file_path = find_datalog_file(
            config.DATALOG_DATA_DIR, _dev_value or config.DEVICE_NAME.lower()
        )
        print(f"[INFO] 自动选取文件：{Path(_file_path).name}")

    # ── --row ─────────────────────────────────────────────────────────────────
    _row_index = -1
    if "--row" in sys.argv:
        idx = sys.argv.index("--row")
        _row_index = int(sys.argv[idx + 1]) - 1

    # ── --keys ────────────────────────────────────────────────────────────────
    _keys: Optional[list[str]] = None
    if "--keys" in sys.argv:
        idx = sys.argv.index("--keys")
        _keys = sys.argv[idx + 1:]

    asyncio.run(_main(_file_path, _row_index, _keys))
