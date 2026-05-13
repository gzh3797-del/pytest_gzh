# -*- coding: utf-8 -*-
"""
cloud_comparator.py — AcuCloud 历史快照 vs 实时 Modbus 数值比对

流程：
  1. 读取 AcuCloud 导出的 xlsx 文件（取最新行或指定行）
  2. 使用设备模块的 build_cloud_col_map() 构建"列标题 → param_key"映射
  3. 与实时 Modbus 寄存器读取值进行比对
  4. 输出控制台摘要 + HTML 报告

容差（config.CLOUD_TOLERANCE_*）比 BACnet-Modbus 对比大，以补偿时序差异。

用法：
  python AcuCloud/cloud_comparator.py                          # 自动选取最新 xlsx，最新数据行
  python AcuCloud/cloud_comparator.py --file <xlsx路径>        # 指定文件
  python AcuCloud/cloud_comparator.py --device acurev4100      # 指定设备（默认 config.DEVICE_NAME）
  python AcuCloud/cloud_comparator.py --row 1                  # 指定数据行（1=第一行，-1=最新）
  python AcuCloud/cloud_comparator.py --keys FREQ_Hz VLN_a_V   # 只比对指定参数
"""

from __future__ import annotations

import asyncio
import html
import importlib
import logging
import sys
import time
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl

sys.path.insert(0, str(Path(__file__).parent.parent))

import config
from modbus_reader import ModbusReader, ModbusResult, get_reader
from template_reader import find_template_file, get_cloud_params

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 比对结果数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class CloudScopeReport:
    """模板 AcuCloud 参数范围 vs 设备模块 col_map 覆盖情况。"""
    template_count:      int
    colmap_count:        int
    matched_keys:        list[str]   # 模板与 col_map 均有
    missing_from_colmap: list[str]   # 模板有但 col_map 未映射
    extra_in_colmap:     list[str]   # col_map 有但模板未列入

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_colmap and not self.extra_in_colmap


@dataclass
class CloudCompareResult:
    """单个参数的 AcuCloud 快照 vs 实时 Modbus 比对结果。"""
    param_key:    str
    cloud_value:  Optional[float] = None   # xlsx 快照值
    modbus_value: Optional[float] = None   # 实时 Modbus 值
    cloud_error:  str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | CLOUD_ERR | MODBUS_ERR | BOTH_ERR

    @property
    def ok(self) -> bool:
        return self.status == "PASS"


# ─────────────────────────────────────────────────────────────────────────────
# xlsx 加载
# ─────────────────────────────────────────────────────────────────────────────

def find_latest_xlsx(data_dir: str, device_hint: str = "") -> str:
    """返回 data_dir 中与设备名匹配的 xlsx 文件路径（跳过 Excel 锁文件 ~$...）。

    若提供 device_hint（如 'acuvim3'），按文件名精确匹配（大小写不敏感）。
    未提供时返回修改时间最新的文件。
    """
    xlsx_files = sorted(
        (p for p in Path(data_dir).glob("*.xlsx") if not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime,
    )
    if not xlsx_files:
        raise FileNotFoundError(f"在 {data_dir} 中未找到 xlsx 文件")
    if device_hint:
        matching = [p for p in xlsx_files if p.stem.lower() == device_hint.lower()]
        if not matching:
            raise FileNotFoundError(
                f"在 {data_dir} 中未找到与设备 '{device_hint}' 匹配的 xlsx 文件"
                f"（期望文件名：{device_hint}.xlsx）"
            )
        return str(matching[-1])
    return str(xlsx_files[-1])


def load_cloud_row(
    xlsx_path: str,
    row_index: int = -1,
) -> tuple[dict[str, Optional[float]], str, str, list[tuple[str, int]]]:
    """
    从 xlsx 加载一行数据。

    Args:
        xlsx_path:  xlsx 文件路径
        row_index:  数据行下标（0=第一数据行，-1=最新行）

    Returns:
        (cloud_map, timestamp_str, sheet_name, dup_headers)
        cloud_map:   dict[header → float | None]
        dup_headers: 有值列中重复出现的列名及出现次数，[(header, count), ...]
    """
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    ws = wb.active
    sheet_name = ws.title

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        raise ValueError(f"xlsx 文件中没有数据行：{xlsx_path}")

    headers = list(rows[0])
    data_rows = rows[1:]

    if row_index == -1:
        row = data_rows[-1]
        row_label = f"第 {len(data_rows)} 数据行（最新）"
    else:
        row = data_rows[row_index]
        row_label = f"第 {row_index + 1} 数据行"

    timestamp_str = str(row[0]) if row[0] else "未知"
    log.info("使用 %s，时间戳：%s", row_label, timestamp_str)

    cloud_map: dict[str, Optional[float]] = {}
    nonempty_header_counts: Counter = Counter()

    for h, v in zip(headers[1:], row[1:]):   # 跳过第一列（时间戳）
        if h is None:
            continue
        if v is None or v == '':
            cloud_map[h] = None
        else:
            nonempty_header_counts[h] += 1
            try:
                cloud_map[h] = float(v)
            except (ValueError, TypeError):
                cloud_map[h] = None

    dup_headers = [(h, cnt) for h, cnt in nonempty_header_counts.items() if cnt > 1]
    if dup_headers:
        log.warning("xlsx 存在重复列名（有值）：%s", [h for h, _ in dup_headers])

    wb.close()
    return cloud_map, timestamp_str, sheet_name, dup_headers


# ─────────────────────────────────────────────────────────────────────────────
# 单参数比对
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one(
    param_key: str,
    cloud_value: Optional[float],
    mr: Optional[ModbusResult],
) -> CloudCompareResult:
    cr = CloudCompareResult(param_key=param_key)

    if cloud_value is None:
        cr.cloud_error = "xlsx 无数据"
    if mr is None or not mr.ok:
        cr.modbus_error = (mr.error if mr else "未读取到")

    if cr.cloud_error and cr.modbus_error:
        cr.status = "BOTH_ERR"
        return cr
    if cr.cloud_error:
        cr.status = "CLOUD_ERR"
        return cr
    if cr.modbus_error:
        cr.status = "MODBUS_ERR"
        return cr

    cv = cloud_value
    mv = mr.value
    cr.cloud_value  = cv
    cr.modbus_value = mv

    diff = abs(cv - mv)
    ref  = max(abs(cv), abs(mv))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0

    tol = max(
        config.CLOUD_TOLERANCE_ABSOLUTE,
        ref * config.CLOUD_TOLERANCE_PERCENT / 100,
    )
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

async def run_cloud_comparison(
    xlsx_path: str,
    row_index: int = -1,
    param_keys: Optional[list[str]] = None,
) -> tuple[CloudScopeReport, list[CloudCompareResult], str, str, list[tuple[str, int]]]:
    """
    执行 AcuCloud 快照 vs 实时 Modbus 比对（含模板范围检查）。

    Returns:
        (scope_report, results, timestamp_str, sheet_name, dup_headers)
    """
    t0 = time.time()

    # ── 加载 xlsx ──────────────────────────────────────────────────────────────
    log.info("加载 xlsx：%s", xlsx_path)
    cloud_map, timestamp_str, sheet_name, dup_headers = load_cloud_row(xlsx_path, row_index)

    # ── 获取设备列标题映射 ─────────────────────────────────────────────────────
    device_module = importlib.import_module(config.DEVICE_MODULE)
    if not hasattr(device_module, "build_cloud_col_map"):
        raise NotImplementedError(
            f"{config.DEVICE_MODULE} 未实现 build_cloud_col_map()"
        )
    col_map: dict[str, str] = device_module.build_cloud_col_map()
    colmap_keys = set(col_map.values())

    # ── 模板范围检查 ───────────────────────────────────────────────────────────
    try:
        tmpl_path   = find_template_file(config.TEMPLATE_DIR, config.DEVICE_NAME)
        tmpl_params = get_cloud_params(tmpl_path)
        tmpl_keys   = {p.param_key for p in tmpl_params}
    except Exception as exc:
        log.warning("无法加载模板文件，范围检查将跳过：%s", exc)
        tmpl_keys = set()

    matched_keys_set  = tmpl_keys & colmap_keys if tmpl_keys else colmap_keys
    missing_from_colmap = sorted(tmpl_keys - colmap_keys)
    extra_in_colmap     = sorted(colmap_keys - tmpl_keys)
    scope_report = CloudScopeReport(
        template_count      = len(tmpl_keys),
        colmap_count        = len(colmap_keys),
        matched_keys        = sorted(matched_keys_set),
        missing_from_colmap = missing_from_colmap,
        extra_in_colmap     = extra_in_colmap,
    )
    log.info("AcuCloud 范围检查：模板=%d  col_map=%d  匹配=%d  缺失=%d  多余=%d",
             len(tmpl_keys), len(colmap_keys), len(matched_keys_set),
             len(missing_from_colmap), len(extra_in_colmap))

    # ── 交集匹配（过滤掉 xlsx 中数值为空的列）────────────────────────────────
    matched: dict[str, float] = {}  # param_key → cloud_value（仅非空值）
    skipped_empty = 0
    for header, pkey in col_map.items():
        if header in cloud_map:
            v = cloud_map[header]
            if v is None:
                skipped_empty += 1
                continue
            if param_keys is None or pkey in param_keys:
                matched[pkey] = v

    if skipped_empty:
        log.info("已过滤空值列：%d 列被跳过", skipped_empty)

    if not matched:
        raise ValueError("xlsx 列标题与设备映射无交集，请检查设备模块或文件")

    log.info("匹配参数数量：%d", len(matched))

    # ── 实时读取 Modbus ────────────────────────────────────────────────────────
    log.info("读取实时 Modbus 寄存器…")
    async with get_reader() as modbus:
        modbus_results = await modbus.read_params(list(matched.keys()))

    modbus_map: dict[str, ModbusResult] = {r.param_key: r for r in modbus_results}

    # ── 比对 ───────────────────────────────────────────────────────────────────
    results: list[CloudCompareResult] = []
    for pkey, cv in matched.items():
        mr = modbus_map.get(pkey)
        results.append(_compare_one(pkey, cv, mr))

    elapsed = time.time() - t0
    log.info("比对完成，耗时 %.1f 秒，共 %d 项", elapsed, len(results))
    return scope_report, results, timestamp_str, sheet_name, dup_headers


# ─────────────────────────────────────────────────────────────────────────────
# 统计摘要
# ─────────────────────────────────────────────────────────────────────────────

def summary(results: list[CloudCompareResult]) -> dict:
    by_status: dict[str, int] = {}
    for r in results:
        by_status[r.status] = by_status.get(r.status, 0) + 1

    total  = len(results)
    passed = by_status.get("PASS", 0)
    failed = by_status.get("FAIL", 0)
    errors = total - passed - failed

    fail_list = sorted(
        [r for r in results if r.status == "FAIL"],
        key=lambda r: r.diff_pct or 0, reverse=True,
    )
    return {
        "total":       total,
        "pass":        passed,
        "fail":        failed,
        "error":       errors,
        "pass_rate":   f"{passed/total*100:.1f}%" if total else "N/A",
        "by_status":   by_status,
        "worst_fails": fail_list[:10],
    }


def print_summary(scope: CloudScopeReport, results: list[CloudCompareResult]) -> None:
    s = summary(results)
    print("\n" + "=" * 70)
    print("  AcuCloud 快照 vs 实时 Modbus 比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"  【参数范围】模板={scope.template_count}  col_map={scope.colmap_count}  "
          f"匹配={len(scope.matched_keys)}  "
          f"缺失={len(scope.missing_from_colmap)}  多余={len(scope.extra_in_colmap)}")
    if scope.missing_from_colmap:
        print(f"  缺失参数（前10）: {scope.missing_from_colmap[:10]}")
    if scope.extra_in_colmap:
        print(f"  多余参数（前10）: {scope.extra_in_colmap[:10]}")
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    if s["worst_fails"]:
        print("\n  差异最大的失败参数（Top 10）：")
        for r in s["worst_fails"]:
            print(f"    {r.param_key:40s}  Cloud={r.cloud_value:.6g}  "
                  f"Modbus={r.modbus_value:.6g}  Δ%={r.diff_pct:.2f}%")
    print("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "PASS":       "#d4edda",
    "FAIL":       "#f8d7da",
    "CLOUD_ERR":  "#fff3cd",
    "MODBUS_ERR": "#fff3cd",
    "BOTH_ERR":   "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS":       "通过",
    "FAIL":       "失败",
    "CLOUD_ERR":  "Cloud异常",
    "MODBUS_ERR": "Modbus异常",
    "BOTH_ERR":   "双路异常",
}


def _fmt(v: Optional[float], digits: int = 6) -> str:
    return "—" if v is None else f"{v:.{digits}g}"


def generate_html_report(
    scope: CloudScopeReport,
    results: list[CloudCompareResult],
    timestamp_str: str = "",
    sheet_name: str = "",
    xlsx_path: str = "",
    dup_headers: Optional[list[tuple[str, int]]] = None,
    output_path: Optional[str] = None,
) -> str:
    """生成两段式 HTML 比对报告（范围检查 / 数值比对），返回文件路径。"""
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"cloud_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    xlsx_basename = Path(xlsx_path).name if xlsx_path else "—"

    # ── Section 1: 范围检查 HTML ──────────────────────────────────────────────
    scope_color = "#d4edda" if scope.scope_ok else "#f8d7da"
    scope_label = "一致" if scope.scope_ok else "不一致"

    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    missing_html = (
        f'<h3 style="color:#721c24;margin-top:12px">模板有但 col_map 未映射（{len(scope.missing_from_colmap)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.missing_from_colmap, "#f8d7da")}</tbody></table>'
    ) if scope.missing_from_colmap else ""

    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">col_map 有但模板未列入（{len(scope.extra_in_colmap)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.extra_in_colmap, "#fff3cd")}</tbody></table>'
    ) if scope.extra_in_colmap else ""

    scope_badge = (f'<span class="badge err-badge">缺失 {len(scope.missing_from_colmap)}</span>'
                   if scope.missing_from_colmap else "") + \
                  (f'<span class="badge warn-badge">多余 {len(scope.extra_in_colmap)}</span>'
                   if scope.extra_in_colmap else "") + \
                  ('<span class="badge ok-badge">一致</span>'
                   if scope.scope_ok else "")
    scope_html = f"""
<details open class="section">
<summary>一、参数范围检查
  <span class="sum-info">模板 {scope.template_count} / col_map {scope.colmap_count} / 匹配 {len(scope.matched_keys)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板 AcuCloud 参数</div></div>
  <div class="card total"><div class="num">{scope.colmap_count}</div><div class="lbl">col_map 覆盖参数</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_colmap)}</div><div class="lbl">模板有/col_map缺</div></div>
  <div class="card err">  <div class="num">{len(scope.extra_in_colmap)}</div><div class="lbl">col_map多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 数值比对 HTML ──────────────────────────────────────────────
    if dup_headers:
        dup_items = "".join(
            f"<li>{html.escape(h)}（出现 {cnt} 次，比对时取最后一列的值）</li>"
            for h, cnt in dup_headers
        )
        dup_warn_html = (
            f'<div class="warn-box">'
            f'<b>⚠ xlsx 数据异常：存在重复列名（共 {len(dup_headers)} 处）</b>'
            f'<ul>{dup_items}</ul>'
            f'</div>'
        )
    else:
        dup_warn_html = ""

    rows_html = []
    for i, r in enumerate(results):
        bg   = _STATUS_COLOR.get(r.status, "#ffffff")
        stat = _STATUS_LABEL.get(r.status, r.status)

        cv_str = _fmt(r.cloud_value)
        mv_str = _fmt(r.modbus_value)
        da_str = _fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"
        dp_str = f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"

        err_hint = ""
        if r.cloud_error:
            err_hint += f"Cloud: {html.escape(r.cloud_error)}"
        if r.modbus_error:
            if err_hint:
                err_hint += "<br>"
            err_hint += f"Modbus: {html.escape(r.modbus_error)}"

        rows_html.append(f"""
        <tr style="background:{bg}">
          <td class="num">{i+1}</td>
          <td class="key">{html.escape(r.param_key)}</td>
          <td class="val">{cv_str}</td>
          <td class="val">{mv_str}</td>
          <td class="val">{da_str}</td>
          <td class="val">{dp_str}</td>
          <td class="stat">{stat}</td>
          <td class="err">{err_hint}</td>
        </tr>""")

    rows_str = "\n".join(rows_html)

    val_badge = (f'<span class="badge err-badge">失败 {s["fail"]}</span>'
                 if s["fail"] else "") + \
                (f'<span class="badge warn-badge">异常 {s["error"]}</span>'
                 if s["error"] else "") + \
                (('<span class="badge ok-badge">全部通过</span>')
                 if not s["fail"] and not s["error"] else "")
    val_html = f"""
<details open class="section">
<summary>二、数值比对（AcuCloud 快照 vs 实时 Modbus）
  <span class="sum-info">共 {s['total']} 项 · 通过率 {s['pass_rate']}</span>
  {val_badge}
</summary>
<div class="section-body">
{dup_warn_html}
<div class="cards">
  <div class="card total"><div class="num">{s['total']}</div><div class="lbl">总参数</div></div>
  <div class="card pass"> <div class="num">{s['pass']}</div> <div class="lbl">通过 PASS ({s['pass_rate']})</div></div>
  <div class="card fail"> <div class="num">{s['fail']}</div> <div class="lbl">失败 FAIL</div></div>
  <div class="card err">  <div class="num">{s['error']}</div><div class="lbl">异常 ERROR</div></div>
</div>
<table>
<colgroup>
  <col class="c-idx">
  <col class="c-key">
  <col class="c-val">
  <col class="c-val">
  <col class="c-diff">
  <col class="c-pct">
  <col class="c-stat">
  <col class="c-err">
</colgroup>
<thead>
  <tr>
    <th>#</th>
    <th>参数名 (param_key)</th>
    <th>Cloud 快照值</th>
    <th>Modbus 实时值</th>
    <th>绝对差值</th>
    <th>相对差值</th>
    <th>结果</th>
    <th>错误信息</th>
  </tr>
</thead>
<tbody>
{rows_str}
</tbody>
</table>
</div>
</details>"""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>AcuCloud vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>
  body {{ font-family: "Microsoft YaHei", Arial, sans-serif; font-size: 13px;
          margin: 20px; background: #f5f5f5; color: #333; }}
  h1   {{ font-size: 20px; margin-bottom: 4px; }}
  h2   {{ font-size: 15px; color: #0056b3; margin: 24px 0 10px; border-left: 4px solid #0056b3;
          padding-left: 8px; }}
  h3   {{ font-size: 13px; margin: 8px 0 4px; }}
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
  .warn-box {{ background: #fff3cd; border: 1px solid #ffc107; border-radius: 6px;
              padding: 10px 16px; margin-bottom: 16px; font-size: 12px; }}
  .warn-box ul {{ margin: 4px 0 0; padding-left: 18px; color: #533f03; }}
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
<h1>AcuCloud 快照 vs 实时 Modbus 比对报告</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  数据文件：{html.escape(xlsx_basename)} &nbsp;|&nbsp;
  快照时间戳：{html.escape(timestamp_str)} &nbsp;|&nbsp;
  来源 Sheet：{html.escape(sheet_name)} &nbsp;|&nbsp;
  Modbus 设备：{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT} &nbsp;|&nbsp;
  容差：±{config.CLOUD_TOLERANCE_PERCENT}% / ±{config.CLOUD_TOLERANCE_ABSOLUTE}
</div>
{scope_html}
{val_html}
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("HTML 报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 设备映射表（与 BACnetIP/comparator.py 保持一致）
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
    "acuiom03":   ("AcuIOM03",   "devices.acuiom03"),
    "acuiom04":   ("AcuIOM04",   "devices.acuiom04"),
}


# ─────────────────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────────────────

async def _main(
    xlsx_path: str,
    row_index: int,
    param_keys: Optional[list[str]],
) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    scope, results, timestamp_str, sheet_name, dup_headers = await run_cloud_comparison(
        xlsx_path, row_index, param_keys,
    )
    print_summary(scope, results)
    report_path = generate_html_report(
        scope,
        results,
        timestamp_str=timestamp_str,
        sheet_name=sheet_name,
        xlsx_path=xlsx_path,
        dup_headers=dup_headers,
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
    _xlsx_path: str
    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        _xlsx_path = sys.argv[idx + 1]
    else:
        _device_hint = _dev_value or ""
        _xlsx_path = find_latest_xlsx(config.CLOUD_DATA_DIR, device_hint=_device_hint)
        print(f"[INFO] 自动选取文件：{Path(_xlsx_path).name}")

    # ── --row ─────────────────────────────────────────────────────────────────
    _row_index = -1
    if "--row" in sys.argv:
        idx = sys.argv.index("--row")
        _row_index = int(sys.argv[idx + 1]) - 1  # 转为 0-based

    # ── --keys ────────────────────────────────────────────────────────────────
    _keys: Optional[list[str]] = None
    if "--keys" in sys.argv:
        idx = sys.argv.index("--keys")
        _keys = sys.argv[idx + 1:]

    asyncio.run(_main(_xlsx_path, _row_index, _keys))
