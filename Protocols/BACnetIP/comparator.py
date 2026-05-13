# -*- coding: utf-8 -*-
"""
comparator.py — BACnet vs Modbus 数值比对模块

流程：
  1. BACnetReader 发现设备对象列表
  2. 过滤出 Modbus 地址表中存在对应地址的参数
  3. BACnet 与 Modbus **并发**读取（节省时间）
  4. 逐参数按容差规则比对
  5. 输出控制台摘要 + HTML 报告

容差规则（来自 config）：
  pass if |diff| <= max(TOLERANCE_ABSOLUTE, ref × TOLERANCE_PERCENT/100)
  其中 ref = max(|bacnet_value|, |modbus_value|)
  当 ref 趋近于零时自动退化为绝对容差，无需手动切换。

用法：
  python BACnetIP/comparator.py              # 比对全部公共参数
  python BACnetIP/comparator.py --quick      # 只比对前 30 个参数（快速验证）
  python BACnetIP/comparator.py --keys FREQ_Hz VLN_a_V I_a_A
"""

from __future__ import annotations

import asyncio
import html
import logging
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

import config
from bacnet_reader import (BACnetReader, AIObject, ReadResult,
                            DeviceInfoResult, ProtocolErrorTestResult, StabilityResult)
from modbus_reader import ModbusReader, ModbusResult, get_reader
from template_reader import TemplateParam, find_template_file, get_bacnet_params, get_bacnet_params_by_range, natural_sort_key

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ScopeReport:
    """模板 BACnet 参数范围 vs 网关实际发布范围的对比报告。"""
    template_count:    int
    gateway_count:     int
    matched_keys:      list[str]   # 两侧均有
    missing_from_gw:   list[str]   # 模板有但网关未发布
    extra_in_gw:       list[str]   # 网关发布但模板未包含

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_gw and not self.extra_in_gw


@dataclass
class MetaCheckResult:
    """单个参数的单位模板 vs BACnet 比对结果。"""
    param_key:   str
    tmpl_unit:   str   # 模板 unit 列
    bacnet_unit: str   # BACnet units 属性
    unit_ok:     bool

    @property
    def ok(self) -> bool:
        return self.unit_ok


@dataclass
class CompareResult:
    """单个参数的 BACnet vs Modbus 比对结果。"""
    param_key:     str
    bacnet_value:  Optional[float] = None
    modbus_value:  Optional[float] = None
    bacnet_error:  str = ""
    modbus_error:  str = ""
    diff_abs:      Optional[float] = None
    diff_pct:      Optional[float] = None
    status:        str = ""   # PASS | FAIL | BACNET_ERR | MODBUS_ERR | BOTH_ERR

    @property
    def ok(self) -> bool:
        return self.status == "PASS"

    def __str__(self) -> str:
        if self.status == "PASS":
            return (f"[PASS] {self.param_key:40s} "
                    f"BACnet={self.bacnet_value:.6g}  "
                    f"Modbus={self.modbus_value:.6g}  "
                    f"Δ={self.diff_abs:.4g} ({self.diff_pct:.2f}%)")
        if self.status == "FAIL":
            return (f"[FAIL] {self.param_key:40s} "
                    f"BACnet={self.bacnet_value:.6g}  "
                    f"Modbus={self.modbus_value:.6g}  "
                    f"Δ={self.diff_abs:.4g} ({self.diff_pct:.2f}%)  ← 超出容差")
        if self.status == "BACNET_ERR":
            return f"[ERR ] {self.param_key:40s} BACnet错误: {self.bacnet_error}"
        if self.status == "MODBUS_ERR":
            return f"[ERR ] {self.param_key:40s} Modbus错误: {self.modbus_error}"
        return f"[ERR ] {self.param_key:40s} 双路错误"


# ─────────────────────────────────────────────────────────────────────────────
# 比对逻辑
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one(
    param_key: str,
    br: ReadResult,
    mr: ModbusResult,
) -> CompareResult:
    """对单个参数执行比对，返回 CompareResult。"""
    cr = CompareResult(param_key=param_key)

    if not br.ok:
        cr.bacnet_error = br.error
    if not mr.ok:
        cr.modbus_error = mr.error

    if not br.ok and not mr.ok:
        cr.status = "BOTH_ERR"
        return cr
    if not br.ok:
        cr.status = "BACNET_ERR"
        return cr
    if not mr.ok:
        cr.status = "MODBUS_ERR"
        return cr

    bv = br.value
    mv = mr.value
    cr.bacnet_value = bv
    cr.modbus_value = mv

    diff = abs(bv - mv)
    ref  = max(abs(bv), abs(mv))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0

    tol = max(config.TOLERANCE_ABSOLUTE, ref * config.TOLERANCE_PERCENT / 100)
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


def compare_all(
    bacnet_map:  dict[str, ReadResult],
    modbus_map:  dict[str, ModbusResult],
    common_keys: list[str],
) -> list[CompareResult]:
    """对所有公共参数执行比对，返回 CompareResult 列表。"""
    results: list[CompareResult] = []
    for key in common_keys:
        br = bacnet_map.get(key)
        mr = modbus_map.get(key)
        if br is None or mr is None:
            continue
        results.append(_compare_one(key, br, mr))
    return results


# ─────────────────────────────────────────────────────────────────────────────
# 主流程：并发读取 + 比对
# ─────────────────────────────────────────────────────────────────────────────

async def run_comparison(
    param_keys: Optional[list[str]] = None,
    check_meta: bool = False,
    check_proto: bool = True,
) -> tuple[ScopeReport, list[MetaCheckResult], list[CompareResult],
           Optional[DeviceInfoResult], list[ProtocolErrorTestResult],
           Optional[StabilityResult]]:
    """
    完整比对流程：范围检查 → （可选）元数据检查 → 数值比对 → 协议规范测试。

    Args:
        param_keys:   要比对的参数名列表；None 表示比对全部匹配参数。
        check_meta:   True 时额外读取 BACnet description/units 并与模板对比。
        check_proto:  True 时执行 Device Object 属性、错误响应、稳定性测试。

    Returns:
        (ScopeReport, list[MetaCheckResult], list[CompareResult],
         DeviceInfoResult, list[ProtocolErrorTestResult], StabilityResult)
    """
    t0 = time.time()

    # ── 加载模板范围 ───────────────────────────────────────────────────────────
    try:
        tmpl_path    = find_template_file(config.TEMPLATE_DIR, config.DEVICE_NAME)
        range_marker = getattr(config, 'BACNET_RANGE_MARKER', '')
        if range_marker:
            tmpl_params = get_bacnet_params_by_range(tmpl_path, range_marker)
        else:
            tmpl_params = get_bacnet_params(tmpl_path)
        tmpl_map    = {p.param_key: p for p in tmpl_params}
        tmpl_keys   = set(tmpl_map)
    except Exception as exc:
        log.warning("无法加载模板文件，范围检查将跳过：%s", exc)
        tmpl_map, tmpl_keys = {}, set()

    async with BACnetReader() as bacnet, get_reader() as modbus:

        # ── 发现 BACnet 对象 ──────────────────────────────────────────────────
        log.info("正在发现 BACnet 对象列表…")
        objects: list[AIObject] = await bacnet.discover_objects()

        gw_keys = {o.param_key for o in objects}

        # ── 范围检查（BACnet 模板 vs 网关实际发布） ────────────────────────────
        matched_keys_set = tmpl_keys & gw_keys if tmpl_keys else gw_keys
        missing_from_gw  = sorted(tmpl_keys - gw_keys, key=natural_sort_key)
        extra_in_gw      = sorted(gw_keys - tmpl_keys, key=natural_sort_key)
        scope_report = ScopeReport(
            template_count  = len(tmpl_keys),
            gateway_count   = len(gw_keys),
            matched_keys    = sorted(matched_keys_set, key=natural_sort_key),
            missing_from_gw = missing_from_gw,
            extra_in_gw     = extra_in_gw,
        )
        log.info("范围检查：模板=%d  网关=%d  匹配=%d  缺失=%d  多余=%d",
                 len(tmpl_keys), len(gw_keys), len(matched_keys_set),
                 len(missing_from_gw), len(extra_in_gw))

        # ── 确定比对参数集合 ──────────────────────────────────────────────────
        modbus_keys = set(modbus.known_params())
        if param_keys is not None:
            target_keys = [k for k in param_keys
                           if k in modbus_keys and k in matched_keys_set]
        else:
            target_keys = [k for k in scope_report.matched_keys if k in modbus_keys]

        target_set     = set(target_keys)
        bacnet_objects = [o for o in objects if o.param_key in target_set]
        log.info("数值比对参数数量：%d", len(target_keys))

        # ── 元数据检查（可选） ────────────────────────────────────────────────
        meta_results: list[MetaCheckResult] = []
        if check_meta and bacnet_objects:
            log.info("读取 BACnet 元数据（description / units）…")
            meta_raw = await bacnet.read_metadata_batch(bacnet_objects)
            for obj, bacnet_desc, bacnet_unit in meta_raw:
                tmpl = tmpl_map.get(obj.param_key)
                if tmpl is None:
                    continue
                meta_results.append(MetaCheckResult(
                    param_key   = obj.param_key,
                    tmpl_unit   = tmpl.unit,
                    bacnet_unit = bacnet_unit,
                    unit_ok     = tmpl.unit == bacnet_unit,
                ))

        # ── 读取数值 ──────────────────────────────────────────────────────────
        log.info("读取 BACnet Present Value…")
        bacnet_results = await bacnet.read_batch(bacnet_objects)

        log.info("读取 Modbus 寄存器…")
        modbus_results = await modbus.read_params(target_keys)

        bacnet_map: dict[str, ReadResult]  = {r.obj.param_key: r for r in bacnet_results}
        modbus_map: dict[str, ModbusResult] = {r.param_key: r for r in modbus_results}

        # ── 协议规范测试（复用同一 BACnet 连接）──────────────────────────────
        dev_info: Optional[DeviceInfoResult] = None
        err_tests: list[ProtocolErrorTestResult] = []
        stability: Optional[StabilityResult] = None

        if check_proto:
            log.info("执行 BACnet 协议规范测试…")
            dev_info  = await bacnet.read_device_info()
            probe_objs = bacnet_objects[:1] if bacnet_objects else objects[:1]
            err_tests  = await bacnet.test_error_responses(probe_objs)
            stability  = await bacnet.check_stability(probe_objs)

    # ── 数值比对 ──────────────────────────────────────────────────────────────
    compare_results = compare_all(bacnet_map, modbus_map, target_keys)
    elapsed = time.time() - t0
    log.info("比对完成，耗时 %.1f 秒", elapsed)

    return scope_report, meta_results, compare_results, dev_info, err_tests, stability


# ─────────────────────────────────────────────────────────────────────────────
# 统计摘要
# ─────────────────────────────────────────────────────────────────────────────

def summary(results: list[CompareResult]) -> dict:
    """返回比对结果统计字典。"""
    by_status: dict[str, int] = {}
    for r in results:
        by_status[r.status] = by_status.get(r.status, 0) + 1

    total    = len(results)
    passed   = by_status.get("PASS", 0)
    failed   = by_status.get("FAIL", 0)
    errors   = total - passed - failed

    fail_list = [r for r in results if r.status == "FAIL"]
    fail_list.sort(key=lambda r: r.diff_pct or 0, reverse=True)

    return {
        "total":        total,
        "pass":         passed,
        "fail":         failed,
        "error":        errors,
        "pass_rate":    f"{passed/total*100:.1f}%" if total else "N/A",
        "by_status":    by_status,
        "worst_fails":  fail_list[:10],
    }


def print_summary(
    scope: ScopeReport,
    meta: list[MetaCheckResult],
    results: list[CompareResult],
    dev_info: Optional[DeviceInfoResult] = None,
    err_tests: Optional[list[ProtocolErrorTestResult]] = None,
    stability: Optional[StabilityResult] = None,
) -> None:
    """打印比对摘要到控制台。"""
    print("\n" + "=" * 70)
    print(f"  BACnet vs Modbus 比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"  【参数范围】模板={scope.template_count}  网关={scope.gateway_count}  "
          f"匹配={len(scope.matched_keys)}  "
          f"缺失={len(scope.missing_from_gw)}  多余={len(scope.extra_in_gw)}")
    if scope.missing_from_gw:
        print(f"  缺失参数（前10）: {scope.missing_from_gw[:10]}")
    if scope.extra_in_gw:
        print(f"  多余参数（前10）: {scope.extra_in_gw[:10]}")
    if meta:
        meta_fail = [m for m in meta if not m.ok]
        print(f"  【单位检查】共 {len(meta)} 项  通过={len(meta)-len(meta_fail)}  "
              f"失败={len(meta_fail)}")
    s = summary(results)
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    if s["worst_fails"]:
        print("\n  差异最大的失败参数（Top 10）：")
        for r in s["worst_fails"]:
            print(f"    {r.param_key:40s}  BACnet={r.bacnet_value:.6g}  "
                  f"Modbus={r.modbus_value:.6g}  Δ%={r.diff_pct:.2f}%")
    if dev_info is not None:
        status = "OK" if dev_info.ok else f"FAIL（{dev_info.error}）"
        print(f"  【Device Object】{status}  设备={dev_info.object_name}  "
              f"厂商={dev_info.vendor_name}  固件={dev_info.firmware_revision}  "
              f"协议版本={dev_info.protocol_version}.{dev_info.protocol_revision}")
    if err_tests:
        passed = sum(1 for t in err_tests if t.passed)
        print(f"  【协议合规性】{passed}/{len(err_tests)} 通过")
    if stability is not None:
        status = "OK" if stability.ok else "FAIL"
        print(f"  【连接稳定性】{status}  {stability.successes}/{stability.attempts} 次成功")
    print("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告生成
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "PASS":       "#d4edda",
    "FAIL":       "#f8d7da",
    "BACNET_ERR": "#fff3cd",
    "MODBUS_ERR": "#fff3cd",
    "BOTH_ERR":   "#e2e3e5",
}

_STATUS_LABEL = {
    "PASS":       "通过",
    "FAIL":       "失败",
    "BACNET_ERR": "BACnet异常",
    "MODBUS_ERR": "Modbus异常",
    "BOTH_ERR":   "双路异常",
}


def _fmt(v: Optional[float], digits: int = 6) -> str:
    if v is None:
        return "—"
    s = f"{v:.{digits}f}"
    return s.rstrip('0').rstrip('.') if '.' in s else s


def generate_html_report(
    scope: ScopeReport,
    meta: list[MetaCheckResult],
    results: list[CompareResult],
    output_path: Optional[str] = None,
    dev_info: Optional[DeviceInfoResult] = None,
    err_tests: Optional[list[ProtocolErrorTestResult]] = None,
    stability: Optional[StabilityResult] = None,
) -> str:
    """生成六段式 HTML 比对报告，返回文件路径。"""
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"bacnet_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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
        f'<h3 style="color:#721c24;margin-top:12px">模板有但网关未发布（{len(scope.missing_from_gw)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.missing_from_gw, "#f8d7da")}</tbody></table>'
    ) if scope.missing_from_gw else ""

    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">网关发布但模板未包含（{len(scope.extra_in_gw)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.extra_in_gw, "#fff3cd")}</tbody></table>'
    ) if scope.extra_in_gw else ""

    scope_badge = (f'<span class="badge err-badge">缺失 {len(scope.missing_from_gw)}</span>'
                   if scope.missing_from_gw else "") + \
                  (f'<span class="badge warn-badge">多余 {len(scope.extra_in_gw)}</span>'
                   if scope.extra_in_gw else "") + \
                  (f'<span class="badge ok-badge">一致</span>'
                   if scope.scope_ok else "")
    scope_html = f"""
<details open class="section">
<summary>一、参数范围检查
  <span class="sum-info">模板 {scope.template_count} / 网关 {scope.gateway_count} / 匹配 {len(scope.matched_keys)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板 BACnet 参数</div></div>
  <div class="card total"><div class="num">{scope.gateway_count}</div><div class="lbl">网关实际发布</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_gw)}</div><div class="lbl">模板有/网关缺</div></div>
  <div class="card err">  <div class="num">{len(scope.extra_in_gw)}</div><div class="lbl">网关多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 元数据检查 HTML ────────────────────────────────────────────
    if meta:
        meta_fail = [m for m in meta if not m.ok]
        meta_pass = len(meta) - len(meta_fail)
        meta_rows = []
        for i, m in enumerate(meta):
            bg_u = "#d4edda" if m.unit_ok else "#f8d7da"
            meta_rows.append(f"""
        <tr>
          <td class="num">{i+1}</td>
          <td class="key">{html.escape(m.param_key)}</td>
          <td>{html.escape(m.tmpl_unit)}</td>
          <td style="background:{bg_u}">{html.escape(m.bacnet_unit)}</td>
          <td style="background:{bg_u};text-align:center">{'✓' if m.unit_ok else '✗'}</td>
        </tr>""")
        meta_badge = (f'<span class="badge err-badge">失败 {len(meta_fail)}</span>'
                      if meta_fail else
                      '<span class="badge ok-badge">全部通过</span>')
        meta_html = f"""
<details open class="section">
<summary>二、单位检查（BACnet units 属性 vs 模板）
  <span class="sum-info">共 {len(meta)} 项</span>
  {meta_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(meta)}</div><div class="lbl">检查总数</div></div>
  <div class="card pass"> <div class="num">{meta_pass}</div><div class="lbl">通过</div></div>
  <div class="card fail"> <div class="num">{len(meta_fail)}</div><div class="lbl">失败</div></div>
</div>
<table>
<colgroup>
  <col style="width:48px"><col style="width:280px">
  <col style="width:120px"><col style="width:120px"><col style="width:48px">
</colgroup>
<thead>
  <tr>
    <th>#</th><th>param_key</th>
    <th>模板单位</th><th>BACnet 单位</th><th>匹配</th>
  </tr>
</thead>
<tbody>{"".join(meta_rows)}</tbody>
</table>
</div>
</details>"""
    else:
        meta_html = f"""
<details class="section">
<summary>二、单位检查（BACnet units 属性 vs 模板）
  <span class="sum-info">未启用</span>
  <span class="badge warn-badge">跳过</span>
</summary>
<div class="section-body">
<p style="color:#888;font-size:12px">元数据检查未启用（使用 --no-meta 可关闭，默认开启）</p>
</div>
</details>"""

    # ── Section 3: 数值比对 HTML ──────────────────────────────────────────────
    val_rows = []
    for i, r in enumerate(results):
        bg   = _STATUS_COLOR.get(r.status, "#ffffff")
        stat = _STATUS_LABEL.get(r.status, r.status)
        bv_str = _fmt(r.bacnet_value)
        mv_str = _fmt(r.modbus_value)
        da_str = _fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"
        dp_str = f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"
        err_hint = ""
        if r.bacnet_error:
            err_hint += f"BACnet: {html.escape(r.bacnet_error)}"
        if r.modbus_error:
            if err_hint:
                err_hint += "<br>"
            err_hint += f"Modbus: {html.escape(r.modbus_error)}"
        val_rows.append(f"""
        <tr style="background:{bg}">
          <td class="num">{i+1}</td>
          <td class="key">{html.escape(r.param_key)}</td>
          <td class="val">{bv_str}</td>
          <td class="val">{mv_str}</td>
          <td class="val">{da_str}</td>
          <td class="val">{dp_str}</td>
          <td class="stat">{stat}</td>
          <td class="err">{err_hint}</td>
        </tr>""")

    val_badge = (f'<span class="badge err-badge">失败 {s["fail"]}</span>'
                 if s["fail"] else "") + \
                (f'<span class="badge warn-badge">异常 {s["error"]}</span>'
                 if s["error"] else "") + \
                (('<span class="badge ok-badge">全部通过</span>')
                 if not s["fail"] and not s["error"] else "")
    val_html = f"""
<details open class="section">
<summary>三、数值比对（BACnet Present Value vs Modbus）
  <span class="sum-info">共 {s['total']} 项 · 通过率 {s['pass_rate']}</span>
  {val_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{s['total']}</div><div class="lbl">总参数</div></div>
  <div class="card pass"> <div class="num">{s['pass']}</div><div class="lbl">通过 PASS ({s['pass_rate']})</div></div>
  <div class="card fail"> <div class="num">{s['fail']}</div><div class="lbl">失败 FAIL</div></div>
  <div class="card err">  <div class="num">{s['error']}</div><div class="lbl">异常 ERROR</div></div>
</div>
<table>
<colgroup>
  <col class="c-idx"><col class="c-key">
  <col class="c-val"><col class="c-val">
  <col class="c-diff"><col class="c-pct">
  <col class="c-stat"><col class="c-err">
</colgroup>
<thead>
  <tr>
    <th>#</th><th>参数名 (param_key)</th>
    <th>BACnet 值</th><th>Modbus 值</th>
    <th>绝对差值</th><th>相对差值</th>
    <th>结果</th><th>错误信息</th>
  </tr>
</thead>
<tbody>{"".join(val_rows)}</tbody>
</table>
</div>
</details>"""

    # ── Section 4: Device Object 属性 ────────────────────────────────────────
    if dev_info is not None:
        dev_ok_color = "#d4edda" if dev_info.ok else "#f8d7da"
        dev_ok_label = "OK" if dev_info.ok else "FAIL"
        _DEV_PROPS = [
            ("设备实例号 (objectIdentifier)",      dev_info.object_identifier),
            ("设备名称 (objectName)",               dev_info.object_name),
            ("系统状态 (systemStatus)",              dev_info.system_status),
            ("厂商名称 (vendorName)",                dev_info.vendor_name),
            ("厂商 ID (vendorIdentifier)",          dev_info.vendor_id),
            ("型号名称 (modelName)",                 dev_info.model_name),
            ("固件版本 (firmwareRevision)",          dev_info.firmware_revision),
            ("应用软件版本 (applicationSoftwareVersion)", dev_info.app_sw_version),
            ("协议版本 (protocolVersion)",           dev_info.protocol_version),
            ("协议修订号 (protocolRevision)",        dev_info.protocol_revision),
            ("最大 APDU 长度 (maxApduLengthAccepted)", dev_info.max_apdu_length),
            ("分段支持 (segmentationSupported)",     dev_info.segmentation),
        ]
        dev_rows = "".join(
            f'<tr style="background:{"#d4edda" if v else "#f8d7da"}">'
            f'<td class="key">{html.escape(name)}</td>'
            f'<td class="val">{html.escape(v) if v else "<em style=\'color:#721c24\'>未获取</em>"}</td></tr>'
            for name, v in _DEV_PROPS
        )
        dev_badge = f'<span class="badge {"ok-badge" if dev_info.ok else "err-badge"}">{dev_ok_label}</span>'
        dev_info_html = f"""
<details open class="section">
<summary>四、BACnet Device Object 属性（ANSI/ASHRAE 135 §12.11）
  <span class="sum-info">必需属性合规性验证</span>
  {dev_badge}
</summary>
<div class="section-body">
<table>
<colgroup><col style="width:360px"><col></colgroup>
<thead><tr><th>属性（BACnet 标准名）</th><th>读取值</th></tr></thead>
<tbody>{dev_rows}</tbody>
</table>
</div>
</details>"""
    else:
        dev_info_html = ""

    # ── Section 5: 协议合规性测试 ─────────────────────────────────────────────
    if err_tests:
        passed_n = sum(1 for t in err_tests if t.passed)
        proto_badge = (f'<span class="badge err-badge">失败 {len(err_tests)-passed_n}</span>'
                       if passed_n < len(err_tests) else
                       '<span class="badge ok-badge">全部通过</span>')
        err_rows = "".join(
            f'<tr style="background:{"#d4edda" if t.passed else "#f8d7da"}">'
            f'<td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(t.test_name)}</td>'
            f'<td class="stat">{"通过" if t.passed else "失败"}</td>'
            f'<td class="err">{html.escape(t.detail)}</td></tr>'
            for i, t in enumerate(err_tests)
        )
        err_tests_html = f"""
<details open class="section">
<summary>五、协议合规性测试（ANSI/ASHRAE 135 §16 错误处理 + AI 必需属性）
  <span class="sum-info">{passed_n}/{len(err_tests)} 通过</span>
  {proto_badge}
</summary>
<div class="section-body">
<table>
<colgroup><col style="width:48px"><col><col style="width:80px"><col style="width:260px"></colgroup>
<thead><tr><th>#</th><th>测试项</th><th>结果</th><th>详情</th></tr></thead>
<tbody>{err_rows}</tbody>
</table>
</div>
</details>"""
    else:
        err_tests_html = ""

    # ── Section 6: 连接稳定性 ────────────────────────────────────────────────
    if stability is not None:
        stab_ok = stability.ok
        stab_badge = ('<span class="badge ok-badge">稳定</span>' if stab_ok else
                      '<span class="badge err-badge">不稳定</span>')
        stab_color = "#d4edda" if stab_ok else "#f8d7da"
        err_list_html = ("".join(
            f'<li style="color:#721c24">{html.escape(e)}</li>' for e in stability.errors
        )) if stability.errors else ""
        stab_err_attr = '' if stab_ok else ' data-has-error="1"'
        stability_html = f"""
<details open class="section"{stab_err_attr}>
<summary>六、连接稳定性测试
  <span class="sum-info">连续读取 {stability.attempts} 次</span>
  {stab_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{stability.attempts}</div><div class="lbl">总次数</div></div>
  <div class="card pass"><div class="num">{stability.successes}</div><div class="lbl">成功</div></div>
  <div class="card fail"><div class="num">{stability.attempts - stability.successes}</div><div class="lbl">失败</div></div>
</div>
{f'<ul style="margin:0;padding-left:20px">{err_list_html}</ul>' if err_list_html else ""}
</div>
</details>"""
    else:
        stability_html = ""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>BACnet vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
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
<button id="err-toggle" onclick="toggleErrOnly()"
  style="position:fixed;top:16px;right:20px;z-index:999;padding:6px 16px;
         background:#dc3545;color:#fff;border:none;border-radius:4px;
         cursor:pointer;font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,.25)">
  仅显示异常
</button>
<script>
function toggleErrOnly() {{
  var btn = document.getElementById('err-toggle');
  var on = btn.dataset.active === '1';
  if (on) {{
    document.querySelectorAll('details.section').forEach(function(d) {{ d.style.display = ''; }});
    document.querySelectorAll('tbody tr').forEach(function(tr) {{ tr.style.display = ''; }});
  }} else {{
    document.querySelectorAll('details.section').forEach(function(d) {{ d.open = true; }});
    document.querySelectorAll('tbody tr').forEach(function(tr) {{
      var _s = (tr.getAttribute('style')||'').toLowerCase();
      tr.style.display = _s.indexOf('d4edda') !== -1 ? 'none' : '';
    }});
    document.querySelectorAll('details.section').forEach(function(d) {{
      var hasErr = d.dataset.hasError === '1' || Array.from(d.querySelectorAll('tbody tr')).some(function(tr) {{
        return (tr.getAttribute('style')||'').toLowerCase().indexOf('d4edda') === -1;
      }});
      if (!hasErr) {{ d.style.display = 'none'; }}
    }});
  }}
  btn.textContent      = on ? '仅显示异常' : '显示全部';
  btn.style.background = on ? '#dc3545'    : '#6c757d';
  btn.dataset.active   = on ? '0' : '1';
}}
</script>
<h1>BACnet vs Modbus 比对报告</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  网关：{config.GATEWAY_IP}:{config.GATEWAY_PORT} &nbsp;|&nbsp;
  Modbus：{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT} &nbsp;|&nbsp;
  容差：±{config.TOLERANCE_PERCENT}% / ±{config.TOLERANCE_ABSOLUTE}
</div>
{scope_html}
{meta_html}
{val_html}
{dev_info_html}
{err_tests_html}
{stability_html}
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("HTML 报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────────────────

async def _main(param_keys: Optional[list[str]], quick: bool,
               check_meta: bool, check_proto: bool) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    if quick and param_keys is None:
        # 快速模式：只取前 30 个常用参数
        from modbus_reader import get_param_map
        all_keys = list(get_param_map().keys())
        param_keys = all_keys[:30]
        log.info("快速模式：仅比对前 %d 个参数", len(param_keys))

    scope, meta, results, dev_info, err_tests, stability = await run_comparison(
        param_keys, check_meta=check_meta, check_proto=check_proto
    )
    print_summary(scope, meta, results, dev_info, err_tests, stability)
    report_path = generate_html_report(scope, meta, results,
                                       dev_info=dev_info,
                                       err_tests=err_tests,
                                       stability=stability)
    print(f"\n  HTML 报告：{report_path}\n")


# 设备名称 → (DEVICE_NAME, DEVICE_MODULE) 映射
_DEVICE_MAP: dict[str, tuple[str, str]] = {
    "acurev4100":  ("AcuRev4100",  "devices.acurev4100"),
    "acurev2100":  ("AcuRev2100",  "devices.acurev2100"),
    "acuvimiiw":   ("AcuvimIIW",   "devices.acuvimiiw"),
    "acuvimiir":   ("AcuvimIIR",   "devices.acuvimiir"),
    "acuvim3":     ("AcuVIM3",     "devices.acuvim3"),
    "pxm350":      ("PXM350",      "devices.pxm350"),
    "acuiom01":    ("AcuIOM01",    "devices.acuiom01"),
    "acuiom02":    ("AcuIOM02",    "devices.acuiom02"),
    "acuiom03":    ("AcuIOM03",    "devices.acuiom03"),
    "acuiom04":    ("AcuIOM04",    "devices.acuiom04"),
}

# AcuIOM 设备需要使用 range 列过滤 BACnet 参数范围（无 BACnetIP 列）
_RANGE_MARKER_MAP: dict[str, str] = {
    "acuiom01": "8",
    "acuiom02": "8",
    "acuiom03": "10",
    "acuiom04": "10",
}


if __name__ == "__main__":
    import sys as _sys

    quick       = "--quick"    in _sys.argv
    check_meta  = "--no-meta"  not in _sys.argv   # 默认开启，--no-meta 可关闭
    check_proto = "--no-proto" not in _sys.argv   # 默认开启，--no-proto 可关闭

    # 无论是否传 --device，都先从 MODBUS_DEVICE_MAP 初始化默认设备的连接参数
    if config.DEVICE_NAME in config.MODBUS_DEVICE_MAP:
        config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
            config.MODBUS_DEVICE_MAP[config.DEVICE_NAME]
    # 根据默认设备设置 BACnet range 过滤标记（AcuIOM 专用）
    _default_key = config.DEVICE_NAME.lower().replace("-", "")
    config.BACNET_RANGE_MARKER = _RANGE_MARKER_MAP.get(_default_key, "")

    # --device <名称>  覆盖 config 中的设备配置
    _dev_value: Optional[str] = None
    if "--device" in _sys.argv:
        idx = _sys.argv.index("--device")
        _dev_value = _sys.argv[idx + 1].lower()
        if _dev_value not in _DEVICE_MAP:
            print(f"[ERROR] 未知设备 '{_dev_value}'，可选：{list(_DEVICE_MAP)}")
            _sys.exit(1)
        config.DEVICE_NAME, config.DEVICE_MODULE = _DEVICE_MAP[_dev_value]
        # 同步更新 Modbus 连接参数和 BACnet range 过滤标记
        if config.DEVICE_NAME in config.MODBUS_DEVICE_MAP:
            config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
                config.MODBUS_DEVICE_MAP[config.DEVICE_NAME]
        config.BACNET_RANGE_MARKER = _RANGE_MARKER_MAP.get(_dev_value, "")
        print(f"[INFO] 已切换设备: {config.DEVICE_NAME}  "
              f"Modbus={config.MODBUS_HOST}:{config.MODBUS_PORT} unit={config.MODBUS_UNIT}")

    # 收集非 flag 参数，同时排除 --device 的值（否则会被误当作 param_key）
    _excluded = {_dev_value} if _dev_value else set()
    args = [a for a in _sys.argv[1:]
            if not a.startswith("--") and a.lower() not in _excluded]

    if "--keys" in _sys.argv:
        idx   = _sys.argv.index("--keys")
        keys  = _sys.argv[idx + 1:]
    elif args:
        keys = args
    else:
        keys = None

    asyncio.run(_main(keys, quick, check_meta, check_proto))
