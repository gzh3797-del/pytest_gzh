# -*- coding: utf-8 -*-
"""
enip_comparator.py — EtherNet/IP vs Modbus TCP 数值比对

流程：
  1. 解析 EDS + 模板 SNMP 列，执行参数范围检查
  2. 对匹配参数执行单位检查（EDS unit vs 模板 unit）
  3. 并发读取 EtherNet/IP Assembly 与 Modbus TCP，逐参数按容差比对
  4. 输出控制台摘要 + HTML 报告（三段式：范围 + 单位 + 数值）

用法（从仓库根目录执行）：
  python Protocols/EtherNetIP/enip_comparator.py
  python Protocols/EtherNetIP/enip_comparator.py --quick
  python Protocols/EtherNetIP/enip_comparator.py --keys FREQ_Hz VLN_a_V
  python Protocols/EtherNetIP/enip_comparator.py --device AcuRev4100
"""
from __future__ import annotations

import asyncio
import html
import logging
import sys
import time
import re as _re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

import config
from enip_reader import (EnipReader, EnipResult, find_eds_file, parse_eds,
                          IdentityResult, ErrorTestResult,
                          AssemblyIntegrityResult, StabilityResult,
                          read_identity, test_error_responses,
                          check_assembly_integrity, check_connection_stability)
from modbus_reader import ModbusReader, ModbusResult, get_reader
from template_reader import TemplateParam, find_template_file, get_snmp_params, natural_sort_key

log = logging.getLogger(__name__)


def _natural_key(s: str) -> list:
    return [int(c) if c.isdigit() else c.lower() for c in _re.split(r'(\d+)', s)]


# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ScopeReport:
    template_count:   int
    eds_count:        int
    matched_keys:     list[str]
    missing_from_eds: list[str]   # 模板 SNMP 有但 EDS Assembly 没有
    extra_in_eds:     list[str]   # EDS 有但模板未标注 SNMP

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_eds and not self.extra_in_eds


@dataclass
class UnitCheckResult:
    """单个参数的单位对比结果（EDS unit vs 模板 unit）。"""
    param_key: str
    tmpl_unit: str
    eds_unit:  str
    unit_ok:   bool

    @property
    def ok(self) -> bool:
        return self.unit_ok


@dataclass
class CompareResult:
    param_key:    str
    enip_value:   Optional[float] = None
    modbus_value: Optional[float] = None
    enip_error:   str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | ENIP_ERR | MODBUS_ERR | BOTH_ERR

    @property
    def ok(self) -> bool:
        return self.status == "PASS"


# ─────────────────────────────────────────────────────────────────────────────
# 比对逻辑
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one(
    param_key: str,
    er: EnipResult,
    mr: ModbusResult,
) -> CompareResult:
    cr = CompareResult(param_key=param_key)

    if not er.ok:
        cr.enip_error = er.error
    if not mr.ok:
        cr.modbus_error = mr.error

    if not er.ok and not mr.ok:
        cr.status = "BOTH_ERR"; return cr
    if not er.ok:
        cr.status = "ENIP_ERR"; return cr
    if not mr.ok:
        cr.status = "MODBUS_ERR"; return cr

    ev, mv = er.value, mr.value
    cr.enip_value   = ev
    cr.modbus_value = mv

    diff = abs(ev - mv)
    ref  = max(abs(ev), abs(mv))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0

    tol = max(config.ENIP_TOLERANCE_ABSOLUTE, ref * config.ENIP_TOLERANCE_PERCENT / 100)
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


def compare_all(
    enip_map:   dict[str, EnipResult],
    modbus_map: dict[str, ModbusResult],
    keys:       list[str],
) -> list[CompareResult]:
    results = []
    for key in keys:
        er = enip_map.get(key)
        mr = modbus_map.get(key)
        if er is None or mr is None:
            continue
        results.append(_compare_one(key, er, mr))
    return results


# ─────────────────────────────────────────────────────────────────────────────
# CIP 对象检查（Identity + 错误响应）
# ─────────────────────────────────────────────────────────────────────────────

async def run_cip_checks() -> tuple[IdentityResult, list[ErrorTestResult], StabilityResult]:
    """独立连接，依次执行：Identity 读取、错误响应测试、连接稳定性（3次重复读 Assembly）。"""
    from pycomm3 import CIPDriver

    def _sync() -> tuple[IdentityResult, list[ErrorTestResult], StabilityResult]:
        driver = CIPDriver(config.ENIP_HOST)
        driver.open()
        log.info("CIP 检查已连接：%s", config.ENIP_HOST)
        try:
            ident = read_identity(driver)
            if ident.ok:
                log.info("Identity 读取成功：%s  Rev=%s.%s  VendorID=%s",
                         ident.product_name, ident.revision_major,
                         ident.revision_minor, ident.vendor_id)
            else:
                log.warning("Identity 读取异常：%s", ident.error)
            tests = test_error_responses(driver)
            passed = sum(1 for t in tests if t.passed)
            log.info("错误响应测试：%d/%d 通过", passed, len(tests))
            stability = check_connection_stability(driver, attempts=3, delay=1.0)
            log.info("稳定性测试：%d/%d 成功", stability.successes, stability.attempts)
            return ident, tests, stability
        finally:
            try:
                driver.close()
            except Exception:
                pass

    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _sync)


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

async def run_comparison(
    param_keys: Optional[list[str]] = None,
) -> tuple[ScopeReport, list[UnitCheckResult], list[CompareResult]]:
    t0 = time.time()

    device = config.DEVICE_NAME

    # ── 1. 模板 SNMP 范围 ────────────────────────────────────────────────────
    tmpl_path   = find_template_file(config.TEMPLATE_DIR, device)
    tmpl_params = get_snmp_params(tmpl_path)
    tmpl_map    = {p.param_key: p for p in tmpl_params}
    tmpl_set    = set(tmpl_map)

    # ── 2. EDS Assembly 范围 ─────────────────────────────────────────────────
    eds_path   = find_eds_file(config.EDS_DIR, device)
    eds_params = parse_eds(eds_path)
    eds_map    = {p.name: p for p in eds_params}
    eds_set    = set(eds_map)
    eds_order  = [p.name for p in eds_params]   # 保持 Assembly 顺序

    # ── 3. 参数范围检查 ──────────────────────────────────────────────────────
    matched_set      = tmpl_set & eds_set
    missing_from_eds = sorted(tmpl_set - eds_set, key=_natural_key)
    extra_in_eds     = sorted(eds_set - tmpl_set, key=_natural_key)
    matched_keys     = sorted(matched_set, key=_natural_key)
    scope_report = ScopeReport(
        template_count   = len(tmpl_set),
        eds_count        = len(eds_set),
        matched_keys     = matched_keys,
        missing_from_eds = missing_from_eds,
        extra_in_eds     = extra_in_eds,
    )
    log.info("范围检查：模板SNMP=%d  EDS=%d  匹配=%d  缺失=%d  多余=%d",
             len(tmpl_set), len(eds_set), len(matched_set),
             len(missing_from_eds), len(extra_in_eds))

    # ── 4. 单位检查 ──────────────────────────────────────────────────────────
    unit_results: list[UnitCheckResult] = []
    for key in matched_keys:
        tmpl_unit = tmpl_map[key].unit
        eds_unit  = eds_map[key].unit
        unit_results.append(UnitCheckResult(
            param_key = key,
            tmpl_unit = tmpl_unit,
            eds_unit  = eds_unit,
            unit_ok   = (tmpl_unit == eds_unit),
        ))

    # ── 5. 确定数值比对目标 ──────────────────────────────────────────────────
    async with EnipReader(config.ENIP_HOST, eds_path, config.ENIP_SLOT) as enip, \
               get_reader() as modbus:

        modbus_keys_set = set(modbus.known_params())

        if param_keys:
            target = [k for k in param_keys
                      if k in matched_set and k in modbus_keys_set]
        else:
            # 按 Assembly 顺序，只取模板 SNMP & Modbus 均有的参数
            target = [k for k in eds_order
                      if k in matched_set and k in modbus_keys_set]

        log.info("数值比对参数数：%d", len(target))

        # ── 6. 并发读取 ──────────────────────────────────────────────────────
        enip_task   = asyncio.create_task(enip.read_all())
        modbus_task = asyncio.create_task(modbus.read_params(target))
        enip_map, modbus_results = await asyncio.gather(enip_task, modbus_task)
        actual_bytes = enip._actual_bytes

    modbus_map = {r.param_key: r for r in modbus_results}
    results    = compare_all(enip_map, modbus_map, target)

    # ── 7. Assembly 结构合规性 ────────────────────────────────────────────────
    integrity = check_assembly_integrity(eds_params, actual_bytes)
    log.info("Assembly 结构：EDS声明=%d B  实际=%d B  越界=%d 个",
             integrity.eds_total_bytes, integrity.actual_bytes,
             len(integrity.out_of_bounds))

    log.info("比对完成，耗时 %.1f 秒", time.time() - t0)
    return scope_report, unit_results, results, integrity


# ─────────────────────────────────────────────────────────────────────────────
# 统计 & 控制台输出
# ─────────────────────────────────────────────────────────────────────────────

def summary(results: list[CompareResult]) -> dict:
    by_status: dict[str, int] = {}
    for r in results:
        by_status[r.status] = by_status.get(r.status, 0) + 1
    total  = len(results)
    passed = by_status.get("PASS", 0)
    failed = by_status.get("FAIL", 0)
    errors = total - passed - failed
    fail_list = sorted([r for r in results if r.status == "FAIL"],
                       key=lambda r: r.diff_pct or 0, reverse=True)
    return {
        "total": total, "pass": passed, "fail": failed, "error": errors,
        "pass_rate": f"{passed/total*100:.1f}%" if total else "N/A",
        "by_status": by_status, "worst_fails": fail_list[:10],
    }


def print_summary(
    scope: ScopeReport,
    unit_results: list[UnitCheckResult],
    results: list[CompareResult],
    integrity: Optional[AssemblyIntegrityResult] = None,
    identity: Optional[IdentityResult] = None,
    err_tests: Optional[list[ErrorTestResult]] = None,
    stability: Optional[StabilityResult] = None,
) -> None:
    s = summary(results)
    print("\n" + "=" * 70)
    print(f"  EtherNet/IP vs Modbus 比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"  【参数范围】模板SNMP={scope.template_count}  EDS={scope.eds_count}  "
          f"匹配={len(scope.matched_keys)}  "
          f"缺失={len(scope.missing_from_eds)}  多余={len(scope.extra_in_eds)}")
    if scope.missing_from_eds:
        print(f"  缺失（前10）: {scope.missing_from_eds[:10]}")
    if scope.extra_in_eds:
        print(f"  多余（前10）: {scope.extra_in_eds[:10]}")
    if unit_results:
        unit_fail = [u for u in unit_results if not u.ok]
        print(f"  【单位检查】共 {len(unit_results)} 项  通过={len(unit_results)-len(unit_fail)}  "
              f"失败={len(unit_fail)}")
        for u in unit_fail[:5]:
            print(f"    {u.param_key:40s}  模板={u.tmpl_unit!r}  EDS={u.eds_unit!r}")
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    if s["worst_fails"]:
        print("\n  差异最大参数（Top 10）：")
        for r in s["worst_fails"]:
            print(f"    {r.param_key:40s}  EIP={r.enip_value:.6g}  "
                  f"Modbus={r.modbus_value:.6g}  Δ%={r.diff_pct:.2f}%")
    if integrity is not None:
        status_str = "OK" if integrity.ok else "FAIL"
        print(f"  【Assembly结构】{status_str}  EDS={integrity.eds_total_bytes}B  "
              f"实际={integrity.actual_bytes}B  越界={len(integrity.out_of_bounds)}")
        for item in integrity.out_of_bounds[:3]:
            print(f"    越界: {item}")
    if identity is not None:
        print(f"  【Identity Object】", "OK" if identity.ok else f"ERROR: {identity.error}")
        if identity.ok:
            print(f"    VendorID={identity.vendor_id}  DevType=0x{identity.device_type or 0:04X}  "
                  f"ProdCode=0x{identity.product_code or 0:04X}  "
                  f"Rev={identity.revision_major}.{identity.revision_minor}")
            print(f"    Serial=0x{identity.serial_number or 0:08X}  "
                  f"Name={identity.product_name!r}")
    if err_tests:
        passed = sum(1 for t in err_tests if t.passed)
        print(f"  【错误响应】共 {len(err_tests)} 项  通过={passed}  失败={len(err_tests)-passed}")
        for t in err_tests:
            icon = "✓" if t.passed else "✗"
            print(f"    {icon} {t.name}: {t.description[:60]}")
    if stability is not None:
        status_str = "OK" if stability.ok else "FAIL"
        print(f"  【连接稳定性】{status_str}  {stability.successes}/{stability.attempts} 次成功")
        for e in stability.errors:
            print(f"    错误: {e}")
    print("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "PASS":       "#d4edda",
    "FAIL":       "#f8d7da",
    "ENIP_ERR":   "#fff3cd",
    "MODBUS_ERR": "#fff3cd",
    "BOTH_ERR":   "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS":       "通过",
    "FAIL":       "失败",
    "ENIP_ERR":   "EIP异常",
    "MODBUS_ERR": "Modbus异常",
    "BOTH_ERR":   "双路异常",
}


def _fmt(v: Optional[float], d: int = 6) -> str:
    if v is None:
        return "—"
    s = f"{v:.{d}f}"
    return s.rstrip('0').rstrip('.') if '.' in s else s


_DEVICE_TYPE_NAME: dict[int, str] = {
    0x00: "Generic Device",
    0x02: "AC Drive",
    0x06: "General Purpose Discrete I/O",
    0x07: "General Purpose Analog I/O",
    0x0C: "Communications Adapter",
    0x12: "Programmable Logic Controller",
    0x1B: "Embedded Component",
}

_STATUS_STATE: dict[int, str] = {
    0: "Non-existent", 1: "Self-testing", 2: "Standby",
    3: "Operating", 4: "Major recov. fault", 5: "Major unrecov. fault",
}


def _decode_cip_status(status: int) -> str:
    state = _STATUS_STATE.get((status >> 2) & 0x0F, f"Unknown({(status >> 2) & 0x0F})")
    flags = []
    if status & 0x0020: flags.append("Minor recov. fault")
    if status & 0x0040: flags.append("Minor unrecov. fault")
    if status & 0x0080: flags.append("Major recov. fault")
    if status & 0x0100: flags.append("Major unrecov. fault")
    return state + (f" | {', '.join(flags)}" if flags else "")


def generate_html_report(
    scope: ScopeReport,
    unit_results: list[UnitCheckResult],
    results: list[CompareResult],
    output_path: Optional[str] = None,
    integrity: Optional[AssemblyIntegrityResult] = None,
    identity: Optional[IdentityResult] = None,
    err_tests: Optional[list[ErrorTestResult]] = None,
    stability: Optional[StabilityResult] = None,
) -> str:
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"enip_compare_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── Section 1: 参数范围检查 ──────────────────────────────────────────────
    scope_ok    = scope.scope_ok
    scope_color = "#d4edda" if scope_ok else "#f8d7da"
    scope_label = "一致" if scope_ok else "不一致"

    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    missing_html = (
        f'<h3 style="color:#721c24;margin-top:12px">模板有但 EDS 缺少（{len(scope.missing_from_eds)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.missing_from_eds, "#f8d7da")}</tbody></table>'
    ) if scope.missing_from_eds else ""

    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">EDS 有但模板未标 SNMP（{len(scope.extra_in_eds)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.extra_in_eds, "#fff3cd")}</tbody></table>'
    ) if scope.extra_in_eds else ""

    scope_badge = (f'<span class="badge err-badge">缺失 {len(scope.missing_from_eds)}</span>'
                   if scope.missing_from_eds else "") + \
                  (f'<span class="badge warn-badge">多余 {len(scope.extra_in_eds)}</span>'
                   if scope.extra_in_eds else "") + \
                  ('<span class="badge ok-badge">一致</span>' if scope_ok else "")

    scope_section = f"""
<details open class="section">
<summary>一、参数范围检查（模板 SNMP 列 vs EDS Assembly）
  <span class="sum-info">模板 {scope.template_count} / EDS {scope.eds_count} / 匹配 {len(scope.matched_keys)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板 SNMP 参数</div></div>
  <div class="card total"><div class="num">{scope.eds_count}</div><div class="lbl">EDS Assembly</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_eds)}</div><div class="lbl">模板有/EDS缺</div></div>
  <div class="card err">  <div class="num">{len(scope.extra_in_eds)}</div><div class="lbl">EDS多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 单位检查 ──────────────────────────────────────────────────
    unit_fail  = [u for u in unit_results if not u.ok]
    unit_ok_count = len(unit_results) - len(unit_fail)
    unit_color  = "#d4edda" if not unit_fail else "#f8d7da"
    unit_label  = "一致" if not unit_fail else "不一致"

    unit_rows_html = "".join(
        f'<tr style="background:{"#f8d7da" if not u.ok else "#d4edda"}">'
        f'<td class="num">{i+1}</td>'
        f'<td class="key">{html.escape(u.param_key)}</td>'
        f'<td class="val">{html.escape(u.tmpl_unit) or "（空）"}</td>'
        f'<td class="val">{html.escape(u.eds_unit) or "（空）"}</td>'
        f'<td class="stat">{"通过" if u.ok else "失败"}</td>'
        f'</tr>'
        for i, u in enumerate(unit_results)
    )

    unit_badge = (f'<span class="badge err-badge">不一致 {len(unit_fail)}</span>'
                  if unit_fail else '<span class="badge ok-badge">全部一致</span>')

    unit_section = f"""
<details {"open " if unit_fail else ""}class="section">
<summary>二、单位检查（EDS unit vs 模板 unit）
  <span class="sum-info">共 {len(unit_results)} 项 · 通过 {unit_ok_count} / 失败 {len(unit_fail)}</span>
  {unit_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(unit_results)}</div><div class="lbl">检查总数</div></div>
  <div class="card pass"> <div class="num">{unit_ok_count}</div><div class="lbl">一致</div></div>
  <div class="card fail"> <div class="num">{len(unit_fail)}</div><div class="lbl">不一致</div></div>
  <div class="card" style="background:{unit_color}"><div class="num" style="font-size:18px">{unit_label}</div><div class="lbl">单位结论</div></div>
</div>
<table>
<colgroup>
  <col style="width:48px"><col style="width:280px">
  <col style="width:120px"><col style="width:120px"><col style="width:80px">
</colgroup>
<thead><tr><th>#</th><th>参数名 (param_key)</th><th>模板单位</th><th>EDS 单位</th><th>结果</th></tr></thead>
<tbody>{unit_rows_html}</tbody>
</table>
</div>
</details>"""

    # ── Section 3: 数值比对 ──────────────────────────────────────────────────
    rows_html = []
    for i, r in enumerate(results):
        bg   = _STATUS_COLOR.get(r.status, "#fff")
        stat = _STATUS_LABEL.get(r.status, r.status)
        err  = ""
        if r.enip_error:   err += f"EIP: {html.escape(r.enip_error)}"
        if r.modbus_error: err += f"{'<br>' if err else ''}Modbus: {html.escape(r.modbus_error)}"
        rows_html.append(f"""
        <tr style="background:{bg}">
          <td class="num">{i+1}</td>
          <td class="key">{html.escape(r.param_key)}</td>
          <td class="val">{_fmt(r.enip_value)}</td>
          <td class="val">{_fmt(r.modbus_value)}</td>
          <td class="val">{_fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"}</td>
          <td class="val">{f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"}</td>
          <td class="stat">{stat}</td>
          <td class="err">{err}</td>
        </tr>""")

    val_badge = (f'<span class="badge err-badge">失败 {s["fail"]}</span>' if s["fail"] else "") + \
                (f'<span class="badge warn-badge">异常 {s["error"]}</span>' if s["error"] else "") + \
                ('<span class="badge ok-badge">全部通过</span>' if not s["fail"] and not s["error"] else "")

    val_section = f"""
<details open class="section">
<summary>三、数值比对（EtherNet/IP Assembly vs Modbus TCP）
  <span class="sum-info">共 {s['total']} 项 · 通过率 {s['pass_rate']}</span>
  {val_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{s['total']}</div><div class="lbl">总参数</div></div>
  <div class="card pass"> <div class="num">{s['pass']}</div><div class="lbl">通过 ({s['pass_rate']})</div></div>
  <div class="card fail"> <div class="num">{s['fail']}</div><div class="lbl">失败</div></div>
  <div class="card err">  <div class="num">{s['error']}</div><div class="lbl">异常</div></div>
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
    <th>EtherNet/IP 值</th><th>Modbus 值</th>
    <th>绝对差值</th><th>相对差值</th>
    <th>结果</th><th>错误信息</th>
  </tr>
</thead>
<tbody>{"".join(rows_html)}</tbody>
</table>
</div>
</details>"""

    # ── Section 4: Assembly 结构合规性 ──────────────────────────────────────────
    if integrity is not None:
        integ_ok    = integrity.ok
        integ_color = "#d4edda" if integ_ok else "#f8d7da"
        integ_badge = ('<span class="badge ok-badge">合规</span>' if integ_ok
                       else '<span class="badge err-badge">异常</span>')
        bytes_row_bg = "#d4edda" if integrity.bytes_match else "#f8d7da"
        oob_rows_html = "".join(
            f'<tr style="background:#f8d7da"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(item)}</td></tr>'
            for i, item in enumerate(integrity.out_of_bounds)
        )
        oob_html = (
            f'<h3 style="color:#721c24;margin-top:12px">'
            f'越界参数（{len(integrity.out_of_bounds)} 个）</h3>'
            f'<table><colgroup><col style="width:48px"><col></colgroup>'
            f'<thead><tr><th>#</th><th>参数（offset, size）</th></tr></thead>'
            f'<tbody>{oob_rows_html}</tbody></table>'
        ) if integrity.out_of_bounds else ""
        integrity_section = f"""
<details {"open " if not integ_ok else ""}class="section">
<summary>四、Assembly 结构合规性
  <span class="sum-info">参数数 {integrity.param_count} · EDS 声明 {integrity.eds_total_bytes} B</span>
  {integ_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{integrity.param_count}</div><div class="lbl">参数总数</div></div>
  <div class="card" style="background:{bytes_row_bg}">
    <div class="num" style="font-size:18px">{integrity.eds_total_bytes}</div>
    <div class="lbl">EDS 声明字节数</div></div>
  <div class="card" style="background:{bytes_row_bg}">
    <div class="num" style="font-size:18px">{integrity.actual_bytes}</div>
    <div class="lbl">实际读取字节数</div></div>
  <div class="card {'fail' if integrity.out_of_bounds else 'pass'}">
    <div class="num">{len(integrity.out_of_bounds)}</div><div class="lbl">越界参数</div></div>
  <div class="card" style="background:{integ_color}">
    <div class="num" style="font-size:18px">{"合规" if integ_ok else "异常"}</div>
    <div class="lbl">结论</div></div>
</div>
{oob_html}
</div>
</details>"""
    else:
        integrity_section = ""

    # ── Section 5: CIP Identity Object ─────────────────────────────────────────
    if identity is not None:
        def _opt_hex(v: Optional[int], w: int = 4) -> str:
            return f"0x{v:0{w}X} ({v})" if v is not None else "—"

        ident_ok    = identity.ok
        ident_color = "#d4edda" if ident_ok else "#f8d7da"
        ident_badge = ('<span class="badge ok-badge">读取成功</span>' if ident_ok
                       else '<span class="badge err-badge">读取失败</span>')
        dev_type_str = ""
        if identity.device_type is not None:
            dev_type_str = _DEVICE_TYPE_NAME.get(identity.device_type, "")
            dev_type_str = f" — {dev_type_str}" if dev_type_str else ""
        status_str = _decode_cip_status(identity.status) if identity.status is not None else "—"

        ident_rows = "".join(f"""
        <tr style="background:#fff">
          <td style="padding:5px 8px;font-weight:bold">{html.escape(label)}</td>
          <td style="padding:5px 8px;text-align:center;color:#999">{attr_no}</td>
          <td style="padding:5px 8px;font-family:monospace">{html.escape(value)}</td>
        </tr>""" for label, attr_no, value in [
            ("Vendor ID",    "1", _opt_hex(identity.vendor_id, 4)),
            ("Device Type",  "2", _opt_hex(identity.device_type, 4) + dev_type_str),
            ("Product Code", "3", _opt_hex(identity.product_code, 4)),
            ("Revision",     "4",
             f"{identity.revision_major}.{identity.revision_minor}"
             if identity.revision_major is not None else "—"),
            ("Status",       "5",
             f"0x{identity.status:04X} — {status_str}"
             if identity.status is not None else "—"),
            ("Serial Number","6", _opt_hex(identity.serial_number, 8)),
            ("Product Name", "7", identity.product_name or "—"),
        ])
        ident_err_html = (
            f'<p style="color:#721c24;margin-top:8px">读取错误：{html.escape(identity.error)}</p>'
            if not ident_ok else ""
        )
        identity_section = f"""
<details {"open " if not ident_ok else ""}class="section">
<summary>五、CIP Identity Object（Class=0x01, Instance=0x01）
  {ident_badge}
</summary>
<div class="section-body">
{ident_err_html}
<table style="max-width:700px">
<colgroup><col style="width:160px"><col style="width:48px"><col></colgroup>
<thead><tr><th>属性</th><th>Attr#</th><th>值</th></tr></thead>
<tbody>{ident_rows}</tbody>
</table>
</div>
</details>"""
    else:
        identity_section = ""

    # ── Section 5: CIP 错误响应测试 ──────────────────────────────────────────
    if err_tests:
        err_pass  = sum(1 for t in err_tests if t.passed)
        err_fail  = len(err_tests) - err_pass
        err_badge = (f'<span class="badge err-badge">失败 {err_fail}</span>' if err_fail
                     else '<span class="badge ok-badge">全部通过</span>')
        err_rows_html = "".join(f"""
        <tr style="background:{'#d4edda' if t.passed else '#f8d7da'}">
          <td class="num">{i+1}</td>
          <td style="padding:5px 8px;font-family:monospace;font-size:12px">{html.escape(t.name)}</td>
          <td style="padding:5px 8px;font-size:12px">{html.escape(t.description)}</td>
          <td class="stat">{"通过" if t.passed else "失败"}</td>
          <td class="err">{html.escape(t.detail)}</td>
        </tr>""" for i, t in enumerate(err_tests))
        err_section = f"""
<details {"open " if err_fail else ""}class="section">
<summary>六、CIP 错误响应测试
  <span class="sum-info">共 {len(err_tests)} 项 · 通过 {err_pass} / 失败 {err_fail}</span>
  {err_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(err_tests)}</div><div class="lbl">测试总数</div></div>
  <div class="card pass"> <div class="num">{err_pass}</div><div class="lbl">通过</div></div>
  <div class="card fail"> <div class="num">{err_fail}</div><div class="lbl">失败</div></div>
</div>
<table>
<colgroup>
  <col style="width:48px"><col style="width:90px"><col><col style="width:80px"><col style="width:280px">
</colgroup>
<thead><tr><th>#</th><th>测试 ID</th><th>描述</th><th>结果</th><th>实际响应</th></tr></thead>
<tbody>{err_rows_html}</tbody>
</table>
</div>
</details>"""
    else:
        err_section = ""

    # ── Section 7: 连接稳定性 ────────────────────────────────────────────────
    if stability is not None:
        stab_ok    = stability.ok
        stab_badge = ('<span class="badge ok-badge">稳定</span>' if stab_ok
                      else '<span class="badge err-badge">不稳定</span>')
        stab_rows_html = "".join(
            f'<tr style="background:#f8d7da"><td class="num">{i+1}</td>'
            f'<td class="err">{html.escape(e)}</td></tr>'
            for i, e in enumerate(stability.errors)
        )
        stab_err_table = (
            f'<table style="margin-top:8px"><colgroup><col style="width:48px"><col></colgroup>'
            f'<thead><tr><th>#</th><th>错误信息</th></tr></thead>'
            f'<tbody>{stab_rows_html}</tbody></table>'
        ) if stability.errors else ""
        stability_section = f"""
<details {"open " if not stab_ok else ""}class="section">
<summary>七、连接稳定性（连续 {stability.attempts} 次读取 Assembly）
  <span class="sum-info">{stability.successes}/{stability.attempts} 次成功</span>
  {stab_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{stability.attempts}</div><div class="lbl">测试次数</div></div>
  <div class="card pass"> <div class="num">{stability.successes}</div><div class="lbl">成功</div></div>
  <div class="card fail"> <div class="num">{stability.attempts - stability.successes}</div><div class="lbl">失败</div></div>
</div>
{stab_err_table}
</div>
</details>"""
    else:
        stability_section = ""

    content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>EtherNet/IP vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>
  body {{ font-family:"Microsoft YaHei",Arial,sans-serif; font-size:13px; margin:20px; background:#f5f5f5; color:#333; }}
  h1   {{ font-size:20px; margin-bottom:4px; }}
  .device-name {{ font-size:15px; font-weight:bold; color:#0056b3; margin-bottom:4px; }}
  .meta {{ color:#666; font-size:12px; margin-bottom:16px; }}
  .cards {{ display:flex; gap:12px; margin-bottom:16px; flex-wrap:wrap; }}
  .card {{ background:#fff; border-radius:6px; padding:14px 20px; box-shadow:0 1px 4px rgba(0,0,0,.1); min-width:110px; text-align:center; }}
  .card .num {{ font-size:28px; font-weight:bold; }}
  .card .lbl {{ font-size:11px; color:#888; margin-top:2px; }}
  .card.pass .num {{ color:#28a745; }} .card.fail .num {{ color:#dc3545; }}
  .card.err  .num {{ color:#ffc107; }} .card.total .num {{ color:#007bff; }}
  table {{ border-collapse:collapse; width:100%; background:#fff; box-shadow:0 1px 4px rgba(0,0,0,.1); border-radius:6px; overflow:hidden; table-layout:fixed; margin-bottom:24px; }}
  colgroup col.c-idx {{ width:48px; }} colgroup col.c-key {{ width:280px; }}
  colgroup col.c-val {{ width:110px; }} colgroup col.c-diff {{ width:90px; }}
  colgroup col.c-pct {{ width:80px; }} colgroup col.c-stat {{ width:90px; }}
  colgroup col.c-err {{ width:auto; }}
  thead tr {{ background:#343a40; color:#fff; }}
  th {{ padding:9px 8px; text-align:center; font-size:12px; white-space:nowrap; }}
  td {{ padding:5px 8px; border-bottom:1px solid #eee; vertical-align:middle; }}
  td.num  {{ text-align:right; color:#999; font-size:11px; }}
  td.key  {{ font-family:monospace; font-size:12px; word-break:break-all; }}
  td.val  {{ text-align:right; font-family:monospace; font-size:12px; }}
  td.stat {{ text-align:center; font-weight:bold; font-size:12px; }}
  td.err  {{ font-size:11px; color:#666; word-break:break-all; }}
  tr:hover td {{ filter:brightness(0.96); }}
  thead th {{ position:sticky; top:0; z-index:1; }}
  details.section {{ background:#fff; border-radius:8px; margin-bottom:20px; box-shadow:0 1px 4px rgba(0,0,0,.1); overflow:hidden; }}
  details.section > summary {{ list-style:none; cursor:pointer; padding:12px 16px;
    background:#f0f4f8; border-left:4px solid #0056b3;
    font-size:15px; font-weight:bold; color:#0056b3;
    display:flex; align-items:center; gap:8px; user-select:none; }}
  details.section > summary::-webkit-details-marker {{ display:none; }}
  details.section > summary::before {{ content:"▶"; font-size:10px; margin-right:4px; }}
  details[open].section > summary::before {{ content:"▼"; }}
  .sum-info {{ font-size:12px; font-weight:normal; color:#666; margin-left:4px; }}
  .badge {{ display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:bold; margin-left:4px; }}
  .ok-badge   {{ background:#d4edda; color:#155724; }}
  .err-badge  {{ background:#f8d7da; color:#721c24; }}
  .warn-badge {{ background:#fff3cd; color:#856404; }}
  .section-body {{ padding:16px; }}
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
<h1>EtherNet/IP vs Modbus TCP 比对报告</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  EtherNet/IP：{html.escape(config.ENIP_HOST)}:44818 &nbsp;|&nbsp;
  Modbus：{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT} &nbsp;|&nbsp;
  容差：±{config.ENIP_TOLERANCE_PERCENT}% / ±{config.ENIP_TOLERANCE_ABSOLUTE}
</div>
{scope_section}
{unit_section}
{val_section}
{integrity_section}
{identity_section}
{err_section}
{stability_section}
</body>
</html>"""

    Path(output_path).write_text(content, encoding="utf-8")
    log.info("HTML 报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 命令行入口
# ─────────────────────────────────────────────────────────────────────────────

async def _main(param_keys: Optional[list[str]], quick: bool) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    if quick and param_keys is None:
        from modbus_reader import get_param_map
        param_keys = list(get_param_map().keys())[:30]
        log.info("快速模式：仅比对前 %d 个参数", len(param_keys))

    if config.DEVICE_NAME in config.MODBUS_DEVICE_MAP:
        config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
            config.MODBUS_DEVICE_MAP[config.DEVICE_NAME]

    scope, unit_results, results, integrity = await run_comparison(param_keys)
    identity, err_tests, stability = await run_cip_checks()
    print_summary(scope, unit_results, results, integrity, identity, err_tests,
                  stability)
    report_path = generate_html_report(scope, unit_results, results,
                                       integrity=integrity, identity=identity,
                                       err_tests=err_tests, stability=stability)
    print(f"\n  HTML 报告：{report_path}\n")


if __name__ == "__main__":
    if "--device" in sys.argv:
        idx = sys.argv.index("--device")
        config.DEVICE_NAME   = sys.argv[idx + 1]
        config.DEVICE_MODULE = f"devices.{config.DEVICE_NAME.lower()}"

    quick = "--quick" in sys.argv
    if "--keys" in sys.argv:
        idx  = sys.argv.index("--keys")
        keys = sys.argv[idx + 1:]
    else:
        keys = None

    asyncio.run(_main(keys, quick))
