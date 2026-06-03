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
  python Protocols/EtherNetIP/enip_comparator.py --all          # 多表批量测试（需配置 ENIP_EDS_PATH + ENIP_MULTI_DEVICES）
  python Protocols/EtherNetIP/enip_comparator.py --all --quick  # 多表快速模式

多表模式说明：
  1. 在 config.ENIP_EDS_PATH 填写本次测试对应的网关 EDS 文件路径（每次更换 EDS 时更新）
  2. 在 config.ENIP_MULTI_DEVICES 填写各设备条目（格式见 config.py 注释）
  3. eds_label 必须与 EDS ConnectionN 的 help string 完全匹配，脚本据此自动定位该设备的 Assembly 实例
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
                          ConnectionManagerResult, ForwardOpenResult,
                          read_identity, test_error_responses,
                          check_assembly_integrity, check_connection_stability,
                          parse_connection_manager, parse_eds_file_revision,
                          parse_eds_device_section, parse_eds_assembly_declared_size,
                          check_eds_orphan_params, test_forward_open,
                          parse_device_assembly_map)
from modbus_reader import ModbusResult, get_reader
from template_reader import find_template_file, get_snmp_params

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

async def run_cip_checks() -> tuple[IdentityResult, list[ErrorTestResult], StabilityResult, ForwardOpenResult]:
    """独立连接，依次执行：Identity 读取、错误响应测试、连接稳定性、Forward_Open 隐式连接测试。"""
    from pycomm3 import CIPDriver

    def _sync() -> tuple[IdentityResult, list[ErrorTestResult], StabilityResult, ForwardOpenResult]:
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
            fw_open = test_forward_open(driver)
            log.info("Forward_Open 测试：%s  %s",
                     "成功" if fw_open.connected else "失败", fw_open.note)
            return ident, tests, stability, fw_open
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
    assem_ids:  Optional[list[int]] = None,
) -> tuple[ScopeReport, list[UnitCheckResult], list[CompareResult]]:
    t0 = time.time()

    device = config.DEVICE_NAME

    # ── 1. 模板 SNMP 范围 ────────────────────────────────────────────────────
    tmpl_path   = find_template_file(config.TEMPLATE_DIR, device)
    tmpl_params = get_snmp_params(tmpl_path)
    tmpl_map    = {p.param_key: p for p in tmpl_params}
    tmpl_set    = set(tmpl_map)

    # ── 2. EDS Assembly 范围 + 静态合规检查 ──────────────────────────────────
    # 优先使用显式 EDS 路径（多表模式），回退到按设备名模糊匹配（单表模式兼容）
    eds_path = (config.ENIP_EDS_PATH
                if getattr(config, "ENIP_EDS_PATH", "")
                else find_eds_file(config.EDS_DIR, device))
    eds_text   = Path(eds_path).read_text(encoding="utf-8", errors="replace")
    eds_params = parse_eds(eds_path, assem_ids)
    eds_map    = {p.name: p for p in eds_params}
    eds_set    = set(eds_map)
    eds_order  = list(dict.fromkeys(p.name for p in eds_params))   # 保持 Assembly 顺序，去重

    # 主 Assembly 实例 ID：多设备 EDS 中取当前设备的第一个 Assembly；单设备默认 10
    primary_assem_id = assem_ids[0] if assem_ids else 10

    # Connection Manager 静态检查（不依赖设备连接）
    cm_result          = parse_connection_manager(eds_text)
    eds_device         = parse_eds_device_section(eds_text)
    asm_declared_size  = parse_eds_assembly_declared_size(eds_text, assem_id=primary_assem_id)
    orphan_params      = check_eds_orphan_params(eds_text, assem_id=primary_assem_id)
    log.info("Connection Manager：字段数=%d  O->T=%s  问题数=%d",
             cm_result.field_count, cm_result.has_ot_direction, len(cm_result.issues))
    log.info("Assembly 声明字节=%d  孤儿Param=%d", asm_declared_size, len(orphan_params))

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
    async with EnipReader(config.ENIP_HOST, eds_path, config.ENIP_SLOT,
                          assem_ids=assem_ids) as enip, \
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

    # ── 7. Assembly 结构合规性（含 LREAL 对齐）────────────────────────────────
    integrity = check_assembly_integrity(eds_params, actual_bytes)
    log.info("Assembly 结构：EDS声明=%d B  实际=%d B  越界=%d 个  对齐问题=%d 个",
             integrity.eds_total_bytes, integrity.actual_bytes,
             len(integrity.out_of_bounds), len(integrity.alignment_issues))

    log.info("比对完成，耗时 %.1f 秒", time.time() - t0)
    return scope_report, unit_results, results, integrity, cm_result, \
           eds_device, asm_declared_size, orphan_params


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
    cm_result: Optional[ConnectionManagerResult] = None,
    fw_open: Optional[ForwardOpenResult] = None,
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
              f"实际={integrity.actual_bytes}B  越界={len(integrity.out_of_bounds)}  "
              f"对齐问题={len(integrity.alignment_issues)}")
        for item in integrity.out_of_bounds[:3]:
            print(f"    越界: {item}")
        for item in integrity.alignment_issues[:3]:
            print(f"    对齐: {item}")
    if cm_result is not None:
        status_str = "OK" if cm_result.ok else "FAIL"
        print(f"  【EDS Connection Manager】{status_str}  字段数={cm_result.field_count}  "
              f"O->T={'有' if cm_result.has_ot_direction else '缺失'}  "
              f"Revision={cm_result.eds_revision}  问题={len(cm_result.issues)}")
        for issue in cm_result.issues:
            print(f"    !! {issue}")
    if fw_open is not None:
        if not fw_open.attempted:
            print("  【Forward_Open 隐式连接】未测试")
        else:
            status_str = "OK" if fw_open.connected else "FAIL"
            print(f"  【Forward_Open 隐式连接】{status_str}  {fw_open.service_used}")
            if fw_open.connected:
                print(f"    T->O_API={fw_open.t_o_api_us}us  {fw_open.note}")
            else:
                print(f"    {fw_open.note}")
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
    cm_result: Optional[ConnectionManagerResult] = None,
    fw_open: Optional[ForwardOpenResult] = None,
    eds_device: Optional[dict] = None,
    asm_declared_size: int = -1,
    orphan_params: Optional[list[str]] = None,
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

    # ── Section 8: EDS 静态合规性检查 ────────────────────────────────────────
    if cm_result is not None or (integrity is not None and integrity.alignment_issues):
        eds_issues: list[str] = []

        # Connection Manager 问题
        cm_issues_html = ""
        if cm_result is not None:
            for iss in cm_result.issues:
                eds_issues.append(iss)
            cm_rows = "".join(
                f'<tr style="background:#f8d7da"><td class="num">{i+1}</td>'
                f'<td style="padding:5px 8px;font-size:12px">{html.escape(iss)}</td></tr>'
                for i, iss in enumerate(cm_result.issues)
            ) if cm_result.issues else (
                '<tr style="background:#d4edda"><td colspan="2" style="padding:5px 8px;text-align:center">'
                '无问题</td></tr>'
            )
            has_ot_str  = '有' if cm_result.has_ot_direction else '<span style="color:#dc3545;font-weight:bold">缺失</span>'
            cm_name_str = html.escape(cm_result.connection_name) if cm_result.connection_name else '<span style="color:#dc3545">（空）</span>'
            cm_path_str = html.escape(cm_result.path) if cm_result.path else '<span style="color:#dc3545">（空）</span>'
            # ── Connection1 原始字段明细表（兜底人工审查，覆盖规则未知的问题）──
            field_rows_html = ""
            if cm_result.raw_fields:
                flagged_positions = cm_result.CHECKED_POSITIONS
                issued_positions: set[int] = set()
                # 判断哪些位置有规则检出问题（用于标红）
                if cm_result.issues:
                    for idx, iss_text in [(7, "T→O Format"), (9, "Proxy Config Format"),
                                          (10, "Target Config Size"), (11, "Target Config Format"),
                                          (12, "Connection Name"), (14, "Path")]:
                        if any(iss_text in iss for iss in cm_result.issues):
                            issued_positions.add(idx)
                        if not cm_result.raw_fields[idx] and idx in flagged_positions:
                            issued_positions.add(idx)

                # 已知"空值合规"的位置及说明（ODVA 规范许可）
                _OK_EMPTY: dict[int, str] = {
                    5:  "空=无 O→T 方向（T→O 方向缺失的伴随症状，已由 Error #1 + 双向检查覆盖）",
                    6:  "空=无 O→T 方向（同上）",
                    8:  "空=无代理配置（Proxy Config Size=0，非模块化设备合规）",
                    13: "可选字段，空值合规",
                }

                for i, (fname, fval) in enumerate(
                    zip(cm_result.FIELD_NAMES, cm_result.raw_fields)
                ):
                    is_flagged  = i in issued_positions
                    is_checked  = i in flagged_positions
                    is_missing  = (fval == '' and i >= cm_result.field_count)
                    is_ok_empty = (not fval and i in _OK_EMPTY)
                    bg = ("#f8d7da" if is_flagged
                          else "#fff3cd" if is_checked and not fval
                          else "#fff")
                    val_display = html.escape(fval) if fval else (
                        '<span style="color:#aaa;font-style:italic">（缺失）</span>'
                        if is_missing else
                        '<span style="color:#aaa;font-style:italic">（空）</span>'
                    )
                    if is_flagged:
                        suffix = ' <span style="color:#dc3545;font-size:11px">← 规则检出</span>'
                    elif is_ok_empty:
                        suffix = (f' <span style="color:#28a745;font-size:11px">'
                                  f'← {_OK_EMPTY[i]}</span>')
                    elif is_checked:
                        suffix = ' <span style="color:#856404;font-size:11px">← 规则覆盖</span>'
                    else:
                        suffix = ''
                    field_rows_html += (
                        f'<tr style="background:{bg}">'
                        f'<td class="num">{i+1}</td>'
                        f'<td style="padding:4px 8px;font-size:12px">{html.escape(fname)}</td>'
                        f'<td style="padding:4px 8px;font-family:monospace;font-size:12px">'
                        f'{val_display}{suffix}</td>'
                        f'</tr>'
                    )

            raw_table_html = f"""
<details style="margin-top:12px">
<summary style="cursor:pointer;font-size:13px;color:#555;padding:4px 0">
  Connection1 原始字段明细（15 个 ODVA 规范位置，供人工核查规则以外的问题）
</summary>
<table style="max-width:800px;margin-top:6px">
<colgroup><col style="width:36px"><col style="width:200px"><col></colgroup>
<thead><tr><th>Pos</th><th>字段名</th><th>实际值</th></tr></thead>
<tbody>{field_rows_html}</tbody>
</table>
<p style="font-size:11px;color:#888;margin:4px 0 0">
  红色=规则检出问题 · 黄色=规则覆盖位置但值为空 · <span style="color:#28a745">绿色注释=空值合规（ODVA 规范许可）</span> · 无色=规则未覆盖位置（需人工判断）
</p>
</details>""" if cm_result.raw_fields else ""

            cm_issues_html = f"""
<h3 style="margin:12px 0 6px">Connection Manager — Connection1 字段检查</h3>
<table style="max-width:700px;margin-bottom:12px">
<colgroup><col style="width:160px"><col></colgroup>
<tbody>
  <tr><td style="padding:4px 8px;font-weight:bold">字段数</td>
      <td style="padding:4px 8px;font-family:monospace">{cm_result.field_count}
        <span style="color:#888;font-size:11px">（双向期望15 / 仅T→O期望12）</span></td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">O→T 方向</td>
      <td style="padding:4px 8px">{has_ot_str}</td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">Connection Name</td>
      <td style="padding:4px 8px;font-family:monospace">{cm_name_str}</td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">Path</td>
      <td style="padding:4px 8px;font-family:monospace">{cm_path_str}</td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">EDS Revision</td>
      <td style="padding:4px 8px;font-family:monospace">{html.escape(cm_result.eds_revision) or "—"}</td></tr>
</tbody>
</table>
<table style="max-width:700px;margin-bottom:8px">
<colgroup><col style="width:48px"><col></colgroup>
<thead><tr><th>#</th><th>问题描述（规则自动检出）</th></tr></thead>
<tbody>{cm_rows}</tbody>
</table>
{raw_table_html}"""

        # LREAL 对齐问题
        align_issues_html = ""
        if integrity is not None and integrity.alignment_issues:
            for iss in integrity.alignment_issues:
                eds_issues.append(f"对齐: {iss}")
            align_rows = "".join(
                f'<tr style="background:#fff3cd"><td class="num">{i+1}</td>'
                f'<td style="padding:5px 8px;font-family:monospace;font-size:12px">'
                f'{html.escape(iss)}</td></tr>'
                for i, iss in enumerate(integrity.alignment_issues)
            )
            align_issues_html = f"""
<h3 style="margin:12px 0 6px">Assembly 数据类型对齐检查</h3>
<p style="font-size:12px;color:#666;margin:0 0 8px">
  Rockwell Logix AOP 要求 LREAL(8B)/DINT/UDINT(4B) 参数的字节偏移满足对应对齐要求。
  偏移不对齐不影响显式消息读取，但会导致 PLC 内部映射异常。
</p>
<table style="max-width:900px">
<colgroup><col style="width:48px"><col></colgroup>
<thead><tr><th>#</th><th>参数（dtype, offset, offset%align）</th></tr></thead>
<tbody>{align_rows}</tbody>
</table>"""

        # EDS Revision vs Identity Revision 交叉比对
        rev_html = ""
        if cm_result is not None and identity is not None and identity.ok and cm_result.eds_revision:
            identity_rev = f"{identity.revision_major}.{identity.revision_minor}"
            rev_match = (cm_result.eds_revision == identity_rev or
                         cm_result.eds_revision.lstrip('0').rstrip('0').rstrip('.') ==
                         identity_rev.lstrip('0').rstrip('0').rstrip('.'))
            rev_bg  = "#d4edda" if rev_match else "#f8d7da"
            rev_str = "一致" if rev_match else "不一致"
            if not rev_match:
                eds_issues.append(
                    f"EDS Revision({cm_result.eds_revision}) 与 "
                    f"Identity Revision({identity_rev}) 不一致"
                )
            rev_html = f"""
<h3 style="margin:12px 0 6px">EDS Revision vs CIP Identity Revision 交叉比对</h3>
<table style="max-width:500px">
<colgroup><col style="width:160px"><col style="width:120px"><col style="width:80px"></colgroup>
<thead><tr><th>来源</th><th>Revision</th><th>结论</th></tr></thead>
<tbody>
  <tr style="background:{rev_bg}">
    <td style="padding:5px 8px">EDS [File]</td>
    <td style="padding:5px 8px;font-family:monospace">{html.escape(cm_result.eds_revision)}</td>
    <td style="padding:5px 8px;text-align:center;font-weight:bold" rowspan="2">{rev_str}</td>
  </tr>
  <tr style="background:{rev_bg}">
    <td style="padding:5px 8px">CIP Identity (Attr 4)</td>
    <td style="padding:5px 8px;font-family:monospace">{html.escape(identity_rev)}</td>
  </tr>
</tbody>
</table>"""

        # ── [Device] section vs CIP Identity 交叉比对 ────────────────────────
        # 严重级别：critical=影响设备识别（不一致→Error）
        #           warning=命名/描述差异（不一致→Warning，不计入 eds_issues）
        dev_html = ""
        if eds_device and identity is not None and identity.ok:
            dev_checks = [
                # (eds_key, eds_val, ident_val, label, is_critical)
                ("VendCode", str(eds_device.get("VendCode", "")),
                 str(identity.vendor_id or ""),      "Vendor ID",        True),
                ("ProdCode", str(eds_device.get("ProdCode", "")),
                 str(identity.product_code or ""),   "Product Code",     True),
                ("MajRev",   str(eds_device.get("MajRev", "")),
                 str(identity.revision_major or ""), "Revision Major",   True),
                ("MinRev",   str(eds_device.get("MinRev", "")),
                 str(identity.revision_minor or ""), "Revision Minor",   True),
                ("ProdName", eds_device.get("ProdName", ""),
                 identity.product_name or "",         "Product Name",     False),
                # 扩展：未来可继续在此追加 Warning 级字段，如 VendName、Catalog
            ]
            dev_rows_html = ""
            for eds_key, eds_val, ident_val, label, is_critical in dev_checks:
                match = (eds_val == ident_val)
                if not match and is_critical:
                    # 关键字段不一致：计入 eds_issues，红色
                    eds_issues.append(
                        f"[Device] {eds_key}={eds_val!r} 与 Identity {label}={ident_val!r} 不一致"
                    )
                    bg   = "#f8d7da"
                    icon = '不一致'
                    note = ''
                elif not match:
                    # 非关键字段不一致：仅显示为警告，不计入 eds_issues，黄色
                    bg   = "#fff3cd"
                    icon = '差异'
                    note = (' <span style="color:#856404;font-size:11px">'
                            '（命名差异，不影响 PLC 识别，无需修复）</span>')
                else:
                    bg   = "#d4edda"
                    icon = '一致'
                    note = ''
                level_badge = (
                    '<span style="font-size:10px;color:#721c24;background:#f8d7da;'
                    'padding:1px 4px;border-radius:3px">关键</span>'
                    if is_critical else
                    '<span style="font-size:10px;color:#856404;background:#fff3cd;'
                    'padding:1px 4px;border-radius:3px">参考</span>'
                )
                dev_rows_html += (
                    f'<tr style="background:{bg}">'
                    f'<td style="padding:4px 8px">{html.escape(label)}&nbsp;{level_badge}</td>'
                    f'<td style="padding:4px 8px;font-family:monospace">{html.escape(eds_key)}</td>'
                    f'<td style="padding:4px 8px;font-family:monospace">{html.escape(eds_val)}</td>'
                    f'<td style="padding:4px 8px;font-family:monospace">'
                    f'{html.escape(ident_val)}{note}</td>'
                    f'<td style="padding:4px 8px;text-align:center;font-weight:bold">{icon}</td>'
                    f'</tr>'
                )
            dev_html = f"""
<h3 style="margin:12px 0 6px">[Device] 段 vs CIP Identity 交叉比对</h3>
<p style="font-size:12px;color:#666;margin:0 0 6px">
  <b>关键</b>字段不一致计入问题数（PLC 用 VendCode+ProdCode 识别设备）；
  <b>参考</b>字段差异仅展示（不影响连接建立，无需修复）。
</p>
<table style="max-width:900px;margin-bottom:16px">
<colgroup><col style="width:180px"><col style="width:120px"><col><col><col style="width:70px"></colgroup>
<thead><tr><th>属性</th><th>EDS 字段</th><th>EDS 值</th><th>Identity 值</th><th>结论</th></tr></thead>
<tbody>{dev_rows_html}</tbody>
</table>"""

        # ── Assembly 字节数三路一致性 ──────────────────────────────────────
        asm_html = ""
        if integrity is not None and (asm_declared_size >= 0 or
                                       cm_result is not None and cm_result.to_size_declared >= 0):
            # 三路一致性针对当前设备的主 Assembly（Assembly ID 最小者），取该组 computed 字节数
            _primary = min(integrity.per_assembly_sizes) if integrity.per_assembly_sizes else 10
            computed   = integrity.per_assembly_sizes.get(_primary, integrity.eds_total_bytes)
            asm_decl   = asm_declared_size      # [Assembly] Assem10 头部声明
            conn_decl  = cm_result.to_size_declared if cm_result else -1  # Connection1 T→O Size

            size_rows  = ""
            all_match  = True
            for label, val in [
                (f"[Assembly] Assem{_primary} 头部声明（bytes）", asm_decl),
                ("[Params] 成员大小之和（computed）",               computed),
                ("Connection1 T→O Size 字段",                       conn_decl),
            ]:
                if val < 0:
                    disp = "— （未声明）"
                    bg   = "#f5f5f5"
                else:
                    match = (val == computed)
                    if not match:
                        all_match = False
                    bg   = "#d4edda" if match else "#f8d7da"
                    disp = str(val)
                size_rows += (
                    f'<tr style="background:{bg}">'
                    f'<td style="padding:4px 8px">{html.escape(label)}</td>'
                    f'<td style="padding:4px 8px;font-family:monospace;text-align:right">{disp}</td>'
                    f'</tr>'
                )

            if not all_match:
                eds_issues.append(
                    f"Assembly 字节数三路不一致：Assem10声明={asm_decl}  "
                    f"computed={computed}  Connection1_T→O_Size={conn_decl}"
                )
            status_str = "一致" if all_match else "不一致"
            asm_html = f"""
<h3 style="margin:12px 0 6px">Assembly 字节数三路一致性（{status_str}）</h3>
<p style="font-size:12px;color:#666;margin:0 0 6px">
  三处数值应相同：[Assembly] 头部声明 = [Params] 成员大小之和 = Connection1 T→O Size
</p>
<table style="max-width:600px;margin-bottom:16px">
<colgroup><col><col style="width:100px"></colgroup>
<thead><tr><th>来源</th><th>字节数</th></tr></thead>
<tbody>{size_rows}</tbody>
</table>"""

        # ── 孤儿 Param 引用检查 ──────────────────────────────────────────
        orphan_html = ""
        if orphan_params:
            for p in orphan_params:
                eds_issues.append(f"孤儿参数引用：{p} 在 [Assembly] 中被引用但 [Params] 中未定义")
            orphan_rows = "".join(
                f'<tr style="background:#f8d7da"><td class="num">{i+1}</td>'
                f'<td style="padding:5px 8px;font-family:monospace">{html.escape(p)}</td></tr>'
                for i, p in enumerate(orphan_params)
            )
            orphan_html = f"""
<h3 style="margin:12px 0 6px">孤儿 Param 引用（{len(orphan_params)} 个）</h3>
<p style="font-size:12px;color:#666;margin:0 0 6px">
  [Assembly] Assem10 中引用但 [Params] 段未定义，数值解析时将使用默认 float32 类型，可能导致数据误读。
</p>
<table style="max-width:500px">
<colgroup><col style="width:48px"><col></colgroup>
<thead><tr><th>#</th><th>Param 引用</th></tr></thead>
<tbody>{orphan_rows}</tbody>
</table>"""

        total_issues = len(eds_issues)
        eds_badge = (f'<span class="badge err-badge">问题 {total_issues}</span>'
                     if total_issues else '<span class="badge ok-badge">无问题</span>')
        eds_static_section = f"""
<details {"open " if total_issues else ""}class="section">
<summary>八、EDS 文件静态合规性检查（无需 PLC）
  <span class="sum-info">Connection Manager · LREAL 对齐 · Revision 一致性</span>
  {eds_badge}
</summary>
<div class="section-body">
<p style="font-size:12px;color:#666;margin:0 0 12px">
  以下检查均基于 EDS 文件静态解析，不依赖设备连接，可发现 PLC（如 Studio 5000）集成前的兼容性问题。
</p>
{cm_issues_html}{align_issues_html}{rev_html}{dev_html}{asm_html}{orphan_html}
</div>
</details>"""
    else:
        eds_static_section = ""

    # ── Section 9: Forward_Open 隐式连接建立 ────────────────────────────────────
    if fw_open is not None and fw_open.attempted:
        fw_ok    = fw_open.connected
        fw_color = "#d4edda" if fw_ok else "#f8d7da"
        fw_label = "成功" if fw_ok else "失败"
        fw_badge = ('<span class="badge ok-badge">连接成功</span>' if fw_ok
                    else '<span class="badge err-badge">连接失败</span>')

        if fw_ok:
            detail_rows = f"""
  <tr><td style="padding:4px 8px;font-weight:bold">T→O 实际包间隔</td>
      <td style="padding:4px 8px;font-family:monospace">{fw_open.t_o_api_us} μs
        <span style="color:#888;font-size:11px">（= {fw_open.t_o_api_us/1000:.1f} ms）</span></td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">详情</td>
      <td style="padding:4px 8px;font-size:12px">{html.escape(fw_open.note)}</td></tr>"""
        else:
            detail_rows = f"""
  <tr style="background:#f8d7da">
    <td style="padding:4px 8px;font-weight:bold">失败原因</td>
    <td style="padding:4px 8px;font-size:12px;word-break:break-all">{html.escape(fw_open.note)}</td>
  </tr>"""

        forward_open_section = f"""
<details {"open " if not fw_ok else ""}class="section">
<summary>九、Forward_Open 隐式连接建立（{fw_open.service_used}）
  {fw_badge}
</summary>
<div class="section-body">
<p style="font-size:12px;color:#666;margin:0 0 12px">
  使用 Large_Forward_Open (0x5B) 以 Input-Only 参数（O→T null / T→O P2P 5656B 500ms）
  向 Assembly 10 发起隐式连接，成功后立即 Forward_Close。<br>
  此测试弥补显式消息绕过 Connection Manager 的盲区：
  若 EDS Connection1 不合规，设备将在此步骤返回 CIP 错误而非在数值比对中体现。
</p>
<table style="max-width:600px">
<colgroup><col style="width:180px"><col></colgroup>
<tbody>
  <tr style="background:{fw_color}">
    <td style="padding:4px 8px;font-weight:bold">结论</td>
    <td style="padding:4px 8px;font-weight:bold">{fw_label}</td>
  </tr>
  <tr><td style="padding:4px 8px;font-weight:bold">使用服务</td>
      <td style="padding:4px 8px;font-family:monospace">{html.escape(fw_open.service_used)}</td></tr>
  <tr><td style="padding:4px 8px;font-weight:bold">连接参数</td>
      <td style="padding:4px 8px;font-size:12px">O→T: null | T→O: P2P, fixed, 5656B, RPI=500ms</td></tr>
{detail_rows}
</tbody>
</table>
</div>
</details>"""
    else:
        forward_open_section = ""

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
{eds_static_section}
{forward_open_section}
</body>
</html>"""

    Path(output_path).write_text(content, encoding="utf-8")
    log.info("HTML 报告已保存：%s", output_path)
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# 多表批量测试
# ─────────────────────────────────────────────────────────────────────────────

async def run_all_devices(
    param_keys: Optional[list[str]] = None,
    quick: bool = False,
) -> None:
    """遍历 config.ENIP_MULTI_DEVICES，逐台设备执行完整比对并生成独立 HTML 报告。

    流程：
      1. 读取 config.ENIP_EDS_PATH（网关配置快照 EDS），解析设备→Assembly 映射
      2. 对 ENIP_MULTI_DEVICES 中每条记录，按 eds_label 匹配 EDS help string 找到该设备的
         Assembly 实例列表，仅解析/读取这些 Assembly（避免多设备参数混用）
      3. 各设备独立生成 HTML 报告，完成后打印多表汇总

    config.ENIP_EDS_PATH 必须填写（指向本次测试对应的网关 EDS 文件）。
    config.ENIP_MULTI_DEVICES 格式：
        (eds_label, device_name, modbus_host, modbus_port, modbus_unit)
    """
    import modbus_reader as _mr

    if not config.ENIP_MULTI_DEVICES:
        log.warning("config.ENIP_MULTI_DEVICES 为空，请填写后再使用 --all")
        return

    eds_path = getattr(config, "ENIP_EDS_PATH", "")
    if not eds_path:
        log.error("--all 模式需要在 config.ENIP_EDS_PATH 中指定 EDS 文件路径")
        return

    # 解析 EDS，建立 help_string → Assembly 实例 ID 列表的映射
    eds_text         = Path(eds_path).read_text(encoding="utf-8", errors="replace")
    device_assem_map = parse_device_assembly_map(eds_text)
    log.info("EDS 设备→Assembly 映射：%s", device_assem_map)

    results_summary: list[dict] = []
    prev_device_module = ""

    for entry in config.ENIP_MULTI_DEVICES:
        eds_label, device_name, mb_host, mb_port, mb_unit = entry
        log.info("═" * 60)
        log.info("开始测试：%s (%s)  Modbus=%s:%d Unit=%d",
                 eds_label, device_name, mb_host, mb_port, mb_unit)

        # 按 help string 查找该设备的 Assembly 实例列表
        assem_ids = device_assem_map.get(eds_label)
        if assem_ids is None:
            msg = f"EDS 中未找到 help string={eds_label!r} 对应的 Connection 条目"
            log.error(msg)
            results_summary.append({"label": eds_label, "device": device_name,
                                     "ok": False, "error_msg": msg})
            continue
        log.info("匹配 Assembly 实例：%s", assem_ids)

        # 动态切换 Modbus 连接配置（EIP 网关统一使用 config.ENIP_HOST）
        config.DEVICE_NAME   = device_name
        config.DEVICE_MODULE = f"devices.{device_name.lower()}"
        config.MODBUS_HOST   = mb_host
        config.MODBUS_PORT   = mb_port
        config.MODBUS_UNIT   = mb_unit

        # 设备类型变更时重置 Modbus 参数缓存
        if config.DEVICE_MODULE != prev_device_module:
            _mr._PARAM_MAP = None
            prev_device_module = config.DEVICE_MODULE

        keys = param_keys
        if quick and keys is None:
            from modbus_reader import get_param_map
            keys = list(get_param_map().keys())[:30]
            log.info("快速模式：仅比对前 %d 个参数", len(keys))

        try:
            scope, unit_results, compare_results, integrity, cm_result, \
                eds_device, asm_declared_size, orphan_params = \
                await run_comparison(keys, assem_ids=assem_ids)
            identity, err_tests, stability, fw_open = await run_cip_checks()

            print_summary(scope, unit_results, compare_results, integrity,
                          identity, err_tests, stability, cm_result, fw_open)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_dir = Path(config.REPORT_DIR)
            report_dir.mkdir(parents=True, exist_ok=True)
            safe_label = eds_label.replace("/", "_").replace(" ", "_")
            report_path = str(report_dir / f"enip_compare_{safe_label}_{ts}.html")
            generate_html_report(
                scope, unit_results, compare_results,
                output_path=report_path,
                integrity=integrity, identity=identity,
                err_tests=err_tests, stability=stability,
                cm_result=cm_result, fw_open=fw_open,
                eds_device=eds_device,
                asm_declared_size=asm_declared_size,
                orphan_params=orphan_params,
            )
            s = summary(compare_results)
            results_summary.append({
                "label": eds_label, "device": device_name, "ok": True,
                "pass": s["pass"], "fail": s["fail"], "error": s["error"],
                "total": s["total"], "pass_rate": s["pass_rate"],
                "report": report_path,
            })
            print(f"\n  HTML 报告：{report_path}\n")

        except Exception as exc:
            log.error("设备 %s 测试失败：%s", eds_label, exc, exc_info=True)
            results_summary.append({
                "label": eds_label, "device": device_name,
                "ok": False, "error_msg": str(exc),
            })

    # 多表汇总
    print("\n" + "═" * 70)
    print("  多表 EtherNet/IP 比对汇总")
    print("═" * 70)
    for r in results_summary:
        if r["ok"]:
            print(
                f"  {r['label']:20s} ({r['device']:15s})  "
                f"PASS={r['pass']:4d}  FAIL={r['fail']:4d}  ERR={r['error']:4d}  "
                f"通过率={r['pass_rate']}"
            )
            print(f"    报告：{r['report']}")
        else:
            print(f"  {r['label']:20s} ({r['device']:15s})  !! 测试异常：{r['error_msg']}")
    print("═" * 70)


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

    scope, unit_results, results, integrity, cm_result, \
        eds_device, asm_declared_size, orphan_params = await run_comparison(param_keys)
    identity, err_tests, stability, fw_open = await run_cip_checks()
    print_summary(scope, unit_results, results, integrity, identity, err_tests,
                  stability, cm_result, fw_open)
    report_path = generate_html_report(scope, unit_results, results,
                                       integrity=integrity, identity=identity,
                                       err_tests=err_tests, stability=stability,
                                       cm_result=cm_result, fw_open=fw_open,
                                       eds_device=eds_device,
                                       asm_declared_size=asm_declared_size,
                                       orphan_params=orphan_params)
    print(f"\n  HTML 报告：{report_path}\n")


if __name__ == "__main__":
    if "--device" in sys.argv:
        idx = sys.argv.index("--device")
        config.DEVICE_NAME   = sys.argv[idx + 1]
        config.DEVICE_MODULE = f"devices.{config.DEVICE_NAME.lower()}"

    quick   = "--quick" in sys.argv
    run_all = "--all"   in sys.argv

    if "--keys" in sys.argv:
        idx  = sys.argv.index("--keys")
        keys = sys.argv[idx + 1:]
    else:
        keys = None

    if run_all:
        asyncio.run(run_all_devices(keys, quick))
    else:
        asyncio.run(_main(keys, quick))
