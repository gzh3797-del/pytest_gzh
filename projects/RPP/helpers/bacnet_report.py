# -*- coding: utf-8 -*-
"""
bacnet_report.py — 六段式 BACnet vs Modbus 比对 HTML 报告生成器

与 tools/Protocols/BACnetIP/comparator.py 的 generate_html_report 视觉一致，
但输入改为 AcuHMI-1-7 UI 用例（test_012~017 的 _verify_full_param_upload）原生
数据结构，无需依赖 tools/Protocols 的 config。每台被测设备各生成一份独立报告，
供 BACnet/IP 参数比对用例（TestCase_AcuHMI-1-7_033_001_012~017）执行时落盘分析。

六段：
  一、参数范围检查（模板 vs 网关实际发布）
  二、单位检查（BACnet units 属性 vs 模板）
  三、数值比对（BACnet Present Value vs Modbus）
  四、BACnet Device Object 属性（§12.11）
  五、协议合规性测试（§16 + AI 必需属性）
  六、连接稳定性测试
"""
from __future__ import annotations

import html
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Optional

from projects.RPP.helpers.hmi_bacnet_client import (
    DeviceInfoResult,
    MetadataItem,
    ProtocolCheckItem,
    StabilityCheckResult,
)

log = logging.getLogger(__name__)

_REPO_ROOT = Path(__file__).resolve().parents[3]

# cmp_rows 的 status → (背景色, 中文标签)
_STATUS_COLOR = {
    "PASS": "#d4edda",
    "FAIL": "#f8d7da",
    "ERR":  "#fff3cd",
}
_STATUS_LABEL = {
    "PASS": "通过",
    "FAIL": "失败",
    "ERR":  "异常",
}

# 数值比对单行类型别名：(param_key, bacnet值, modbus值, 绝对差, 相对差%, status)
CmpRow = tuple[str, Optional[float], Optional[float], Optional[float], Optional[float], str]


def get_bacnet_report_dir() -> Path:
    """BACnet 比对报告输出目录（惰性 mkdir）。

    与本次运行的 pytest-html 报告（bacnet_ui.html）落在**同一目录**：
    优先用根 conftest / framework runner 注入的环境变量 REPORT_DIR
    （指向 reports/<项目>/<模块>/<时间戳>/，即 bacnet_ui.html 所在的 run 根目录），
    报告直接写在该目录下，与 bacnet_ui.html 同级。

    未注入 REPORT_DIR 时（如裸跑未传 --html），回退到
    reports/RPP/BacnetIP/<时间戳>/，结构与正式运行保持一致；
    时间戳复用 RUN_TS（若有），保证同次运行各设备报告同目录。
    """
    report_dir = os.getenv("REPORT_DIR")
    if report_dir:
        base = Path(report_dir)
    else:
        ts = os.getenv("RUN_TS") or datetime.now().strftime("%Y%m%d_%H%M%S")
        base = _REPO_ROOT / "reports" / "RPP" / "BacnetIP" / ts
    base.mkdir(parents=True, exist_ok=True)
    return base


def _fmt(v: Optional[float], digits: int = 6) -> str:
    if v is None:
        return "—"
    s = f"{v:.{digits}f}"
    return s.rstrip("0").rstrip(".") if "." in s else s


def _cmp_summary(rows: list[CmpRow]) -> dict:
    """数值比对统计（参考 comparator.summary）。"""
    total = len(rows)
    passed = sum(1 for r in rows if r[5] == "PASS")
    failed = sum(1 for r in rows if r[5] == "FAIL")
    errors = total - passed - failed
    worst = sorted(
        (r for r in rows if r[5] == "FAIL"),
        key=lambda r: r[4] or 0,
        reverse=True,
    )
    return {
        "total": total,
        "pass": passed,
        "fail": failed,
        "error": errors,
        "pass_rate": f"{passed / total * 100:.1f}%" if total else "N/A",
        "worst_fails": worst[:10],
    }


def generate_six_segment_report(
    *,
    device_name: str,
    template_name: str,
    template_keys: set[str],
    published_keys: set[str],
    matched: list[str],
    missing: list[str],
    extra: list[str],
    meta_results: list[MetadataItem],
    cmp_rows: list[CmpRow],
    cmp_note: str,
    dev_info: Optional[DeviceInfoResult],
    compliance_results: list[ProtocolCheckItem],
    stability: Optional[StabilityCheckResult],
    gateway_ip: str = "",
    gateway_port: object = "",
    tol_pct: float = 1.0,
    tol_abs: float = 0.05,
    output_path: Optional[str] = None,
) -> str:
    """生成六段式 HTML 比对报告，返回文件路径。

    输入对应 _verify_full_param_upload 各段已算好的数据，本函数只做渲染，不做断言。
    """
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_dev = "".join(c if c.isalnum() or c in "-_" else "_" for c in device_name)
        output_path = str(get_bacnet_report_dir() / f"bacnet_{safe_dev}_{ts}.html")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ── Section 1: 参数范围检查 ───────────────────────────────────────────────
    scope_ok = not missing and not extra
    scope_color = "#d4edda" if scope_ok else "#f8d7da"
    scope_label = "一致" if scope_ok else "不一致"

    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i + 1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    missing_html = (
        f'<h3 style="color:#721c24;margin-top:12px">模板有但网关未发布（{len(missing)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(missing, "#f8d7da")}</tbody></table>'
    ) if missing else ""

    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">网关发布但模板未包含（{len(extra)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(extra, "#fff3cd")}</tbody></table>'
    ) if extra else ""

    scope_badge = (f'<span class="badge err-badge">缺失 {len(missing)}</span>'
                   if missing else "") + \
                  (f'<span class="badge warn-badge">多余 {len(extra)}</span>'
                   if extra else "") + \
                  ('<span class="badge ok-badge">一致</span>' if scope_ok else "")
    scope_html = f"""
<details open class="section">
<summary>一、参数范围检查
  <span class="sum-info">模板 {len(template_keys)} / 网关 {len(published_keys)} / 匹配 {len(matched)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(template_keys)}</div><div class="lbl">模板 BACnet 参数</div></div>
  <div class="card total"><div class="num">{len(published_keys)}</div><div class="lbl">网关实际发布</div></div>
  <div class="card pass"> <div class="num">{len(matched)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(missing)}</div><div class="lbl">模板有/网关缺</div></div>
  <div class="card err">  <div class="num">{len(extra)}</div><div class="lbl">网关多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 单位检查 ───────────────────────────────────────────────────
    if meta_results:
        meta_read_failed = [m for m in meta_results
                            if getattr(m, "unit_read_failed", False)]
        meta_skipped = [m for m in meta_results
                        if m.unit_skipped and not getattr(m, "unit_read_failed", False)]
        meta_fail = [m for m in meta_results
                     if not m.unit_ok and not m.unit_skipped
                     and not getattr(m, "unit_read_failed", False)]
        meta_pass = (len(meta_results) - len(meta_fail)
                     - len(meta_skipped) - len(meta_read_failed))
        meta_rows = []
        for i, m in enumerate(meta_results):
            disp_unit = html.escape(m.bacnet_unit)
            if getattr(m, "unit_read_failed", False):
                # units 读取失败（超时/异常），未拿到值，不判错
                bg_u = "#fff3cd"
                flag = "读取失败"
                disp_unit = "<em style='color:#856404'>读取失败</em>"
            elif m.unit_skipped:
                bg_u = "#e2e3e5"
                flag = "跳过"
            elif m.unit_ok:
                bg_u = "#d4edda"
                flag = "✓"
            elif m.bacnet_unit == "":
                # 读取成功但网关返回 no-units（空单位），而模板要求有单位：网关缺单位
                bg_u = "#f8d7da"
                flag = "网关无单位"
                disp_unit = "<em style='color:#721c24'>no-units</em>"
            else:
                # 读到了单位但与模板不一致：单位写错
                bg_u = "#f8d7da"
                flag = "✗"
            meta_rows.append(f"""
        <tr>
          <td class="num">{i + 1}</td>
          <td class="key">{html.escape(m.param_key)}</td>
          <td>{html.escape(m.tmpl_unit)}</td>
          <td style="background:{bg_u}">{disp_unit}</td>
          <td style="background:{bg_u};text-align:center">{flag}</td>
        </tr>""")
        meta_badge = (
            (f'<span class="badge err-badge">失败 {len(meta_fail)}</span>'
             if meta_fail else "")
            + (f'<span class="badge warn-badge">读取失败 {len(meta_read_failed)}</span>'
               if meta_read_failed else "")
            + ('<span class="badge ok-badge">全部通过</span>'
               if not meta_fail and not meta_read_failed else "")
        )
        rf_card = (
            f'<div class="card"><div class="num" style="color:#856404">{len(meta_read_failed)}</div>'
            f'<div class="lbl">读取失败(未拿到单位)</div></div>'
        ) if meta_read_failed else ""
        skip_card = (
            f'<div class="card"><div class="num" style="color:#6c757d">{len(meta_skipped)}</div>'
            f'<div class="lbl">跳过(模板无单位)</div></div>'
        ) if meta_skipped else ""
        meta_html = f"""
<details open class="section">
<summary>二、单位检查（BACnet units 属性 vs 模板）
  <span class="sum-info">共 {len(meta_results)} 项</span>
  {meta_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(meta_results)}</div><div class="lbl">检查总数</div></div>
  <div class="card pass"> <div class="num">{meta_pass}</div><div class="lbl">通过</div></div>
  <div class="card fail"> <div class="num">{len(meta_fail)}</div><div class="lbl">失败</div></div>
  {rf_card}{skip_card}
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
        meta_html = """
<details class="section">
<summary>二、单位检查（BACnet units 属性 vs 模板）
  <span class="sum-info">无对照项</span>
  <span class="badge warn-badge">跳过</span>
</summary>
<div class="section-body">
<p style="color:#888;font-size:12px">未构建任何有模板对照的单位检查项</p>
</div>
</details>"""

    # ── Section 3: 数值比对 ───────────────────────────────────────────────────
    s = _cmp_summary(cmp_rows)
    if cmp_note and not cmp_rows:
        val_html = f"""
<details class="section">
<summary>三、数值比对（BACnet Present Value vs Modbus）
  <span class="sum-info">已跳过</span>
  <span class="badge warn-badge">跳过</span>
</summary>
<div class="section-body">
<p style="color:#888;font-size:12px">{html.escape(cmp_note)}</p>
</div>
</details>"""
    else:
        val_rows = []
        for i, r in enumerate(cmp_rows):
            key, bv, mv, da, dp, status = r
            bg = _STATUS_COLOR.get(status, "#ffffff")
            stat = _STATUS_LABEL.get(status, status)
            val_rows.append(f"""
        <tr style="background:{bg}">
          <td class="num">{i + 1}</td>
          <td class="key">{html.escape(key)}</td>
          <td class="val">{_fmt(bv)}</td>
          <td class="val">{_fmt(mv)}</td>
          <td class="val">{_fmt(da, 4) if da is not None else "—"}</td>
          <td class="val">{f"{dp:.3f}%" if dp is not None else "—"}</td>
          <td class="stat">{stat}</td>
        </tr>""")
        val_badge = (f'<span class="badge err-badge">失败 {s["fail"]}</span>'
                     if s["fail"] else "") + \
                    (f'<span class="badge warn-badge">异常 {s["error"]}</span>'
                     if s["error"] else "") + \
                    ('<span class="badge ok-badge">全部通过</span>'
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
  <col class="c-diff"><col class="c-pct"><col class="c-stat">
</colgroup>
<thead>
  <tr>
    <th>#</th><th>参数名 (param_key)</th>
    <th>BACnet 值</th><th>Modbus 值</th>
    <th>绝对差值</th><th>相对差值</th><th>结果</th>
  </tr>
</thead>
<tbody>{"".join(val_rows)}</tbody>
</table>
</div>
</details>"""

    # ── Section 4: Device Object 属性 ────────────────────────────────────────
    if dev_info is not None:
        _dev_props = [
            ("设备实例号 (objectIdentifier)", dev_info.object_identifier),
            ("设备名称 (objectName)", dev_info.object_name),
            ("系统状态 (systemStatus)", dev_info.system_status),
            ("厂商名称 (vendorName)", dev_info.vendor_name),
            ("厂商 ID (vendorIdentifier)", dev_info.vendor_id),
            ("型号名称 (modelName)", dev_info.model_name),
            ("固件版本 (firmwareRevision)", dev_info.firmware_revision),
            ("应用软件版本 (applicationSoftwareVersion)", dev_info.app_sw_version),
            ("协议版本 (protocolVersion)", dev_info.protocol_version),
            ("协议修订号 (protocolRevision)", dev_info.protocol_revision),
            ("最大 APDU 长度 (maxApduLengthAccepted)", dev_info.max_apdu_length),
            ("分段支持 (segmentationSupported)", dev_info.segmentation),
        ]
        # Python 3.11 的 f-string 表达式内不允许反斜杠，缺失占位串提前抽成常量
        _missing = "<em style='color:#721c24'>未获取</em>"
        dev_rows = "".join(
            f'<tr style="background:{"#d4edda" if v else "#f8d7da"}">'
            f'<td class="key">{html.escape(name)}</td>'
            f'<td class="val">{html.escape(str(v)) if v else _missing}</td></tr>'
            for name, v in _dev_props
        )
        dev_badge = (f'<span class="badge {"ok-badge" if dev_info.ok else "err-badge"}">'
                     f'{"OK" if dev_info.ok else "FAIL"}</span>')
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
    if compliance_results:
        passed_n = sum(1 for t in compliance_results if t.passed)
        proto_badge = (f'<span class="badge err-badge">失败 {len(compliance_results) - passed_n}</span>'
                       if passed_n < len(compliance_results) else
                       '<span class="badge ok-badge">全部通过</span>')
        err_rows = "".join(
            f'<tr style="background:{"#d4edda" if t.passed else "#f8d7da"}">'
            f'<td class="num">{i + 1}</td>'
            f'<td class="key">{html.escape(t.test_name)}</td>'
            f'<td class="stat">{"通过" if t.passed else "失败"}</td>'
            f'<td class="err">{html.escape(t.detail)}</td></tr>'
            for i, t in enumerate(compliance_results)
        )
        proto_html = f"""
<details open class="section">
<summary>五、协议合规性测试（ANSI/ASHRAE 135 §16 错误处理 + AI 必需属性）
  <span class="sum-info">{passed_n}/{len(compliance_results)} 通过</span>
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
        proto_html = ""

    # ── Section 6: 连接稳定性 ────────────────────────────────────────────────
    if stability is not None:
        stab_ok = stability.ok
        stab_badge = ('<span class="badge ok-badge">稳定</span>' if stab_ok else
                      '<span class="badge err-badge">不稳定</span>')
        err_list_html = "".join(
            f'<li style="color:#721c24">{html.escape(e)}</li>' for e in stability.errors
        ) if stability.errors else ""
        stab_err_attr = "" if stab_ok else ' data-has-error="1"'
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
<title>BACnet vs Modbus 比对报告 — {html.escape(device_name)}</title>
<style>
  body {{ font-family: "Microsoft YaHei", Arial, sans-serif; font-size: 13px;
          margin: 20px; background: #f5f5f5; color: #333; }}
  h1   {{ font-size: 20px; margin-bottom: 4px; }}
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
<div class="device-name">设备：{html.escape(device_name)}（模板 {html.escape(template_name)}）</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  网关：{html.escape(str(gateway_ip))}:{html.escape(str(gateway_port))} &nbsp;|&nbsp;
  容差：±{tol_pct}% / ±{tol_abs}
</div>
{scope_html}
{meta_html}
{val_html}
{dev_info_html}
{proto_html}
{stability_html}
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("[bacnet_report] HTML 报告已保存：%s", output_path)
    return output_path
