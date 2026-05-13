# -*- coding: utf-8 -*-
"""
mqtt_comparator.py — MQTT 快照 vs 实时 Modbus 三段式比对

比对流程：
  1. 范围检查：模板全量参数 vs JSON reading 中实际发布的参数
  2. 单位检查：JSON unit 字段 vs 模板 unit 列
  3. 数值比对：JSON value vs 实时 Modbus 寄存器值（并发读取）

MQTT JSON 格式：
  {
    "timestamp": <unix>,
    "comm_head": { "model": "ACM-41-WEB2", "sn": "..." },
    "modules": [
      {
        "name": "...", "model": "AcuRev-4110-mA", "sn": "...", "online": true,
        "reading": [
          { "param": "FREQ_Hz", "value": "50.000", "unit": "Hz" },
          ...
        ]
      }
    ]
  }

param 字段直接对应 param_key，无需列标题映射。

用法：
  python MQTT/mqtt_comparator.py                          # 自动选最新 JSON，第一个在线模块
  python MQTT/mqtt_comparator.py --device acurev4100      # 指定设备
  python MQTT/mqtt_comparator.py --file <json路径>        # 指定文件
  python MQTT/mqtt_comparator.py --module 0               # 指定模块下标（默认第一个在线模块）
  python MQTT/mqtt_comparator.py --no-meta                # 跳过单位检查
  python MQTT/mqtt_comparator.py --keys FREQ_Hz VLN_a_V   # 只比对指定参数
"""

from __future__ import annotations

import asyncio
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
from modbus_reader import ModbusReader, ModbusResult
from template_reader import TemplateParam, find_template_file, load_template

log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class MQTTScopeReport:
    """模板全量参数 vs JSON 实际发布参数的范围对比。"""
    template_count:  int
    json_count:      int
    matched_keys:    list[str]   # 两侧均有
    missing_from_json: list[str] # 模板有但 JSON 未发布
    extra_in_json:   list[str]   # JSON 发布但模板中无

    @property
    def scope_ok(self) -> bool:
        return not self.missing_from_json and not self.extra_in_json


@dataclass
class MQTTUnitResult:
    """单个参数的单位检查结果：模板 unit vs JSON unit。"""
    param_key:  str
    tmpl_unit:  str
    json_unit:  str
    unit_ok:    bool

    @property
    def ok(self) -> bool:
        return self.unit_ok


@dataclass
class MQTTCompareResult:
    """单个参数的 MQTT 快照 vs 实时 Modbus 数值比对结果。"""
    param_key:    str
    mqtt_value:   Optional[float] = None
    modbus_value: Optional[float] = None
    mqtt_error:   str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | MQTT_ERR | MODBUS_ERR | BOTH_ERR

    @property
    def ok(self) -> bool:
        return self.status == "PASS"


# ─────────────────────────────────────────────────────────────────────────────
# JSON 加载
# ─────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', s.lower())


def find_json_file(data_dir: str, device_hint: str = "") -> str:
    """在 data_dir 下查找匹配设备名的 JSON 文件，找不到时返回最新文件。"""
    json_files = sorted(
        Path(data_dir).glob("*.json"),
        key=lambda p: p.stat().st_mtime,
    )
    if not json_files:
        raise FileNotFoundError(f"在 {data_dir} 中未找到 JSON 文件")
    if device_hint:
        needle = _norm(device_hint)
        matching = [p for p in json_files if needle in _norm(p.stem)]
        if matching:
            return str(matching[-1])
        log.warning("未找到与 '%s' 匹配的 JSON 文件，回退至最新文件", device_hint)
    return str(json_files[-1])


def load_mqtt_json(
    json_path: str,
    module_index: Optional[int] = None,
) -> tuple[dict[str, float], dict[str, str], str, str, str]:
    """
    解析 MQTT JSON 文件。

    Returns:
        (value_map, unit_map, timestamp_str, module_name, module_model)
        value_map: {param_key: float}
        unit_map:  {param_key: unit_str}（原始字符串，用于单位检查）
    """
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    ts = data.get("timestamp", 0)
    timestamp_str = datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S") if ts else "未知"

    modules = data.get("modules", [])
    if not modules:
        raise ValueError(f"JSON 中无 modules 字段：{json_path}")

    if module_index is not None:
        if module_index >= len(modules):
            raise IndexError(
                f"module_index={module_index} 超出范围（共 {len(modules)} 个模块）"
            )
        mod = modules[module_index]
    else:
        online = [m for m in modules if m.get("online", False)]
        mod = online[0] if online else modules[0]
        if not online:
            log.warning("所有模块均 offline，使用第一个模块")

    module_name  = mod.get("name", "—")
    module_model = mod.get("model", "—")
    log.info("使用模块：name=%s  model=%s  online=%s",
             module_name, module_model, mod.get("online"))

    value_map: dict[str, float] = {}
    unit_map:  dict[str, str]   = {}
    for item in mod.get("reading", []):
        param = item.get("param", "").strip()
        if not param:
            continue
        unit_map[param] = str(item.get("unit", "")).strip()
        try:
            value_map[param] = float(item["value"])
        except (KeyError, ValueError, TypeError):
            log.debug("参数 %s 的 value 无法转浮点，已跳过", param)

    log.info("读取到 %d 个参数（模块 %s）", len(value_map), module_name)
    return value_map, unit_map, timestamp_str, module_name, module_model


# ─────────────────────────────────────────────────────────────────────────────
# 单位规范化（容许常见等价写法）
# ─────────────────────────────────────────────────────────────────────────────

_UNIT_ALIAS: dict[str, str] = {
    "hz":    "hz",
    "v":     "v",
    "a":     "a",
    "kw":    "kw",
    "kvar":  "kvar",
    "kva":   "kva",
    "kwh":   "kwh",
    "kvarh": "kvarh",
    "kvah":  "kvah",
    "%":     "%",
    "°":     "deg",
    "deg":   "deg",
}


def _norm_unit(u: str) -> str:
    """统一单位字符串以减少大小写/符号差异导致的误判。"""
    s = u.strip().lower()
    return _UNIT_ALIAS.get(s, s)


def check_unit(param_key: str, tmpl_unit: str, json_unit: str) -> MQTTUnitResult:
    return MQTTUnitResult(
        param_key = param_key,
        tmpl_unit = tmpl_unit,
        json_unit = json_unit,
        unit_ok   = _norm_unit(tmpl_unit) == _norm_unit(json_unit),
    )


# ─────────────────────────────────────────────────────────────────────────────
# 数值比对
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one(
    param_key: str,
    mqtt_value: Optional[float],
    mr: Optional[ModbusResult],
) -> MQTTCompareResult:
    cr = MQTTCompareResult(param_key=param_key)

    if mqtt_value is None:
        cr.mqtt_error = "JSON 无数据"
    if mr is None or not mr.ok:
        cr.modbus_error = (mr.error if mr else "未读取到")

    if cr.mqtt_error and cr.modbus_error:
        cr.status = "BOTH_ERR"
        return cr
    if cr.mqtt_error:
        cr.status = "MQTT_ERR"
        return cr
    if cr.modbus_error:
        cr.status = "MODBUS_ERR"
        return cr

    mv_mqtt   = mqtt_value
    mv_modbus = mr.value
    cr.mqtt_value   = mv_mqtt
    cr.modbus_value = mv_modbus

    diff = abs(mv_mqtt - mv_modbus)
    ref  = max(abs(mv_mqtt), abs(mv_modbus))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0

    tol = max(
        config.MQTT_TOLERANCE_ABSOLUTE,
        ref * config.MQTT_TOLERANCE_PERCENT / 100,
    )
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


# ─────────────────────────────────────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────────────────────────────────────

async def run_mqtt_comparison(
    json_path: str,
    module_index: Optional[int] = None,
    param_keys: Optional[list[str]] = None,
    check_meta: bool = True,
) -> tuple[MQTTScopeReport, list[MQTTUnitResult], list[MQTTCompareResult], str, str, str]:
    """
    三段式比对：范围检查 → 单位检查 → 数值比对。

    Returns:
        (scope_report, unit_results, compare_results, timestamp_str, module_name, module_model)
    """
    t0 = time.time()

    # ── 加载 JSON ──────────────────────────────────────────────────────────────
    log.info("加载 JSON：%s", json_path)
    value_map, unit_map, timestamp_str, module_name, module_model = load_mqtt_json(
        json_path, module_index
    )

    # ── 加载模板 ───────────────────────────────────────────────────────────────
    try:
        tmpl_path   = find_template_file(config.TEMPLATE_DIR, config.DEVICE_NAME)
        tmpl_params = load_template(tmpl_path)           # 全量，不按协议列过滤
        tmpl_map    = {p.param_key: p for p in tmpl_params}
        tmpl_keys   = set(tmpl_map)
        log.info("已加载模板：%s（%d 个参数）", tmpl_path, len(tmpl_keys))
    except Exception as exc:
        log.warning("无法加载模板文件，范围/单位检查将跳过：%s", exc)
        tmpl_map, tmpl_keys = {}, set()

    json_keys = set(value_map.keys())

    # ── 范围检查 ───────────────────────────────────────────────────────────────
    if tmpl_keys:
        matched_keys_set   = tmpl_keys & json_keys
        missing_from_json  = sorted(tmpl_keys - json_keys)
        extra_in_json      = sorted(json_keys - tmpl_keys)
    else:
        matched_keys_set  = json_keys
        missing_from_json = []
        extra_in_json     = []

    scope_report = MQTTScopeReport(
        template_count   = len(tmpl_keys),
        json_count       = len(json_keys),
        matched_keys     = sorted(matched_keys_set),
        missing_from_json = missing_from_json,
        extra_in_json    = extra_in_json,
    )
    log.info("范围检查：模板=%d  JSON=%d  匹配=%d  缺失=%d  多余=%d",
             len(tmpl_keys), len(json_keys), len(matched_keys_set),
             len(missing_from_json), len(extra_in_json))

    # ── 单位检查 ───────────────────────────────────────────────────────────────
    unit_results: list[MQTTUnitResult] = []
    if check_meta and tmpl_map:
        for pkey in sorted(matched_keys_set):
            tmpl_unit = tmpl_map[pkey].unit
            json_unit = unit_map.get(pkey, "")
            unit_results.append(check_unit(pkey, tmpl_unit, json_unit))
        unit_fail = sum(1 for r in unit_results if not r.ok)
        log.info("单位检查：共 %d 项，不匹配 %d 项", len(unit_results), unit_fail)

    # ── 确定数值比对参数范围 ───────────────────────────────────────────────────
    # 比对范围：模板与 JSON 的交集（或全部 JSON 参数，若无模板）
    compare_scope = matched_keys_set if tmpl_keys else json_keys
    comparable: dict[str, float] = {
        k: v for k, v in value_map.items()
        if k in compare_scope and (param_keys is None or k in param_keys)
    }
    if not comparable:
        raise ValueError("无可比对参数（JSON 与模板无交集，或指定的 --keys 均不存在）")

    log.info("数值比对参数数量：%d", len(comparable))

    # ── 实时读取 Modbus ────────────────────────────────────────────────────────
    log.info("读取实时 Modbus 寄存器…")
    async with ModbusReader() as modbus:
        modbus_results = await modbus.read_params(list(comparable.keys()))

    modbus_map: dict[str, ModbusResult] = {r.param_key: r for r in modbus_results}

    # ── 数值比对 ───────────────────────────────────────────────────────────────
    compare_results: list[MQTTCompareResult] = []
    for pkey, mv in comparable.items():
        mr = modbus_map.get(pkey)
        compare_results.append(_compare_one(pkey, mv, mr))

    elapsed = time.time() - t0
    log.info("比对完成，耗时 %.1f 秒，共 %d 项", elapsed, len(compare_results))
    return scope_report, unit_results, compare_results, timestamp_str, module_name, module_model


# ─────────────────────────────────────────────────────────────────────────────
# 统计摘要
# ─────────────────────────────────────────────────────────────────────────────

def summary(results: list[MQTTCompareResult]) -> dict:
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


def print_summary(
    scope: MQTTScopeReport,
    unit_results: list[MQTTUnitResult],
    results: list[MQTTCompareResult],
    module_name: str = "",
    module_model: str = "",
) -> None:
    s = summary(results)
    unit_fail = sum(1 for r in unit_results if not r.ok)
    print("\n" + "=" * 70)
    print("  MQTT 快照 vs 实时 Modbus 比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    if module_name or module_model:
        print(f"  模块: {module_name}  ({module_model})")
    print("=" * 70)
    print(f"  【范围检查】模板={scope.template_count}  JSON={scope.json_count}  "
          f"匹配={len(scope.matched_keys)}  "
          f"缺失={len(scope.missing_from_json)}  多余={len(scope.extra_in_json)}")
    if scope.missing_from_json:
        print(f"  缺失（前10）: {scope.missing_from_json[:10]}")
    if scope.extra_in_json:
        print(f"  多余（前10）: {scope.extra_in_json[:10]}")
    if unit_results:
        print(f"  【单位检查】共 {len(unit_results)} 项，不匹配 {unit_fail} 项")
        if unit_fail:
            for r in unit_results:
                if not r.ok:
                    print(f"    {r.param_key:40s}  模板={r.tmpl_unit!r}  JSON={r.json_unit!r}")
    print(f"  【数值比对】总={s['total']}  PASS={s['pass']} ({s['pass_rate']})  "
          f"FAIL={s['fail']}  ERR={s['error']}")
    if s["worst_fails"]:
        print("\n  差异最大的失败参数（Top 10）：")
        for r in s["worst_fails"]:
            print(f"    {r.param_key:40s}  MQTT={r.mqtt_value:.6g}  "
                  f"Modbus={r.modbus_value:.6g}  Δ%={r.diff_pct:.2f}%")
    print("=" * 70)


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

_STATUS_COLOR = {
    "PASS":       "#d4edda",
    "FAIL":       "#f8d7da",
    "MQTT_ERR":   "#fff3cd",
    "MODBUS_ERR": "#fff3cd",
    "BOTH_ERR":   "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS":       "通过",
    "FAIL":       "失败",
    "MQTT_ERR":   "MQTT异常",
    "MODBUS_ERR": "Modbus异常",
    "BOTH_ERR":   "双路异常",
}


def _fmt(v: Optional[float], digits: int = 6) -> str:
    return "—" if v is None else f"{v:.{digits}g}"


def generate_html_report(
    scope: MQTTScopeReport,
    unit_results: list[MQTTUnitResult],
    results: list[MQTTCompareResult],
    timestamp_str: str = "",
    module_name: str = "",
    module_model: str = "",
    json_path: str = "",
    output_path: Optional[str] = None,
) -> str:
    """生成三段式 HTML 比对报告（范围检查 / 单位检查 / 数值比对），返回文件路径。"""
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"mqtt_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str       = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    json_basename = Path(json_path).name if json_path else "—"
    unit_fail     = sum(1 for r in unit_results if not r.ok)

    # ── Section 1: 范围检查 ───────────────────────────────────────────────────
    scope_color = "#d4edda" if scope.scope_ok else "#f8d7da"
    scope_label = "一致" if scope.scope_ok else "不一致"

    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    missing_html = (
        f'<h3 style="color:#721c24;margin-top:12px">模板有但 JSON 未发布（{len(scope.missing_from_json)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.missing_from_json, "#f8d7da")}</tbody></table>'
    ) if scope.missing_from_json else ""

    extra_html = (
        f'<h3 style="color:#856404;margin-top:12px">JSON 发布但模板中无（{len(scope.extra_in_json)} 条）</h3>'
        f'<table><colgroup><col style="width:48px"><col></colgroup>'
        f'<thead><tr><th>#</th><th>param_key</th></tr></thead>'
        f'<tbody>{_list_rows(scope.extra_in_json, "#fff3cd")}</tbody></table>'
    ) if scope.extra_in_json else ""

    scope_badge = (f'<span class="badge err-badge">缺失 {len(scope.missing_from_json)}</span>'
                   if scope.missing_from_json else "") + \
                  (f'<span class="badge warn-badge">多余 {len(scope.extra_in_json)}</span>'
                   if scope.extra_in_json else "") + \
                  ('<span class="badge ok-badge">一致</span>' if scope.scope_ok else "")

    scope_html = f"""
<details open class="section">
<summary>一、参数范围检查
  <span class="sum-info">模板 {scope.template_count} / JSON {scope.json_count} / 匹配 {len(scope.matched_keys)}</span>
  {scope_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板参数</div></div>
  <div class="card total"><div class="num">{scope.json_count}</div><div class="lbl">JSON 发布</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_json)}</div><div class="lbl">模板有/JSON缺</div></div>
  <div class="card err">  <div class="num">{len(scope.extra_in_json)}</div><div class="lbl">JSON多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:18px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}
</div>
</details>"""

    # ── Section 2: 单位检查 ───────────────────────────────────────────────────
    if unit_results:
        unit_badge = (f'<span class="badge err-badge">不匹配 {unit_fail}</span>'
                      if unit_fail else '<span class="badge ok-badge">全部一致</span>')
        unit_rows = "".join(
            f'<tr style="background:{"#f8d7da" if not r.ok else "#d4edda"}">'
            f'<td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(r.param_key)}</td>'
            f'<td class="val">{html.escape(r.tmpl_unit)}</td>'
            f'<td class="val">{html.escape(r.json_unit)}</td>'
            f'<td class="stat">{"不匹配" if not r.ok else "一致"}</td>'
            f'</tr>'
            for i, r in enumerate(unit_results)
        )
        unit_html = f"""
<details {"open" if unit_fail else ""} class="section">
<summary>二、单位检查（模板 unit vs JSON unit）
  <span class="sum-info">共 {len(unit_results)} 项</span>
  {unit_badge}
</summary>
<div class="section-body">
<div class="cards">
  <div class="card total"><div class="num">{len(unit_results)}</div><div class="lbl">检查参数</div></div>
  <div class="card pass"> <div class="num">{len(unit_results)-unit_fail}</div><div class="lbl">单位一致</div></div>
  <div class="card fail"> <div class="num">{unit_fail}</div><div class="lbl">单位不匹配</div></div>
</div>
<table>
<colgroup><col class="c-idx"><col class="c-key"><col class="c-val"><col class="c-val"><col class="c-stat"></colgroup>
<thead><tr><th>#</th><th>param_key</th><th>模板 unit</th><th>JSON unit</th><th>结果</th></tr></thead>
<tbody>{unit_rows}</tbody>
</table>
</div>
</details>"""
        val_section_num = "三"
    else:
        unit_html = """
<details class="section">
<summary>二、单位检查<span class="sum-info">（已跳过）</span></summary>
<div class="section-body"><p style="color:#888">使用 --no-meta 跳过或模板文件不可用。</p></div>
</details>"""
        val_section_num = "三"

    # ── Section 3: 数值比对 ───────────────────────────────────────────────────
    rows_html = []
    for i, r in enumerate(results):
        bg   = _STATUS_COLOR.get(r.status, "#ffffff")
        stat = _STATUS_LABEL.get(r.status, r.status)
        cv_str = _fmt(r.mqtt_value)
        mv_str = _fmt(r.modbus_value)
        da_str = _fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"
        dp_str = f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"
        err_hint = ""
        if r.mqtt_error:
            err_hint += f"MQTT: {html.escape(r.mqtt_error)}"
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
                ('<span class="badge ok-badge">全部通过</span>'
                 if not s["fail"] and not s["error"] else "")

    val_html = f"""
<details open class="section">
<summary>{val_section_num}、数值比对（MQTT 快照 vs 实时 Modbus）
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
  <col class="c-idx"><col class="c-key">
  <col class="c-val"><col class="c-val">
  <col class="c-diff"><col class="c-pct">
  <col class="c-stat"><col class="c-err">
</colgroup>
<thead>
  <tr>
    <th>#</th><th>参数名 (param_key)</th>
    <th>MQTT 快照值</th><th>Modbus 实时值</th>
    <th>绝对差值</th><th>相对差值</th>
    <th>结果</th><th>错误信息</th>
  </tr>
</thead>
<tbody>{rows_str}</tbody>
</table>
</div>
</details>"""

    css = """
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
  .section-body {{ padding: 16px; }}"""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>MQTT vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>{css}
</style>
</head>
<body>
<h1>MQTT 快照 vs 实时 Modbus 比对报告</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  数据文件：{html.escape(json_basename)} &nbsp;|&nbsp;
  快照时间戳：{html.escape(timestamp_str)} &nbsp;|&nbsp;
  模块：{html.escape(module_name)} ({html.escape(module_model)}) &nbsp;|&nbsp;
  Modbus：{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT} &nbsp;|&nbsp;
  容差：±{config.MQTT_TOLERANCE_PERCENT}% / ±{config.MQTT_TOLERANCE_ABSOLUTE}
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
# 设备映射表
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
    json_path: str,
    module_index: Optional[int],
    param_keys: Optional[list[str]],
    check_meta: bool,
) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )
    scope, unit_results, results, timestamp_str, module_name, module_model = \
        await run_mqtt_comparison(json_path, module_index, param_keys, check_meta)
    print_summary(scope, unit_results, results, module_name, module_model)
    report_path = generate_html_report(
        scope, unit_results, results,
        timestamp_str=timestamp_str,
        module_name=module_name,
        module_model=module_model,
        json_path=json_path,
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
        _json_path = sys.argv[idx + 1]
    else:
        _device_hint = _dev_value or ""
        _json_path = find_json_file(config.MQTT_DATA_DIR, device_hint=_device_hint)
        print(f"[INFO] 自动选取文件：{Path(_json_path).name}")

    # ── --module ──────────────────────────────────────────────────────────────
    _module_index: Optional[int] = None
    if "--module" in sys.argv:
        idx = sys.argv.index("--module")
        _module_index = int(sys.argv[idx + 1])

    # ── --no-meta ─────────────────────────────────────────────────────────────
    _check_meta = "--no-meta" not in sys.argv

    # ── --keys ────────────────────────────────────────────────────────────────
    _keys: Optional[list[str]] = None
    if "--keys" in sys.argv:
        idx = sys.argv.index("--keys")
        _keys = sys.argv[idx + 1:]

    asyncio.run(_main(_json_path, _module_index, _keys, _check_meta))
