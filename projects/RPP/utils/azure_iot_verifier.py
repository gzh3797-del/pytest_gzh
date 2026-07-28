"""
Azure IoT 消息订阅验证脚本（AcuHMI-1-7 工程版）

四阶段验证：
  [1] 设备列表：上报设备是否与期望设备一致
  [2] 参数完整性：每台设备上报的参数是否覆盖模板 MQTT 列全量参数
  [3] 单位一致性：每个参数的上报单位是否与模板 unit 列一致
  [4] 数值比对：Azure IoT 上报值 vs 设备实时 Modbus 读取值

用法：
    python utils/azure_iot_verifier.py --config tests/azure_iot/config.yaml
"""
import asyncio
import html
import json
import logging
import os
import re
import sys
import time
import queue as _queue
import threading
from dataclasses import dataclass
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from azure.eventhub import EventHubConsumerClient
import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent   # AcuHMI-1-7/
REPORT_DIR   = str(PROJECT_ROOT / "reports")

sys.path.insert(0, str(PROJECT_ROOT))

from utils.modbus_reader import ModbusResult, read_device_params

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


# ─── Device name mappings ──────────────────────────────────────────────────────

# Gateway-reported name → canonical excel_param_map device name
_DEV_ALIAS_MAP: dict[str, str] = {
    "4100242":   "AcuRev4100",
    "Acu4100":   "AcuRev4100",
}

# Canonical name → excel_param_map key (matches _EXCEL_FILES in excel_param_map.py)
_DEV_EXCEL_MAP: dict[str, str] = {
    "AcuRev4100": "AcuRev4100",
    "AcuRev2100": "AcuRev2100",
    "AcuvimIIW":  "AcuvimIIW",
    "AcuvimIIR":  "AcuvimIIR",
    "AcuVIM3":    "Acuvim3",
    "Acuvim3":    "Acuvim3",
    "AcuRev1300": "AcuRev1300",
}

_UNIT_ALIAS: dict[str, str] = {
    "hz": "hz", "v": "v", "a": "a",
    "kw": "kw", "kvar": "kvar", "kva": "kva",
    "kwh": "kwh", "kvarh": "kvarh", "kvah": "kvah",
    "%": "%", "deg": "deg", "°": "deg",
}


def _norm_unit(u: str) -> str:
    s = u.strip().lower()
    return _UNIT_ALIAS.get(s, s)


# ─── Data structures ───────────────────────────────────────────────────────────

@dataclass
class ValueCompareResult:
    param_key:    str
    azure_value:  Optional[float] = None
    modbus_value: Optional[float] = None
    azure_error:  str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | AZURE_ERR | MODBUS_ERR | BOTH_ERR


# ─── Config helpers ────────────────────────────────────────────────────────────

def load_config(path: str) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def _abs_path(path: str) -> str:
    if os.path.isabs(path):
        return path
    return str(PROJECT_ROOT / path)


# ─── Template loading ──────────────────────────────────────────────────────────

def _load_mqtt_template(device_name: str) -> dict[str, str] | None:
    from utils.excel_param_map import load_device_params

    canonical = _DEV_ALIAS_MAP.get(device_name, device_name)
    excel_name = _DEV_EXCEL_MAP.get(canonical, canonical)

    # Try several candidate names
    candidates = [excel_name, canonical, device_name]
    for num in re.findall(r'\d{3,4}', device_name):
        candidates.append(num)

    seen = set()
    for cand in candidates:
        if cand in seen:
            continue
        seen.add(cand)
        try:
            params = load_device_params(cand, snmp_only=False, mqtt_only=True)
            if params:
                logging.info(f"  [模板] {device_name} → {cand}: MQTT 参数 {len(params)} 个")
                return {ptype: info["unit"] for ptype, info in params.items()}
        except Exception:
            continue
    return None


# ─── Payload parsing ───────────────────────────────────────────────────────────

def _parse_modules(payload: dict) -> list[dict]:
    device = payload.get('device')
    if isinstance(device, dict):
        name = (device.get('name') or device.get('deviceName') or '').strip()
        if name:
            readings = device.get('readings') or device.get('reading') or []
            return [{'name': name, 'reading': readings,
                     'online': payload.get('online', True),
                     'model': device.get('model', '')}]

    modules = payload.get('modules')
    if isinstance(modules, list):
        result = []
        for m in modules:
            name = (m.get('name') or m.get('deviceName') or '').strip()
            if not name:
                continue
            readings = m.get('readings') or m.get('reading') or []
            result.append({'name': name, 'reading': readings,
                           'online': m.get('online', True), 'model': m.get('model', '')})
        return result

    for key in ('deviceName', 'device_name', 'name'):
        name = payload.get(key)
        if isinstance(name, str) and name.strip():
            readings = payload.get('readings') or payload.get('reading') or []
            return [{'name': name.strip(), 'reading': readings, 'online': True, 'model': ''}]

    data = payload.get('data')
    if isinstance(data, dict):
        return _parse_modules(data)
    return []


# ─── Parameter check ───────────────────────────────────────────────────────────

def _check_params(json_params: dict[str, str], tmpl_map: dict[str, str]) -> dict:
    tmpl_keys = set(tmpl_map.keys())
    json_keys  = set(json_params.keys())
    missing = sorted(tmpl_keys - json_keys)
    extra   = sorted(json_keys - tmpl_keys)
    unit_mismatch = []
    for param in sorted(tmpl_keys & json_keys):
        if _norm_unit(tmpl_map[param]) != _norm_unit(json_params[param]):
            unit_mismatch.append((param, tmpl_map[param], json_params[param]))
    return {
        'tmpl_total': len(tmpl_keys), 'report_total': len(json_keys),
        'missing': missing, 'extra': extra, 'unit_mismatch': unit_mismatch,
        'ok': not missing and not unit_mismatch,
        '_tmpl_map': tmpl_map, '_reported_units': json_params,
        '_all_params': sorted(tmpl_keys | json_keys),
    }


# ─── Value comparison ──────────────────────────────────────────────────────────

def _compare_one_value(
    param_key: str, azure_value: Optional[float],
    mr: Optional[ModbusResult], tol_pct: float, tol_abs: float,
) -> ValueCompareResult:
    import math
    cr = ValueCompareResult(param_key=param_key)
    if azure_value is None:
        cr.azure_error = "无值"
    if mr is None or not mr.ok:
        cr.modbus_error = mr.error if mr else "未读取"
    if cr.azure_error and cr.modbus_error:
        cr.status = "BOTH_ERR"; return cr
    if cr.azure_error:
        cr.status = "AZURE_ERR"; return cr
    if cr.modbus_error:
        cr.status = "MODBUS_ERR"; return cr
    cr.azure_value = azure_value
    cr.modbus_value = mr.value
    if math.isnan(azure_value) and math.isnan(mr.value):
        cr.diff_abs = 0.0; cr.diff_pct = 0.0; cr.status = "PASS"; return cr
    diff = abs(azure_value - mr.value)
    ref  = max(abs(azure_value), abs(mr.value))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0
    tol = max(tol_abs, ref * tol_pct / 100)
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


async def _read_modbus_for_device(
    dev_name: str,
    param_keys: list[str],
    modbus_device_map: dict,
) -> tuple[dict[str, ModbusResult] | None, str]:
    canonical = _DEV_ALIAS_MAP.get(dev_name, dev_name)
    modbus_info = modbus_device_map.get(canonical)
    if modbus_info is None:
        for k, v in modbus_device_map.items():
            if k.lower() == canonical.lower():
                modbus_info = v
                break
    if not modbus_info:
        return None, f"未在 config.yaml modbus.devices 中配置 {canonical!r}"

    host = str(modbus_info[0])
    port = int(modbus_info[1])
    unit = int(modbus_info[2])
    excel_name = _DEV_EXCEL_MAP.get(canonical, canonical)

    try:
        result = await read_device_params(excel_name, param_keys, host, port, unit)
        return result, ""
    except Exception as e:
        return None, str(e)


async def _compare_all_devices(
    device_values: dict,
    tol_pct: float,
    tol_abs: float,
    modbus_device_map: dict,
    preread_modbus: 'dict | None' = None,
    preread_errors: 'dict | None' = None,
    paired_azure:   'dict | None' = None,
) -> dict:
    results: dict = {}
    for dev_name in sorted(device_values.keys()):
        val_map = (paired_azure or {}).get(dev_name) or device_values[dev_name]
        comparable_keys = [k for k, v in val_map.items() if v is not None]
        if not comparable_keys:
            results[dev_name] = {'skip': '无可比对数值（所有参数缺失 value 字段）'}
            continue
        if preread_modbus is not None and dev_name in preread_modbus:
            modbus_map = preread_modbus[dev_name]
            err = (preread_errors or {}).get(dev_name, '')
            if modbus_map is None:
                results[dev_name] = {'skip': err or '配对 Modbus 读取失败'}
                logging.info(f"  [Modbus] {dev_name} 跳过（配对失败）：{err}")
                continue
            logging.info(f"  [Modbus] {dev_name}：使用同步配对读取（{len(modbus_map)} 个参数）")
        else:
            logging.info(f"  [Modbus] {dev_name}：无配对结果，实时读取 {len(comparable_keys)} 个参数…")
            modbus_map, err = await _read_modbus_for_device(
                dev_name, comparable_keys, modbus_device_map
            )
            if modbus_map is None:
                results[dev_name] = {'skip': err}
                logging.info(f"  [Modbus] {dev_name} 跳过：{err}")
                continue
        cmp_list = [
            _compare_one_value(k, val_map[k], modbus_map.get(k), tol_pct, tol_abs)
            for k in sorted(comparable_keys)
        ]
        results[dev_name] = cmp_list
        p = sum(1 for r in cmp_list if r.status == "PASS")
        f = sum(1 for r in cmp_list if r.status == "FAIL")
        logging.info(f"  [Modbus] {dev_name}：PASS={p}  FAIL={f}  ERR={len(cmp_list)-p-f}")
    return results


# ─── Azure Event Hub message receiver ─────────────────────────────────────────

def _receive_messages(cfg: dict, timeout: int) -> list[dict]:
    """通过 Azure Event Hub 兼容端点接收设备消息。"""
    azure = cfg["azure_iot"]
    eventhub_conn_str = azure.get("eventhub_conn_str", "")
    if not eventhub_conn_str:
        logging.warning("[Azure] eventhub_conn_str 未配置，跳过消息接收")
        return []

    messages: list[dict] = []
    _stop = threading.Event()

    def on_event(partition_context, event):
        if _stop.is_set() or event is None:
            return
        try:
            messages.append(json.loads(event.body_as_str()))
        except (json.JSONDecodeError, UnicodeDecodeError, ValueError):
            pass

    starting_position = datetime.now(timezone.utc) - timedelta(seconds=30)
    client = EventHubConsumerClient.from_connection_string(
        conn_str=eventhub_conn_str,
        consumer_group="$Default",
    )

    def _run():
        try:
            client.receive(on_event=on_event, starting_position=starting_position)
        except Exception:
            pass

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    t.join(timeout=timeout)
    _stop.set()
    try:
        client.close()
    except Exception:
        pass
    return messages


# ─── HTML report ───────────────────────────────────────────────────────────────

_CSS = """
  body{font-family:"Microsoft YaHei",Arial,sans-serif;font-size:13px;margin:20px;background:#f5f5f5;color:#333}
  h1{font-size:20px;margin-bottom:4px}
  .meta{color:#666;font-size:12px;margin-bottom:16px}
  .cards{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}
  .card{background:#fff;border-radius:6px;padding:14px 20px;box-shadow:0 1px 4px rgba(0,0,0,.1);min-width:110px;text-align:center}
  .card .num{font-size:28px;font-weight:bold}
  .card .lbl{font-size:11px;color:#888;margin-top:2px}
  .card.pass .num{color:#28a745}.card.fail .num{color:#dc3545}
  .card.warn .num{color:#ffc107}.card.total .num{color:#007bff}
  table{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.1);
        border-radius:6px;overflow:hidden;table-layout:fixed;margin-bottom:16px}
  colgroup col.c-idx{width:48px}colgroup col.c-key{width:280px}
  colgroup col.c-val{width:110px}colgroup col.c-diff{width:90px}
  colgroup col.c-pct{width:80px}colgroup col.c-stat{width:90px}colgroup col.c-err{width:auto}
  thead tr{background:#343a40;color:#fff}
  th{padding:9px 8px;text-align:center;font-size:12px;white-space:nowrap}
  td{padding:5px 8px;border-bottom:1px solid #eee;vertical-align:middle}
  td.num{text-align:right;color:#999;font-size:11px}
  td.key{font-family:monospace;font-size:12px;word-break:break-all}
  td.val{text-align:right;font-family:monospace;font-size:12px}
  td.stat{text-align:center;font-weight:bold;font-size:12px}
  td.err{font-size:11px;color:#666;word-break:break-all}
  tr:hover td{filter:brightness(.96)}
  thead th{position:sticky;top:0;z-index:1}
  details.section{background:#fff;border-radius:8px;margin-bottom:20px;
                  box-shadow:0 1px 4px rgba(0,0,0,.1);overflow:hidden}
  details.section>summary{list-style:none;cursor:pointer;padding:12px 16px;
    background:#f0f4f8;border-left:4px solid #0056b3;font-size:15px;font-weight:bold;
    color:#0056b3;display:flex;align-items:center;gap:8px;user-select:none}
  details.section>summary::-webkit-details-marker{display:none}
  details.section>summary::before{content:"▶";font-size:10px;margin-right:4px}
  details[open].section>summary::before{content:"▼"}
  details.subsection{background:#f8f9fa;border-radius:6px;margin-bottom:12px;
                     border:1px solid #dee2e6;overflow:hidden}
  details.subsection>summary{list-style:none;cursor:pointer;padding:8px 12px;
    background:#e9ecef;font-size:13px;font-weight:bold;color:#495057;
    display:flex;align-items:center;gap:6px}
  details.subsection>summary::-webkit-details-marker{display:none}
  details.subsection>summary::before{content:"▶";font-size:9px;margin-right:4px}
  details[open].subsection>summary::before{content:"▼"}
  .sum-info{font-size:12px;font-weight:normal;color:#666;margin-left:4px}
  .badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:bold;margin-left:4px}
  .ok-badge{background:#d4edda;color:#155724}
  .err-badge{background:#f8d7da;color:#721c24}
  .warn-badge{background:#fff3cd;color:#856404}
  .section-body{padding:16px}
  .overall-banner{padding:16px 20px;border-radius:8px;margin-bottom:20px;
                  font-size:18px;font-weight:bold;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.1)}
"""

_STATUS_COLOR = {
    "PASS": "#d4edda", "FAIL": "#f8d7da",
    "AZURE_ERR": "#fff3cd", "MODBUS_ERR": "#fff3cd", "BOTH_ERR": "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS": "通过", "FAIL": "失败",
    "AZURE_ERR": "Azure异常", "MODBUS_ERR": "Modbus异常", "BOTH_ERR": "双路异常",
}


def _badge(text: str, style: str) -> str:
    cls = {'ok': 'ok-badge', 'err': 'err-badge', 'warn': 'warn-badge'}[style]
    return f'<span class="badge {cls}">{html.escape(text)}</span>'


def _fmt(v: Optional[float], digits: int = 6) -> str:
    if v is None:
        return "—"
    s = f"{v:.{digits}f}"
    return s.rstrip('0').rstrip('.') if '.' in s else s


def _gen_device_list_section(reported: set[str], expected: set[str]) -> str:
    missing_d = sorted(expected - reported)
    extra_d   = sorted(reported - expected)
    ok = not missing_d and not extra_d
    badge = _badge('一致', 'ok') if ok else (_badge('不一致', 'err') if missing_d else _badge('多余设备', 'warn'))
    rows = ""
    for d in sorted(reported):
        if d in expected:
            rows += (f'<tr style="background:#d4edda"><td class="num"></td>'
                     f'<td class="key">{html.escape(d)}</td>'
                     f'<td class="stat" style="color:#28a745">✓ 已上报</td></tr>')
        else:
            rows += (f'<tr style="background:#fff3cd"><td class="num"></td>'
                     f'<td class="key">{html.escape(d)}</td>'
                     f'<td class="stat" style="color:#856404">! 多余</td></tr>')
    for d in missing_d:
        rows += (f'<tr style="background:#f8d7da"><td class="num"></td>'
                 f'<td class="key">{html.escape(d)}</td>'
                 f'<td class="stat" style="color:#721c24">✗ 缺少</td></tr>')
    err_attr = '' if ok else ' data-has-error="1"'
    return f"""
<details open class="section"{err_attr}>
<summary>一、设备列表检查
  <span class="sum-info">期望 {len(expected)} 台 / 上报 {len(reported)} 台</span>
  {badge}
</summary>
<div class="section-body">
<table>
  <colgroup><col class="c-idx"><col class="c-key"><col style="width:200px"></colgroup>
  <thead><tr><th>#</th><th>设备名称</th><th>状态</th></tr></thead>
  <tbody>{rows}</tbody>
</table>
</div>
</details>"""


def _gen_param_unit_section(reported: set[str], device_results: dict) -> str:
    total_pass = total_fail = total_warn = 0
    subs = ""
    for dev_name in sorted(reported):
        r = device_results.get(dev_name)
        if r is None:
            total_warn += 1
            subs += (f'<details class="subsection"><summary>{html.escape(dev_name)}'
                     f'<span class="sum-info">无 MQTT 模板</span>{_badge("跳过","warn")}</summary>'
                     f'<div class="section-body"><p style="color:#856404">未找到模板，跳过。</p></div></details>')
            continue
        if r['ok']:
            total_pass += 1
            dev_b = _badge('通过', 'ok')
        else:
            total_fail += 1
            dev_b = _badge('失败', 'err')
        rows = ""
        miss_set  = set(r['missing'])
        extra_set = set(r['extra'])
        mis_set   = {p for p, _, _ in r['unit_mismatch']}
        for i, param in enumerate(r['_all_params']):
            ru = r['_reported_units'].get(param, '')
            tu = r['_tmpl_map'].get(param, '')
            if param in miss_set:
                bg, st = '#f8d7da', '<span style="color:#721c24;font-weight:bold">✗ 缺失</span>'
                rc, tc = '<td class="val">—</td>', f'<td class="val">{html.escape(tu)}</td>'
            elif param in mis_set:
                _, t_u, j_u = next(x for x in r['unit_mismatch'] if x[0] == param)
                bg, st = '#fff3cd', '<span style="color:#856404;font-weight:bold">! 单位不符</span>'
                rc = f'<td class="val" style="color:#856404">{html.escape(j_u)}</td>'
                tc = f'<td class="val">{html.escape(t_u)}</td>'
            elif param in extra_set:
                bg, st = '#e8f4fd', '<span style="color:#0056b3">+ 多余</span>'
                rc, tc = f'<td class="val">{html.escape(ru)}</td>', '<td class="val">—</td>'
            else:
                bg, st = '#d4edda', '<span style="color:#28a745">✓</span>'
                rc = f'<td class="val">{html.escape(ru)}</td>'
                tc = f'<td class="val">{html.escape(tu)}</td>'
            rows += (f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
                     f'<td class="key">{html.escape(param)}</td>{tc}{rc}'
                     f'<td class="stat">{st}</td></tr>')
        extras = " ".join(filter(None, [
            f"缺失 {len(r['missing'])}" if r['missing'] else "",
            f"单位不符 {len(r['unit_mismatch'])}" if r['unit_mismatch'] else "",
            f"多余 {len(r['extra'])}" if r['extra'] else "",
        ]))
        ea = ' data-has-error="1"' if not r['ok'] else ''
        subs += f"""<details {"open" if not r['ok'] else ""} class="subsection"{ea}>
<summary>{html.escape(dev_name)}
  <span class="sum-info">模板 {r['tmpl_total']} / 上报 {r['report_total']}{(" / "+extras) if extras else ""}</span>
  {dev_b}
</summary>
<div class="section-body">
<table>
  <colgroup><col class="c-idx"><col class="c-key">
    <col style="width:100px"><col style="width:100px"><col style="width:120px">
  </colgroup>
  <thead><tr><th>#</th><th>参数名</th><th>模板单位</th><th>上报单位</th><th>状态</th></tr></thead>
  <tbody>{rows}</tbody>
</table></div></details>"""
    ob = (_badge('全部通过', 'ok') if total_fail == 0 and total_warn == 0
          else _badge(f'{total_fail} 个失败', 'err') if total_fail else _badge(f'{total_warn} 个跳过', 'warn'))
    ea = ' data-has-error="1"' if total_fail else ''
    return f"""
<details open class="section"{ea}>
<summary>二、参数完整性 &amp; 单位检查
  <span class="sum-info">{total_pass} 通过 / {total_fail} 失败 / {total_warn} 跳过</span>
  {ob}
</summary>
<div class="section-body">{subs}</div>
</details>"""


def _gen_value_section(
    value_results: dict,
    tol_pct: float,
    tol_abs: float,
) -> str:
    total_pass = total_fail = total_err = total_skip = 0
    subs = ""
    for dev_name in sorted(value_results.keys()):
        res = value_results[dev_name]
        if isinstance(res, dict):
            total_skip += 1
            reason = res.get('skip', res.get('error', ''))
            subs += (f'<details class="subsection"><summary>{html.escape(dev_name)}'
                     f'<span class="sum-info">{html.escape(reason)}</span>'
                     f'{_badge("跳过","warn")}</summary>'
                     f'<div class="section-body"><p style="color:#856404">{html.escape(reason)}</p>'
                     f'</div></details>')
            continue
        cmp_list: list[ValueCompareResult] = res
        dp = sum(1 for r in cmp_list if r.status == "PASS")
        df = sum(1 for r in cmp_list if r.status == "FAIL")
        de = len(cmp_list) - dp - df
        total_pass += dp; total_fail += df; total_err += de
        dev_b = (_badge(f'失败 {df}', 'err') if df
                 else _badge(f'异常 {de}', 'warn') if de
                 else _badge('通过', 'ok'))
        rows = ""
        for i, r in enumerate(cmp_list):
            bg   = _STATUS_COLOR.get(r.status, "#fff")
            stat = _STATUS_LABEL.get(r.status, r.status)
            err_hint = ""
            if r.azure_error:  err_hint += f"Azure: {html.escape(r.azure_error)}"
            if r.modbus_error: err_hint += f"{'<br>' if err_hint else ''}Modbus: {html.escape(r.modbus_error)}"
            da = _fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"
            dp_s = f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"
            rows += (f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
                     f'<td class="key">{html.escape(r.param_key)}</td>'
                     f'<td class="val">{_fmt(r.azure_value)}</td>'
                     f'<td class="val">{_fmt(r.modbus_value)}</td>'
                     f'<td class="val">{da}</td><td class="val">{dp_s}</td>'
                     f'<td class="stat">{stat}</td><td class="err">{err_hint}</td></tr>')
        pr = f"{dp/len(cmp_list)*100:.1f}%" if cmp_list else "N/A"
        ea = ' data-has-error="1"' if df else ''
        subs += f"""<details {"open" if df else ""} class="subsection"{ea}>
<summary>{html.escape(dev_name)}
  <span class="sum-info">共 {len(cmp_list)} 个 / PASS {dp} ({pr}) / FAIL {df} / ERR {de}</span>
  {dev_b}
</summary>
<div class="section-body">
<table>
<colgroup><col class="c-idx"><col class="c-key"><col class="c-val"><col class="c-val">
  <col class="c-diff"><col class="c-pct"><col class="c-stat"><col class="c-err">
</colgroup>
<thead><tr><th>#</th><th>参数名</th><th>Azure IoT 值</th><th>Modbus 值</th>
  <th>绝对差</th><th>相对差</th><th>结果</th><th>错误信息</th></tr></thead>
<tbody>{rows}</tbody>
</table></div></details>"""

    overall_total = total_pass + total_fail + total_err
    pr_all = f"{total_pass/overall_total*100:.1f}%" if overall_total else "N/A"
    ob = (_badge(f'失败 {total_fail}', 'err') if total_fail
          else _badge(f'异常 {total_err}', 'warn') if total_err
          else _badge(f'跳过 {total_skip}', 'warn') if total_skip and not overall_total
          else _badge('全部通过', 'ok'))
    ea = ' data-has-error="1"' if total_fail else ''
    return f"""
<details open class="section"{ea}>
<summary>三、数值比对（Azure IoT vs 实时 Modbus）
  <span class="sum-info">容差 ±{tol_pct}% / ±{tol_abs} · {total_pass} 通过({pr_all}) / {total_fail} 失败 / {total_err} 异常 / {total_skip} 跳过</span>
  {ob}
</summary>
<div class="section-body">{subs}</div>
</details>"""


def generate_html_report(
    config_info: dict,
    msg_count: int,
    device_params: dict[str, dict[str, str]],
    expected: set[str],
    device_results: dict,
    value_results: dict,
    tol_pct: float,
    tol_abs: float,
    overall_ok: bool,
    now_str: str,
) -> str:
    reported = set(device_params.keys())
    oc = '#d4edda' if overall_ok else '#f8d7da'
    ol = 'PASS' if overall_ok else 'FAIL'
    otc = '#28a745' if overall_ok else '#dc3545'

    p2 = sum(1 for r in device_results.values() if r and r.get('ok'))
    f2 = sum(1 for r in device_results.values() if r and not r.get('ok'))
    w2 = sum(1 for r in device_results.values() if r is None)
    vp = sum(
        sum(1 for r in res if r.status == "PASS")
        for res in value_results.values() if isinstance(res, list)
    )
    vf = sum(
        sum(1 for r in res if r.status == "FAIL")
        for res in value_results.values() if isinstance(res, list)
    )
    vs = sum(1 for res in value_results.values() if isinstance(res, dict))

    s1 = _gen_device_list_section(reported, expected)
    s2 = _gen_param_unit_section(reported, device_results)
    s3 = _gen_value_section(value_results, tol_pct, tol_abs)

    content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Azure IoT 验证报告</title>
<style>{_CSS}</style>
</head>
<body>
<button id="err-toggle" onclick="toggleErrOnly()"
  style="position:fixed;top:16px;right:20px;z-index:999;padding:6px 16px;
         background:#dc3545;color:#fff;border:none;border-radius:4px;
         cursor:pointer;font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,.25)">
  仅显示异常
</button>
<script>
function toggleErrOnly(){{
  var btn=document.getElementById('err-toggle');
  var on=btn.dataset.active==='1';
  if(on){{
    document.querySelectorAll('details.section,details.subsection').forEach(function(d){{d.style.display=''}});
    document.querySelectorAll('tbody tr').forEach(function(tr){{tr.style.display=''}});
  }}else{{
    document.querySelectorAll('details.section,details.subsection').forEach(function(d){{d.open=true}});
    document.querySelectorAll('tbody tr').forEach(function(tr){{
      var bg=(tr.getAttribute('style')||'').toLowerCase();
      tr.style.display=(bg.indexOf('d4edda')!==-1||bg.indexOf('e8f4fd')!==-1)?'none':'';
    }});
    document.querySelectorAll('details.section,details.subsection').forEach(function(d){{
      var hasErr=d.dataset.hasError==='1'||
        d.querySelector('[data-has-error="1"]')!==null||
        Array.from(d.querySelectorAll('tbody tr')).some(function(tr){{
          var s=(tr.getAttribute('style')||'').toLowerCase();
          return tr.style.display!=='none'&&s.indexOf('d4edda')===-1&&s.indexOf('e8f4fd')===-1;
        }});
      if(!hasErr)d.style.display='none';
    }});
  }}
  btn.textContent=on?'仅显示异常':'显示全部';
  btn.style.background=on?'#dc3545':'#6c757d';
  btn.dataset.active=on?'0':'1';
}}
</script>
<h1>Azure IoT 消息验证报告</h1>
<div class="meta">
  生成时间：{html.escape(now_str)} &nbsp;|&nbsp;
  EventHub：{html.escape(config_info.get('eventhub_conn_str','')[:60])} &nbsp;|&nbsp;
  收到消息：{msg_count} 条 / 解析设备：{len(reported)} 台
</div>
<div class="overall-banner" style="background:{oc};color:{otc}">总体结论：{ol}</div>
<div class="cards">
  <div class="card total"><div class="num">{len(reported)}</div><div class="lbl">上报设备</div></div>
  <div class="card pass"><div class="num">{p2}</div><div class="lbl">参数通过</div></div>
  <div class="card fail"><div class="num">{f2}</div><div class="lbl">参数失败</div></div>
  <div class="card warn"><div class="num">{w2}</div><div class="lbl">无模板</div></div>
  <div class="card pass"><div class="num">{vp}</div><div class="lbl">数值通过</div></div>
  <div class="card fail"><div class="num">{vf}</div><div class="lbl">数值失败</div></div>
  <div class="card warn"><div class="num">{vs}</div><div class="lbl">数值跳过</div></div>
  <div class="card total"><div class="num">{msg_count}</div><div class="lbl">消息数</div></div>
</div>
{s1}
{s2}
{s3}
</body>
</html>"""

    os.makedirs(REPORT_DIR, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_path = os.path.join(REPORT_DIR, f'azure_iot_{ts}.html')
    Path(report_path).write_text(content, encoding='utf-8')
    return report_path


# ─── Core verifier ─────────────────────────────────────────────────────────────

def run_verifier(cfg: dict, timeout: int = 90, skip_web: bool = True) -> bool:
    """Run 4-stage Azure IoT verification. skip_web=True is the normal mode."""
    return _run_verifier_core(cfg, timeout)


def _run_verifier_core(cfg: dict, timeout: int) -> bool:
    azure    = cfg['azure_iot']
    expected = set(azure.get('expected_devices', []))

    tol_pct = float(azure.get('tolerance_pct', 5.0))
    tol_abs = float(azure.get('tolerance_abs', 1.0))
    modbus_device_map = cfg.get('modbus', {}).get('devices', {})

    device_params: dict[str, dict[str, str]]             = {}
    device_values: dict[str, dict[str, Optional[float]]] = {}
    msg_count = [0]
    done_event = threading.Event()
    grace_timer: list = [None]

    # ── 同步配对读取：消息到达时立刻读 Modbus，解决时序问题 ────────────────
    subscribe_ts = time.time()
    FRESH_WINDOW = 120          # 超过此秒数的消息视为缓存旧数据，不触发配对读取

    _paired_azure:  dict[str, dict]        = {}   # 触发 Modbus 时的 Azure 值快照
    _paired_modbus: dict[str, 'dict | None'] = {}
    _paired_err:    dict[str, str]         = {}
    _modbus_q:      _queue.Queue           = _queue.Queue()
    _modbus_lock    = threading.Lock()             # 串行化 Modbus 读取
    _queued_devices: set                   = set()

    def _modbus_worker():
        """后台线程：串行处理配对 Modbus 读取队列。"""
        while True:
            item = _modbus_q.get()
            if item is None:          # sentinel
                _modbus_q.task_done()
                break
            dev_name, param_keys, azure_snap = item
            with _modbus_lock:
                loop = asyncio.new_event_loop()
                try:
                    result, err = loop.run_until_complete(
                        _read_modbus_for_device(dev_name, param_keys, modbus_device_map)
                    )
                    _paired_azure[dev_name]  = azure_snap
                    _paired_modbus[dev_name] = result
                    _paired_err[dev_name]    = err
                    if err:
                        logging.warning(f"  [同步Modbus] {dev_name} 读取失败：{err}")
                    else:
                        logging.info(
                            f"  [同步Modbus] {dev_name}：读取 {len(result or {})} 个参数"
                            f"（与消息时间同步）"
                        )
                except Exception as exc:
                    _paired_azure[dev_name]  = azure_snap
                    _paired_modbus[dev_name] = None
                    _paired_err[dev_name]    = str(exc)
                    logging.warning(f"  [同步Modbus] {dev_name} 异常：{exc}")
                finally:
                    loop.close()
            _modbus_q.task_done()

    _worker_thread = threading.Thread(target=_modbus_worker, daemon=True)
    _worker_thread.start()

    def _is_fresh(payload: dict) -> bool:
        """消息 timestamp 在订阅时刻之前超过 FRESH_WINDOW 秒则视为缓存旧消息。"""
        ts = payload.get('timestamp')
        if ts is None:
            return True
        age = subscribe_ts - float(ts)
        return age <= FRESH_WINDOW

    def _schedule_stop():
        def _fire():
            logging.info("宽限期结束，停止收集")
            done_event.set()
        t = threading.Timer(5.0, _fire)
        t.daemon = True
        t.start()
        grace_timer[0] = t
        logging.info("所有期望设备已就绪，5 秒后停止收集剩余分块…")

    def _process_payload(payload: dict) -> None:
        """处理单条消息 payload，更新 device_params / device_values 并触发配对 Modbus。"""
        msg_count[0] += 1
        preview = str(payload)[:200]
        logging.info(f"[消息 #{msg_count[0]}] {preview}")
        fresh = _is_fresh(payload)
        if not fresh:
            msg_ts = payload.get('timestamp', 0)
            age = subscribe_ts - float(msg_ts)
            logging.info(f"  -> 旧缓存消息（{age:.0f}s 前推送），跳过数值配对")
        for mod in _parse_modules(payload):
            name = mod['name']
            if name not in device_params:
                device_params[name] = {}
                device_values[name] = {}
            for item in mod['reading']:
                param = (item.get('param') or '').strip()
                if not param:
                    continue
                device_params[name][param] = str(item.get('unit', '')).strip()
                try:
                    device_values[name][param] = float(item['value'])
                except (KeyError, ValueError, TypeError):
                    if param not in device_values[name]:
                        device_values[name][param] = None
            logging.info(f"  -> 设备：{name}，本块 {len(mod['reading'])} 个参数，"
                         f"累计 {len(device_params[name])} 个")
            # 新鲜消息：立刻触发该设备的同步 Modbus 读取（每台设备仅触发一次）
            if fresh and name not in _queued_devices:
                _queued_devices.add(name)
                param_keys = [k for k, v in device_values[name].items() if v is not None]
                if param_keys:
                    azure_snap = dict(device_values[name])  # 当前消息的值快照
                    _modbus_q.put((name, param_keys, azure_snap))
                    logging.info(f"  -> [{name}] 已触发同步 Modbus 读取")
        if expected and device_params.keys() >= expected and grace_timer[0] is None:
            _schedule_stop()

    # ── 通过 Azure Event Hub 兼容端点接收消息 ────────────────────────────────
    logging.info(f"正在连接 Azure Event Hub，等待消息（超时 {timeout}s）…")
    raw_messages = _receive_messages(cfg, timeout)
    done_event.set()

    for payload in raw_messages:
        _process_payload(payload)

    # 等待后台 Modbus 配对读取全部完成
    logging.info("等待同步 Modbus 读取完成…")
    _modbus_q.put(None)   # 发送结束哨兵
    _modbus_q.join()
    _worker_thread.join(timeout=30)

    if not raw_messages:
        print("\n[FAIL] 超时内未收到任何消息，请检查网络/EventHub 连接字符串配置")
        return False
    if not device_params:
        print("\n[WARN] 收到消息但无法解析设备数据")
        return False

    SEP  = "=" * 70
    SEP2 = "-" * 70
    overall_ok = True
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    print(f"\n{SEP}")
    print(f"共收到 {len(raw_messages)} 条消息，解析到 {len(device_params)} 台设备")
    print(SEP)

    # ── [1] 设备列表 ──────────────────────────────────────────────────────────
    print("\n[1] 设备列表检查")
    reported  = set(device_params.keys())
    missing_d = expected - reported
    extra_d   = reported - expected
    if not missing_d and not extra_d:
        print(f"  [OK] 上报设备与期望设备完全一致：{sorted(reported)}")
    else:
        overall_ok = False
        if missing_d:
            print(f"  [FAIL] 缺少设备：{sorted(missing_d)}")
        if extra_d:
            print(f"  [WARN] 多余设备：{sorted(extra_d)}")

    # ── [2][3] 参数完整性 & 单位 ──────────────────────────────────────────────
    print("\n[2][3] 参数完整性 & 单位检查（模板 MQTT 列）")
    device_results: dict[str, dict | None] = {}
    for dev_name in sorted(reported):
        tmpl_map = _load_mqtt_template(dev_name)
        print(f"\n{SEP2}")
        print(f"  设备：{dev_name}  累计上报参数：{len(device_params[dev_name])} 个")
        if tmpl_map is None:
            print(f"  [WARN] 未找到 MQTT 模板，跳过参数检查")
            device_results[dev_name] = None
            continue
        r = _check_params(device_params[dev_name], tmpl_map)
        device_results[dev_name] = r
        if r['ok']:
            print(f"  [OK] 参数完整（{r['tmpl_total']} 个）且单位全部一致")
        else:
            overall_ok = False
        if r['missing']:
            print(f"  [FAIL] 缺少参数 {len(r['missing'])} 个")
            for p in r['missing'][:20]:
                print(f"         {p:<40} 模板单位: {tmpl_map[p]!r}")
        if r['extra']:
            print(f"  [WARN] 多余参数 {len(r['extra'])} 个")
        if r['unit_mismatch']:
            print(f"  [FAIL] 单位不一致 {len(r['unit_mismatch'])} 个")
            for param, t_u, j_u in r['unit_mismatch'][:20]:
                print(f"         {param:<40} {t_u!r} -> {j_u!r}")

    # ── [4] 数值比对 ──────────────────────────────────────────────────────────
    print(f"\n[4] 数值比对（Azure IoT vs 实时 Modbus，容差 ±{tol_pct}% / ±{tol_abs}）")

    value_results = asyncio.run(_compare_all_devices(
        device_values, tol_pct, tol_abs,
        modbus_device_map=modbus_device_map,
        preread_modbus=_paired_modbus,
        preread_errors=_paired_err,
        paired_azure=_paired_azure,
    ))

    for dev_name in sorted(value_results.keys()):
        res = value_results[dev_name]
        print(f"\n{SEP2}")
        print(f"  设备：{dev_name}")
        if isinstance(res, dict):
            print(f"  [WARN] 跳过：{res.get('skip', res.get('error', ''))}")
            continue
        dp = sum(1 for r in res if r.status == "PASS")
        df = sum(1 for r in res if r.status == "FAIL")
        de = len(res) - dp - df
        pr = f"{dp/len(res)*100:.1f}%" if res else "N/A"
        if df:
            overall_ok = False
            print(f"  [FAIL] PASS={dp} ({pr})  FAIL={df}  ERR={de}")
            for r in res:
                if r.status == "FAIL":
                    print(f"         {r.param_key:<40} "
                          f"Azure={_fmt(r.azure_value)}  Modbus={_fmt(r.modbus_value)}  "
                          f"Δ%={r.diff_pct:.2f}%")
        else:
            print(f"  [OK] PASS={dp} ({pr})  ERR={de}")

    print(f"\n{SEP}")
    print("总体结论：[PASS] 全部通过" if overall_ok else "总体结论：[FAIL] 存在问题，请检查上述详情")
    print(SEP)

    report_path = generate_html_report(
        config_info    = azure,
        msg_count      = len(raw_messages),
        device_params  = device_params,
        expected       = expected,
        device_results = device_results,
        value_results  = value_results,
        tol_pct        = tol_pct,
        tol_abs        = tol_abs,
        overall_ok     = overall_ok,
        now_str        = now_str,
    )
    print(f"\nHTML 报告已保存：{report_path}")
    return overall_ok


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default="tests/azure_iot/config.yaml")
    parser.add_argument("--timeout", type=int, default=90)
    args = parser.parse_args()
    cfg = load_config(args.config)
    ok = run_verifier(cfg, timeout=args.timeout)
    sys.exit(0 if ok else 1)
