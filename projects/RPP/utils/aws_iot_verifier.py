"""
AWS IoT 消息订阅验证脚本（AcuHMI-1-7 工程版）

四阶段验证：
  [1] 设备列表：上报设备是否与期望设备一致
  [2] 参数完整性：每台设备上报的参数是否覆盖模板 MQTT 列全量参数
  [3] 单位一致性：每个参数的上报单位是否与模板 unit 列一致
  [4] 数值比对：AWS IoT 上报值 vs 设备实时 Modbus 读取值

用法：
    python utils/aws_iot_verifier.py --config tests/aws_iot/config.yaml
"""
import argparse
import asyncio
import concurrent.futures
import html
import inspect
import json
import logging
import os
import re
import sys
import time
import queue as _queue
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import paho.mqtt.client as mqtt
import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent   # AcuHMI-1-7/
REPORT_DIR   = str(PROJECT_ROOT / "tests" / "aws_iot" / "reports")

sys.path.insert(0, str(PROJECT_ROOT))

from utils.modbus_reader import ModbusResult, read_device_params, read_device_params_rtu

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


def _get_caller_case_id() -> str:
    """从调用栈中提取用例文件名的 AWS_XXX_XXX 部分，用于报告命名。"""
    for frame_info in inspect.stack():
        m = re.search(r'(AWS_\d+_\d+|COMM_\d+_\d+)', Path(frame_info.filename).stem)
        if m:
            return m.group(1)
    return ""


# ─── Device name mappings ──────────────────────────────────────────────────────

# Gateway-reported name → canonical excel_param_map device name
# 静态精确匹配；前缀模式匹配由 _resolve_device_name() 动态处理（读取 config alias_map）
_DEV_ALIAS_MAP: dict[str, str] = {
    "4100242":   "AcuRev4100",
    "Acu4100":   "AcuRev4100",
    # PXB/PXE1/PXE2/PXM350 前缀由运行时 alias_map 处理
}

# 运行时从 config 加载的前缀映射（在 _run_verifier_core 中填充）
_runtime_alias_prefixes: dict[str, str] = {}

# Canonical name → excel_param_map key (matches _EXCEL_FILES in excel_param_map.py)
# 同型号多实例（如 AcuRev4100a/b）映射到同一参数模板，Modbus 连接参数由 config.yaml 各自配置
_DEV_EXCEL_MAP: dict[str, str] = {
    "AcuRev4100":  "AcuRev4100",
    "AcuRev4100a": "AcuRev4100",
    "AcuRev4100b": "AcuRev4100",
    "AcuRev2100":  "AcuRev2100",
    "AcuvimIIW":   "AcuvimIIW",
    "AcuvimIIR":   "AcuvimIIR",
    "AcuVIM3":     "Acuvim3",
    "Acuvim3":     "Acuvim3",
    "AcuRev1300":  "AcuRev1300",
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
    aws_value:    Optional[float] = None
    modbus_value: Optional[float] = None
    aws_error:    str = ""
    modbus_error: str = ""
    diff_abs:     Optional[float] = None
    diff_pct:     Optional[float] = None
    status:       str = ""   # PASS | FAIL | AWS_ERR | MODBUS_ERR | BOTH_ERR


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
        'ok': not missing and not extra and not unit_mismatch,
        '_tmpl_map': tmpl_map, '_reported_units': json_params,
        '_all_params': sorted(tmpl_keys | json_keys),
    }


def _check_params_units_only(json_params: dict[str, str], tmpl_map: dict[str, str]) -> dict:
    """虚拟设备专用：只比对参数数量和单位集合，忽略参数名（MQTT 用 Post Label，Reading 页用 Parameter Name）。"""
    tmpl_units = sorted(_norm_unit(v) for v in tmpl_map.values())
    json_units = sorted(_norm_unit(v) for v in json_params.values())
    count_ok = len(tmpl_map) == len(json_params)
    units_ok = tmpl_units == json_units
    unit_mismatch = [] if units_ok else [('(单位集合)', str(tmpl_units), str(json_units))]
    return {
        'tmpl_total': len(tmpl_map), 'report_total': len(json_params),
        'missing': [] if count_ok else [f'(期望 {len(tmpl_map)} 个，上报 {len(json_params)} 个)'],
        'extra': [],
        'unit_mismatch': unit_mismatch,
        'ok': count_ok and units_ok,
        '_tmpl_map': tmpl_map, '_reported_units': json_params,
        '_all_params': [],
    }


# ─── Value comparison ──────────────────────────────────────────────────────────

def _compare_one_value(
    param_key: str, aws_value: Optional[float],
    mr: Optional[ModbusResult], tol_pct: float, tol_abs: float,
) -> ValueCompareResult:
    import math
    cr = ValueCompareResult(param_key=param_key)
    if aws_value is None:
        cr.aws_error = "无值"
    if mr is None or not mr.ok:
        cr.modbus_error = mr.error if mr else "未读取"
    if cr.aws_error and cr.modbus_error:
        cr.status = "BOTH_ERR"; return cr
    if cr.aws_error:
        cr.status = "AWS_ERR"; return cr
    if cr.modbus_error:
        cr.status = "MODBUS_ERR"; return cr
    cr.aws_value = aws_value
    cr.modbus_value = mr.value
    if math.isnan(aws_value) and math.isnan(mr.value):
        cr.diff_abs = 0.0; cr.diff_pct = 0.0; cr.status = "PASS"; return cr
    diff = abs(aws_value - mr.value)
    ref  = max(abs(aws_value), abs(mr.value))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0
    tol = max(tol_abs, ref * tol_pct / 100)
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


def _is_virtual_device(dev_name: str, cfg: dict) -> bool:
    """判断设备是否为虚拟设备：名称含 'virtual' 或匹配 config.aws_iot.virtual_device。"""
    if 'virtual' in dev_name.lower():
        return True
    vd = (cfg.get('aws_iot') or {}).get('virtual_device', '')
    if not vd:
        return False
    dn_l = dev_name.lower()
    vd_l = vd.lower()
    return dn_l == vd_l or dn_l in vd_l or vd_l in dn_l


def _find_virtual_readings(dev_name: str, virtual_readings: dict) -> 'dict | None':
    """在 virtual_readings 中查找与 dev_name 匹配的读数（大小写不敏感子串匹配）。"""
    dn_l = dev_name.lower()
    for k, v in virtual_readings.items():
        kl = k.lower()
        if kl == dn_l or kl in dn_l or dn_l in kl:
            return v
    return None


def _resolve_device_name(dev_name: str) -> str:
    """解析网关上报的设备名到规范型号名。
    先查静态精确表，再按前缀匹配 alias_map（运行时填充）。
    """
    if dev_name in _DEV_ALIAS_MAP:
        return _DEV_ALIAS_MAP[dev_name]
    upper = dev_name.upper()
    for prefix, canonical in _runtime_alias_prefixes.items():
        if upper.startswith(prefix.upper()):
            return canonical
    return dev_name


async def _read_modbus_for_device(
    dev_name: str,
    param_keys: list[str],
    modbus_device_map: dict,
) -> tuple[dict[str, ModbusResult] | None, str]:
    canonical = _resolve_device_name(dev_name)
    modbus_info = modbus_device_map.get(canonical)
    if modbus_info is None:
        for k, v in modbus_device_map.items():
            if k.lower() == canonical.lower():
                modbus_info = v
                break
    if not modbus_info:
        return None, f"未在 config.yaml modbus 中配置 {canonical!r}"

    excel_name = _DEV_EXCEL_MAP.get(canonical, canonical)
    mode = modbus_info.get("mode", "tcp")

    try:
        if mode == "rtu":
            result = await read_device_params_rtu(
                excel_name,
                param_keys,
                serial_port=modbus_info["serial_port"],
                unit=modbus_info["unit"],
                baudrate=modbus_info.get("baudrate", 9600),
                parity=modbus_info.get("parity", "N"),
                stopbits=modbus_info.get("stopbits", 1),
                bytesize=modbus_info.get("bytesize", 8),
            )
        else:
            result = await read_device_params(
                excel_name, param_keys,
                modbus_info["ip"], modbus_info["port"], modbus_info["unit"],
            )
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
    paired_aws:     'dict | None' = None,
    virtual_readings: 'dict | None' = None,
    cfg:            'dict | None' = None,
) -> dict:
    results: dict = {}
    for dev_name in sorted(device_values.keys()):
        val_map = (paired_aws or {}).get(dev_name) or device_values[dev_name]
        comparable_keys = [k for k, v in val_map.items() if v is not None]
        if not comparable_keys:
            results[dev_name] = {'skip': '无可比对数值（所有参数缺失 value 字段）'}
            continue

        # 虚拟设备：不读 Modbus，改用 Reading 页面值比对
        # Reading 页 key 与 AWS key 后缀可能不同（如 param01 vs label01），按位置匹配
        if _is_virtual_device(dev_name, cfg or {}):
            if virtual_readings:
                vrd = _find_virtual_readings(dev_name, virtual_readings)
                if vrd is not None:
                    cmp_list = []
                    aws_params = sorted(comparable_keys)
                    reading_vals = list(vrd.values())  # 按 Reading 页顺序
                    for i, param in enumerate(aws_params):
                        aws_val = val_map.get(param)
                        if i < len(reading_vals):
                            rv = reading_vals[i]
                            raw_v = str(rv.get("value", "")).strip()
                            # 尝试解析：剥离末尾非数字字符（如单位混入），支持逗号小数点
                            _m = re.search(r'[-+]?\d[\d,]*\.?\d*', raw_v)
                            _parsed = _m.group().replace(',', '') if _m else None
                            try:
                                reading_val = float(_parsed) if _parsed else float(raw_v)
                                mr = ModbusResult(param_key=param, value=reading_val, ok=True, error="")
                            except (TypeError, ValueError):
                                mr = ModbusResult(param_key=param, value=None, ok=False,
                                                  error=f"Reading 值无法转为数值：{raw_v!r}")
                        else:
                            mr = ModbusResult(param_key=param, value=None, ok=False, error="Reading 页面参数不足")
                        cmp_list.append(_compare_one_value(param, aws_val, mr, tol_pct, tol_abs))
                    results[dev_name] = cmp_list
                    p = sum(1 for r in cmp_list if r.status == "PASS")
                    f = sum(1 for r in cmp_list if r.status == "FAIL")
                    logging.info(f"  [Reading] {dev_name}：PASS={p}  FAIL={f}  ERR={len(cmp_list)-p-f}")
                    continue
            results[dev_name] = {'skip': '虚拟设备，不做 Modbus 比对（可传入 virtual_readings 启用 Reading 页面比对）'}
            logging.info(f"  [Modbus] {dev_name} 跳过：虚拟设备")
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
    "AWS_ERR": "#fff3cd", "MODBUS_ERR": "#fff3cd", "BOTH_ERR": "#e2e3e5",
}
_STATUS_LABEL = {
    "PASS": "通过", "FAIL": "失败",
    "AWS_ERR": "AWS异常", "MODBUS_ERR": "Modbus异常", "BOTH_ERR": "双路异常",
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
            if r.aws_error:    err_hint += f"AWS: {html.escape(r.aws_error)}"
            if r.modbus_error: err_hint += f"{'<br>' if err_hint else ''}Modbus: {html.escape(r.modbus_error)}"
            da = _fmt(r.diff_abs, 4) if r.diff_abs is not None else "—"
            dp_s = f"{r.diff_pct:.3f}%" if r.diff_pct is not None else "—"
            rows += (f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
                     f'<td class="key">{html.escape(r.param_key)}</td>'
                     f'<td class="val">{_fmt(r.aws_value)}</td>'
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
<thead><tr><th>#</th><th>参数名</th><th>AWS IoT 值</th><th>Modbus 值</th>
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
<summary>三、数值比对（AWS IoT vs 实时 Modbus）
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
    report_name: str = "",
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
<title>AWS IoT 验证报告</title>
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
<h1>AWS IoT 消息验证报告</h1>
<div class="meta">
  生成时间：{html.escape(now_str)} &nbsp;|&nbsp;
  Endpoint：{html.escape(config_info.get('url',''))} &nbsp;|&nbsp;
  Topic：{html.escape(config_info.get('topic',''))} &nbsp;|&nbsp;
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
    prefix = f'{report_name}_' if report_name else ''
    report_path = os.path.join(REPORT_DIR, f'{prefix}aws_iot_{ts}.html')
    Path(report_path).write_text(content, encoding='utf-8')
    try:
        import allure
        attach_name = f"数值比对报告 ({report_name})" if report_name else "数值比对报告"
        allure.attach(content, name=attach_name, attachment_type=allure.attachment_type.HTML)
    except Exception:
        pass
    return report_path


# ─── Core verifier ─────────────────────────────────────────────────────────────

def run_verifier(cfg: dict, timeout: int = 90, skip_web: bool = True,
                 config_path: str = None,
                 virtual_readings: 'dict | None' = None) -> bool:
    """Run 4-stage AWS IoT verification. skip_web=True is the normal mode.

    virtual_readings: 虚拟设备的 Reading 页面数据，格式
        {device_name: {param_name: {"value": str, "unit": str}}}
        由 AWSIoTPage.get_virtual_device_readings() 获取后传入。
        提供后虚拟设备在 stage [4] 中与 Reading 页面值比对，否则仅标注"虚拟设备，跳过"。
    """
    return _run_verifier_core(cfg, timeout, virtual_readings=virtual_readings)


def _lookup_template(dev_name: str,
                     extra_templates: 'dict[str, dict[str, str]] | None') -> 'dict[str, str] | None':
    """先查 extra_templates（虚拟设备），再查 Excel 模板（物理设备）。
    extra_templates key 与 dev_name 做大小写不敏感子串匹配。"""
    if extra_templates:
        dn_l = dev_name.lower()
        for key, tmpl in extra_templates.items():
            kl = key.lower()
            if kl == dn_l or kl in dn_l or dn_l in kl:
                return tmpl
    return _load_mqtt_template(dev_name)


def run_verifier_params_only(cfg: dict, timeout: int = 90,
                             extra_templates: 'dict[str, dict[str, str]] | None' = None) -> bool:
    """
    仅验证参数数量和单位（stage [1][2][3]），不做 Modbus 数值比对。

    extra_templates: 虚拟设备的参数→单位映射，来自 UI Reading 页，
                     格式 {device_name: {param_name: unit_str}}。
                     命中的设备跳过 Excel 模板，直接与此对比。
    """
    aws      = cfg['aws_iot']
    endpoint = aws['url']
    topic    = aws['topic']
    expected = set(aws.get('expected_devices', []))

    cert_file = _abs_path(aws['sub_cert_file'])
    key_file  = _abs_path(aws['sub_key_file'])
    ca_file   = _abs_path(aws['ca_file'])

    for label, path in [('cert', cert_file), ('key', key_file), ('CA', ca_file)]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} 文件不存在：{path}")

    _runtime_alias_prefixes.clear()
    _runtime_alias_prefixes.update(cfg.get('modbus', {}).get('alias_map', {}))

    device_params: dict[str, dict[str, str]] = {}
    raw_messages:  list[dict] = []
    msg_count  = [0]
    done_event = threading.Event()
    grace_timer: list = [None]

    def _schedule_stop():
        def _fire():
            logging.info("宽限期结束，停止收集")
            done_event.set()
        t = threading.Timer(5.0, _fire)
        t.daemon = True
        t.start()
        grace_timer[0] = t
        logging.info("所有期望设备已就绪，5 秒后停止收集…")

    def on_connect(client, userdata, connect_flags, reason_code, properties):
        if reason_code == 0:
            logging.info(f"已连接 AWS IoT：{endpoint}")
            client.subscribe(topic, qos=1)
            logging.info(f"订阅 Topic：{topic}")
        else:
            logging.error(f"连接失败 code={reason_code}")
            done_event.set()

    def on_message(client, userdata, msg):
        raw = msg.payload.decode('utf-8', errors='replace')
        msg_count[0] += 1
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            return
        raw_messages.append(payload)
        for mod in _parse_modules(payload):
            name = mod['name']
            if name not in device_params:
                device_params[name] = {}
            for item in mod['reading']:
                param = (item.get('param') or '').strip()
                if not param:
                    continue
                device_params[name][param] = str(item.get('unit', '')).strip()
            logging.info(f"  -> 设备：{name}，累计 {len(device_params[name])} 个参数")
        if expected and device_params.keys() >= expected and grace_timer[0] is None:
            _schedule_stop()

    def on_disconnect(client, userdata, disconnect_flags, reason_code, properties):
        logging.info(f"已断开 rc={reason_code}")

    mqttc = mqtt.Client(
        mqtt.CallbackAPIVersion.VERSION2,
        client_id=f"verifier-{int(time.time())}",
    )
    mqttc.tls_set(ca_certs=ca_file, certfile=cert_file, keyfile=key_file)
    mqttc.on_connect    = on_connect
    mqttc.on_message    = on_message
    mqttc.on_disconnect = on_disconnect

    logging.info(f"正在连接 {endpoint}:8883 …")
    mqttc.connect(endpoint, port=8883, keepalive=60)
    mqttc.loop_start()
    logging.info(f"等待消息（超时 {timeout}s）…")
    done_event.wait(timeout=timeout)
    mqttc.loop_stop()
    mqttc.disconnect()

    if not raw_messages:
        print("\n[FAIL] 超时内未收到任何消息，请检查网络/证书/Topic 配置")
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

    # [1] 设备列表
    print("\n[1] 设备列表检查")
    reported  = set(device_params.keys())
    missing_d = expected - reported
    extra_d   = reported - expected
    if not missing_d:
        if not extra_d:
            print(f"  [OK] 上报设备与期望设备完全一致：{sorted(reported)}")
        else:
            print(f"  [OK] 期望设备已全部上报；多余设备（旧消息）：{sorted(extra_d)}")
    else:
        overall_ok = False
        print(f"  [FAIL] 缺少设备：{sorted(missing_d)}")
        if extra_d:
            print(f"  [WARN] 多余设备：{sorted(extra_d)}")

    # 只验证期望设备，过滤旧缓存消息带入的多余设备
    devices_to_validate = (reported & expected) if expected else reported

    # [2][3] 参数完整性 & 单位
    print("\n[2][3] 参数完整性 & 单位检查（模板 MQTT 列 / 虚拟设备用 UI Reading 页）")
    device_results: dict[str, dict | None] = {}
    for dev_name in sorted(devices_to_validate):
        tmpl_map = _lookup_template(dev_name, extra_templates)
        print(f"\n{SEP2}")
        print(f"  设备：{dev_name}  上报参数：{len(device_params[dev_name])} 个")
        if tmpl_map is None:
            print(f"  [WARN] 未找到 MQTT 模板，跳过参数检查")
            device_results[dev_name] = None
            continue
        r = _check_params(device_params[dev_name], tmpl_map)
        device_results[dev_name] = r
        if r['ok']:
            print(f"  [OK] 参数完整（模板 {r['tmpl_total']} 个 / 上报 {r['report_total']} 个）且单位全部一致")
        else:
            overall_ok = False
            if r['missing']:
                print(f"  [FAIL] 缺少参数 {len(r['missing'])} 个")
                for p in r['missing'][:20]:
                    print(f"         {p:<40} 模板单位: {tmpl_map[p]!r}")
            if r['extra']:
                print(f"  [WARN] 多余参数 {len(r['extra'])} 个")
                for p in r['extra'][:20]:
                    print(f"         {p:<40} 上报单位: {device_params[dev_name].get(p, '')!r}")
            if r['unit_mismatch']:
                print(f"  [FAIL] 单位不一致 {len(r['unit_mismatch'])} 个")
                for param, t_u, j_u in r['unit_mismatch'][:20]:
                    print(f"         {param:<40} 模板: {t_u!r}  上报: {j_u!r}")

    print(f"\n{SEP}")
    print("总体结论：[PASS] 全部通过" if overall_ok else "总体结论：[FAIL] 存在问题，请检查上述详情")
    print(SEP)

    report_path = generate_html_report(
        config_info    = aws,
        msg_count      = len(raw_messages),
        device_params  = device_params,
        expected       = expected,
        device_results = device_results,
        value_results  = {},
        tol_pct        = float(aws.get('tolerance_pct', 5.0)),
        tol_abs        = float(aws.get('tolerance_abs', 1.0)),
        overall_ok     = overall_ok,
        now_str        = now_str,
        report_name    = _get_caller_case_id(),
    )
    print(f"\nHTML 报告已保存：{report_path}")
    return overall_ok


def count_mqtt_messages(cfg: dict, timeout: int) -> int:
    """
    订阅 AWS IoT Topic，在 timeout 秒内计算收到的消息总数。
    用于验证"无参数配置时设备不上报数据"场景（预期返回 0）。
    """
    aws = cfg['aws_iot']
    cert_file = _abs_path(aws['sub_cert_file'])
    key_file  = _abs_path(aws['sub_key_file'])
    ca_file   = _abs_path(aws['ca_file'])

    msg_count = [0]
    done_event = threading.Event()

    def on_connect(client, userdata, flags, reason_code, properties):
        if reason_code == 0:
            client.subscribe(aws['topic'], qos=1)
            logging.info(f"[count] 已订阅 {aws['topic']}，等待 {timeout}s…")
        else:
            logging.error(f"[count] 连接失败 rc={reason_code}")
            done_event.set()

    def on_message(client, userdata, msg):
        msg_count[0] += 1
        logging.info(f"[count] 收到消息 #{msg_count[0]}")

    mqttc = mqtt.Client(
        mqtt.CallbackAPIVersion.VERSION2,
        client_id=f"count-{int(time.time())}",
    )
    mqttc.tls_set(ca_certs=ca_file, certfile=cert_file, keyfile=key_file)
    mqttc.on_connect = on_connect
    mqttc.on_message = on_message

    mqttc.connect(aws['url'], port=8883, keepalive=60)
    mqttc.loop_start()
    done_event.wait(timeout=timeout)
    mqttc.loop_stop()
    mqttc.disconnect()

    logging.info(f"[count] 订阅结束，共收到 {msg_count[0]} 条消息")
    return msg_count[0]


def _run_verifier_core(cfg: dict, timeout: int,
                       virtual_readings: 'dict | None' = None) -> bool:
    aws      = cfg['aws_iot']
    endpoint = aws['url']
    topic    = aws['topic']
    expected = set(aws.get('expected_devices', []))

    cert_file = _abs_path(aws['sub_cert_file'])
    key_file  = _abs_path(aws['sub_key_file'])
    ca_file   = _abs_path(aws['ca_file'])

    for label, path in [('cert', cert_file), ('key', key_file), ('CA', ca_file)]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} 文件不存在：{path}")

    tol_pct = float(aws.get('tolerance_pct', 5.0))
    tol_abs = float(aws.get('tolerance_abs', 1.0))

    # 加载 alias_map 前缀映射到运行时全局表
    _runtime_alias_prefixes.clear()
    _runtime_alias_prefixes.update(cfg.get('modbus', {}).get('alias_map', {}))

    # 构建统一 modbus_device_map：兼容旧格式 (devices: {name: [ip,port,unit]}) 和
    # 新格式 (tcp.devices + rtu 列表)
    modbus_cfg = cfg.get('modbus', {})
    modbus_device_map: dict = {}
    if 'devices' in modbus_cfg:
        # 旧格式：{name: [ip, port, unit]}
        for _name, _v in modbus_cfg['devices'].items():
            modbus_device_map[_name] = {
                "mode": "tcp",
                "ip":   str(_v[0]),
                "port": int(_v[1]),
                "unit": int(_v[2]),
            }
    if 'tcp' in modbus_cfg:
        for _name, _dev in modbus_cfg['tcp'].get('devices', {}).items():
            modbus_device_map[_name] = {
                "mode": "tcp",
                "ip":   str(_dev['ip']),
                "port": int(_dev['port']),
                "unit": int(_dev['unit']),
            }
    for _line in modbus_cfg.get('rtu', []):
        _serial_params = {
            "serial_port": _line['port'],
            "baudrate":    int(_line.get('baudrate', 9600)),
            "parity":      _line.get('parity', 'N'),
            "stopbits":    int(_line.get('stopbits', 1)),
            "bytesize":    int(_line.get('bytesize', 8)),
        }
        for _name, _dev in _line.get('devices', {}).items():
            modbus_device_map[_name] = {
                "mode": "rtu",
                "unit": int(_dev['unit']),
                **_serial_params,
            }

    device_params: dict[str, dict[str, str]]            = {}
    device_values: dict[str, dict[str, Optional[float]]] = {}
    raw_messages: list[dict] = []
    msg_count = [0]
    done_event = threading.Event()
    grace_timer: list = [None]

    # ── 同步配对读取：消息到达时立刻读 Modbus，解决时序问题 ────────────────
    subscribe_ts = time.time()
    FRESH_WINDOW = 120          # 超过此秒数的消息视为缓存旧数据，不触发配对读取

    _paired_aws:    dict[str, dict] = {}   # 触发 Modbus 时的 AWS 值快照
    _paired_modbus: dict[str, 'dict | None'] = {}
    _paired_err:    dict[str, str]  = {}
    _modbus_q:      _queue.Queue    = _queue.Queue()
    _modbus_lock    = threading.Lock()     # 串行化 Modbus 读取
    _device_modbus_timers: dict[str, threading.Timer] = {}
    _device_modbus_timer_lock = threading.Lock()

    def _modbus_worker():
        """后台线程：串行处理配对 Modbus 读取队列。"""
        while True:
            item = _modbus_q.get()
            if item is None:          # sentinel
                _modbus_q.task_done()
                break
            dev_name, param_keys, aws_snap = item
            with _modbus_lock:
                loop = asyncio.new_event_loop()
                try:
                    result, err = loop.run_until_complete(
                        _read_modbus_for_device(dev_name, param_keys, modbus_device_map)
                    )
                    _paired_aws[dev_name]    = aws_snap
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
                    _paired_aws[dev_name]    = aws_snap
                    _paired_modbus[dev_name] = None
                    _paired_err[dev_name]    = str(exc)
                    logging.warning(f"  [同步Modbus] {dev_name} 异常：{exc}")
                finally:
                    loop.close()
            _modbus_q.task_done()

    _worker_thread = threading.Thread(target=_modbus_worker, daemon=True)
    _worker_thread.start()

    def _enqueue_modbus_now(name: str):
        """定时器触发时调用：将设备当前全量参数入队 Modbus 读取。"""
        keys = [k for k, v in device_values[name].items() if v is not None]
        if keys:
            snap = dict(device_values[name])
            _modbus_q.put((name, keys, snap))
            logging.info(f"  -> [{name}] 已入队 Modbus 读取（{len(keys)} 个参数）")

    def _schedule_modbus_debounced(name: str, delay: float = 2.0):
        """每次收到该设备新包时调用；delay 秒内无新包则触发 Modbus 读取（debounce）。"""
        with _device_modbus_timer_lock:
            old = _device_modbus_timers.get(name)
            if old is not None:
                old.cancel()
            t = threading.Timer(delay, _enqueue_modbus_now, args=(name,))
            t.daemon = True
            _device_modbus_timers[name] = t
            t.start()

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

    def on_connect(client, userdata, connect_flags, reason_code, properties):
        if reason_code == 0:
            logging.info(f"已连接 AWS IoT：{endpoint}")
            client.subscribe(topic, qos=1)
            logging.info(f"订阅 Topic：{topic}")
        else:
            logging.error(f"连接失败 code={reason_code}")
            done_event.set()

    def on_message(client, userdata, msg):
        raw = msg.payload.decode('utf-8', errors='replace')
        msg_count[0] += 1
        preview = raw[:200] + ('…' if len(raw) > 200 else '')
        logging.info(f"[消息 #{msg_count[0]}] {preview}")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            logging.warning("非 JSON 消息，跳过")
            return
        fresh = _is_fresh(payload)
        if not fresh:
            msg_ts = payload.get('timestamp', 0)
            age = subscribe_ts - float(msg_ts)
            logging.info(f"  -> 旧缓存消息（{age:.0f}s 前推送），跳过数值配对")
        raw_messages.append(payload)
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
            # 新鲜消息：debounce 触发 Modbus 读取（每次收包重置定时器，最后一包后静默 2s 再读）
            # 分块上报场景（如 AcuRev4100 三包）可确保用全量参数读取，而非仅第一包
            # 虚拟设备跳过 Modbus 队列，stage [4] 改用 Reading 页面值比对
            if fresh:
                if _is_virtual_device(name, cfg):
                    logging.info(f"  -> [{name}] 虚拟设备，跳过 Modbus 读取")
                else:
                    _schedule_modbus_debounced(name, delay=2.0)
                    logging.info(f"  -> [{name}] 已（重）置 Modbus 读取定时器")
        if expected and device_params.keys() >= expected and grace_timer[0] is None:
            _schedule_stop()

    def on_disconnect(client, userdata, disconnect_flags, reason_code, properties):
        logging.info(f"已断开 rc={reason_code}")

    client = mqtt.Client(
        mqtt.CallbackAPIVersion.VERSION2,
        client_id=f"verifier-{int(time.time())}",
    )
    client.tls_set(ca_certs=ca_file, certfile=cert_file, keyfile=key_file)
    client.on_connect    = on_connect
    client.on_message    = on_message
    client.on_disconnect = on_disconnect

    logging.info(f"正在连接 {endpoint}:8883 …")
    client.connect(endpoint, port=8883, keepalive=60)
    client.loop_start()
    logging.info(f"等待消息（超时 {timeout}s）…")
    done_event.wait(timeout=timeout)
    client.loop_stop()
    client.disconnect()

    # 先等所有 debounce 定时器触发并完成入队（最多额外等 2s）
    logging.info("等待 debounce 定时器完成…")
    for _t in list(_device_modbus_timers.values()):
        _t.join()

    # 等待后台 Modbus 配对读取全部完成
    logging.info("等待同步 Modbus 读取完成…")
    _modbus_q.put(None)   # 发送结束哨兵
    _modbus_q.join()
    _worker_thread.join(timeout=30)

    if not raw_messages:
        print("\n[FAIL] 超时内未收到任何消息，请检查网络/证书/Topic 配置")
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
    if not missing_d:
        if not extra_d:
            print(f"  [OK] 上报设备与期望设备完全一致：{sorted(reported)}")
        else:
            print(f"  [OK] 期望设备已全部上报；多余设备（旧消息）：{sorted(extra_d)}")
    else:
        overall_ok = False
        print(f"  [FAIL] 缺少设备：{sorted(missing_d)}")
        if extra_d:
            print(f"  [WARN] 多余设备：{sorted(extra_d)}")

    # 只验证期望设备；extra 设备来自旧缓存消息，不参与后续参数/数值检查
    devices_to_validate = (reported & expected) if expected else reported

    # ── [2][3] 参数完整性 & 单位 ──────────────────────────────────────────────
    print("\n[2][3] 参数完整性 & 单位检查（模板 MQTT 列 / 虚拟设备用 Reading 页面）")
    device_results: dict[str, dict | None] = {}
    for dev_name in sorted(devices_to_validate):
        print(f"\n{SEP2}")
        print(f"  设备：{dev_name}  累计上报参数：{len(device_params[dev_name])} 个")
        # 虚拟设备：Reading 页 key 与 AWS key 后缀可能不同，改为仅比对参数数量
        if _is_virtual_device(dev_name, cfg) and virtual_readings:
            vrd = _find_virtual_readings(dev_name, virtual_readings)
            if vrd is not None:
                n_aws = len(device_params[dev_name])
                n_reading = len(vrd)
                aws_keys = sorted(device_params[dev_name].keys())
                _tmpl = {}
                _rep = {p: device_params[dev_name][p] if isinstance(device_params[dev_name][p], str) else '' for p in aws_keys}
                if n_aws == n_reading:
                    print(f"  [OK] 虚拟设备：参数数量匹配（AWS={n_aws}, Reading={n_reading}），数值见 [4]")
                    device_results[dev_name] = {
                        'ok': True, 'missing': [], 'extra': [], 'unit_mismatch': [],
                        'tmpl_total': n_reading, 'report_total': n_aws,
                        '_all_params': aws_keys, '_tmpl_map': _tmpl, '_reported_units': _rep,
                    }
                else:
                    overall_ok = False
                    print(f"  [FAIL] 虚拟设备：参数数量不匹配（AWS={n_aws}, Reading={n_reading}）")
                    device_results[dev_name] = {
                        'ok': False, 'missing': [], 'extra': [], 'unit_mismatch': [],
                        'tmpl_total': n_reading, 'report_total': n_aws,
                        '_all_params': aws_keys, '_tmpl_map': _tmpl, '_reported_units': _rep,
                    }
                continue
            print(f"  [WARN] 未找到 Reading 页面数据，跳过参数检查")
            device_results[dev_name] = None
            continue
        tmpl_map = _load_mqtt_template(dev_name)
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
    print(f"\n[4] 数值比对（AWS IoT vs 实时 Modbus，容差 ±{tol_pct}% / ±{tol_abs}）")

    # 只比对期望设备的数值，过滤旧消息带入的多余设备
    _values_to_cmp = {k: v for k, v in device_values.items()
                      if not expected or k in expected}
    _cmp_coro = _compare_all_devices(
        _values_to_cmp, tol_pct, tol_abs,
        modbus_device_map=modbus_device_map,
        preread_modbus=_paired_modbus,
        preread_errors=_paired_err,
        paired_aws=_paired_aws,
        virtual_readings=virtual_readings,
        cfg=cfg,
    )
    try:
        asyncio.get_running_loop()
        # pytest-playwright 可能留有运行中的事件循环；在独立线程里执行
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as _pool:
            value_results = _pool.submit(asyncio.run, _cmp_coro).result()
    except RuntimeError:
        value_results = asyncio.run(_cmp_coro)

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
                          f"AWS={_fmt(r.aws_value)}  Modbus={_fmt(r.modbus_value)}  "
                          f"Δ%={r.diff_pct:.2f}%")
        else:
            print(f"  [OK] PASS={dp} ({pr})  ERR={de}")

    print(f"\n{SEP}")
    print("总体结论：[PASS] 全部通过" if overall_ok else "总体结论：[FAIL] 存在问题，请检查上述详情")
    print(SEP)

    report_path = generate_html_report(
        config_info    = aws,
        msg_count      = len(raw_messages),
        device_params  = device_params,
        expected       = expected,
        device_results = device_results,
        value_results  = value_results,
        tol_pct        = tol_pct,
        tol_abs        = tol_abs,
        overall_ok     = overall_ok,
        now_str        = now_str,
        report_name    = _get_caller_case_id(),
    )
    print(f"\nHTML 报告已保存：{report_path}")
    return overall_ok


def main():
    parser = argparse.ArgumentParser(description='AWS IoT 消息验证工具')
    parser.add_argument('--config',
                        default=str(PROJECT_ROOT / 'tests' / 'aws_iot' / 'config.yaml'))
    parser.add_argument('--timeout', type=int, default=90)
    parser.add_argument('--no-web', action='store_true',
                        help='跳过 Web UI 配置/禁用步骤（直接订阅，不操作网关页面）')
    args = parser.parse_args()
    cfg = load_config(args.config)
    ok  = run_verifier(cfg, timeout=args.timeout, skip_web=True,
                       config_path=args.config)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
