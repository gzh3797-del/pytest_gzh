"""
AWS IoT 消息订阅验证脚本

四阶段验证：
  [1] 设备列表：上报设备是否与页面勾选设备一致
  [2] 参数完整性：每台设备上报的参数是否覆盖模板 MQTT 列全量参数
  [3] 单位一致性：每个参数的上报单位是否与模板 unit 列一致
  [4] 数值比对：AWS IoT 上报值 vs 设备实时 Modbus 读取值

用法：
    python Protocols/IoT/AWS_IoT/aws_iot_verifier.py
    python Protocols/IoT/AWS_IoT/aws_iot_verifier.py --timeout 120
"""
import argparse
import asyncio
import html
import json
import logging
import os
import sys
import time
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import paho.mqtt.client as mqtt
import yaml

PROJECT_ROOT  = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
PROTOCOLS_DIR = os.path.join(PROJECT_ROOT, 'Protocols')
sys.path.insert(0, PROJECT_ROOT)
if PROTOCOLS_DIR not in sys.path:
    sys.path.insert(0, PROTOCOLS_DIR)

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as _ChromeOptions
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from test_case.WEB2_4100.operation.LoginPage import LoginPage
    from test_case.WEB2_4100.operation.AWSIoTPage import AWSIoTPage
    _SELENIUM_OK = True
except ImportError:
    _SELENIUM_OK = False

from Protocols.template_reader import find_template_file, get_mqtt_params
import config as _pcfg
import modbus_reader as _mreader
from modbus_reader import ModbusResult


def _patch_pymodbus_device_id() -> None:
    """modbus_reader.py passes device_id= (maps to Modbus unit ID).
    pymodbus 3.6.x silently drops it via **kwargs and defaults slave=0.
    This patch maps device_id → slave so the correct unit ID is forwarded."""
    try:
        from pymodbus.client.mixin import ModbusClientMixin

        def _make_patched(orig):
            import functools
            @functools.wraps(orig)
            def patched(self, *args, **kwargs):
                if 'device_id' in kwargs and 'slave' not in kwargs:
                    kwargs['slave'] = kwargs.pop('device_id')
                return orig(self, *args, **kwargs)
            return patched

        for _name in ('read_holding_registers', 'read_coils', 'read_discrete_inputs'):
            setattr(ModbusClientMixin, _name, _make_patched(getattr(ModbusClientMixin, _name)))
    except (ImportError, AttributeError):
        pass


_patch_pymodbus_device_id()

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


# ─────────────────────────────────────────────────────────────────────────────
# Web UI 配置 / 禁用
# ─────────────────────────────────────────────────────────────────────────────

def _create_driver():
    options = _ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors=yes')
    options.add_argument('--allow-running-insecure-content')
    options.add_argument('--disable-web-security')
    options.add_argument('--start-maximized')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    return webdriver.Chrome(options=options)


def _web_setup(cfg: dict, config_path: str = None):
    """登录网关，配置并启用 AWS IoT，返回 (driver, page) 供 teardown 复用。

    同时读取页面设备列表，自动同步 config.yaml 的 expected_devices。
    """
    if not _SELENIUM_OK:
        raise RuntimeError("selenium 未安装，无法执行 Web 配置")
    gw  = cfg['gateway']
    aws = cfg['aws_iot']
    cert_file = _abs_path(aws['cert_file'])
    key_file  = _abs_path(aws['key_file'])
    for label, path in [('cert', cert_file), ('key', key_file)]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} 文件不存在：{path}")

    driver = _create_driver()
    try:
        driver.get(gw['url'])
        time.sleep(2)
        LoginPage(driver).login(gw['username'], gw['password'])
        try:
            WebDriverWait(driver, 30).until(EC.url_changes(gw['url']))
        except Exception:
            raise RuntimeError(
                f"登录超时（30s 内页面未跳转），请检查账号密码是否正确。"
                f"当前 URL: {driver.current_url}"
            )
        if 'login' in driver.current_url.lower():
            raise RuntimeError(
                f"登录失败（页面未跳转离开登录页），请检查账号密码。"
                f"当前 URL: {driver.current_url}"
            )
        time.sleep(2)
        logging.info(f"[Web] 登录成功：{driver.current_url}")

        page = AWSIoTPage(driver)
        page.navigate_to_aws_iot()
        result = page.configure(
            client_id=aws['client_id'],
            url=aws['url'],
            topic=aws['topic'],
            cert_file=cert_file,
            key_file=key_file,
            interval=aws.get('interval', '30 seconds'),
        )
        logging.info(f"[Web] Test Connection 结果：{result}")

        # ── 同步页面设备列表 → config.yaml ──────────────────────────────
        page_devices = page.get_checked_device_names()
        if page_devices:
            current  = sorted(aws.get('expected_devices', []))
            updated  = sorted(page_devices)
            # 大小写不敏感比对，避免 Acurev1300 vs AcuRev1300 被误报为新增/移除
            current_lower = {d.lower(): d for d in current}
            updated_lower = {d.lower(): d for d in updated}
            added   = sorted(k for k in updated_lower if k not in current_lower)
            removed = sorted(k for k in current_lower if k not in updated_lower)
            # added/removed 显示页面实际名称
            added   = [updated_lower[k] for k in added]
            removed = [current_lower[k] for k in removed]
            if added or removed:
                if added:
                    logging.info(f"[Web] 检测到新增设备：{added}")
                if removed:
                    logging.info(f"[Web] 检测到移除设备：{removed}")
                cfg['aws_iot']['expected_devices'] = updated
                if config_path and os.path.exists(config_path):
                    with open(config_path, 'w', encoding='utf-8') as f:
                        yaml.dump(cfg, f, allow_unicode=True,
                                  default_flow_style=False, sort_keys=False)
                    logging.info(f"[Web] config.yaml 已更新，期望设备：{updated}")
            else:
                logging.info(f"[Web] 设备列表无变化，共 {len(updated)} 台：{updated}")

        return driver, page
    except Exception:
        driver.quit()
        raise


def _web_teardown(driver, page) -> None:
    """导航回 AWS IoT 页面，点击 Disable 并保存，然后关闭浏览器。"""
    try:
        page.navigate_to_aws_iot()
        page.disable()
        logging.info("[Web] AWS IoT 已禁用")
    except Exception as e:
        logging.warning(f"[Web] 禁用流程出错（已忽略）：{e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

TEMPLATE_DIR = os.path.join(PROJECT_ROOT, 'knowledge', 'shared', 'templates', 'raw')
REPORT_DIR   = os.path.join(PROJECT_ROOT, 'Protocols', 'reports')

# AWS IoT 上报设备名 → Protocols/devices/ 模块名
_DEV_MODULE_MAP: dict[str, str] = {
    "AcuRev4100": "devices.acurev4100",
    "AcuRev2100": "devices.acurev2100",
    "AcuvimIIW":  "devices.acuvimiiw",
    "AcuvimIIR":  "devices.acuvimiir",
    "AcuVIM3":    "devices.acuvim3",
    "Acuvim3":    "devices.acuvim3",   # 网关上报名与标准名大小写不同
    "AcuRev1300": "devices.pxm350",
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


# ─────────────────────────────────────────────────────────────────────────────
# 数据结构
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
# 配置 & 解析工具
# ─────────────────────────────────────────────────────────────────────────────

def load_config(path: str) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def _abs_path(path: str) -> str:
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(PROJECT_ROOT, path))


def _load_mqtt_template(device_name: str) -> dict[str, str] | None:
    try:
        xlsx = find_template_file(TEMPLATE_DIR, device_name)
        params = get_mqtt_params(xlsx)
        if not params:
            return None
        logging.info(f"  [模板] {device_name}: {os.path.basename(xlsx)}，MQTT 参数 {len(params)} 个")
        return {p.param_key: p.unit for p in params}
    except FileNotFoundError:
        return None


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


def _check_params(
    json_params: dict[str, str],
    tmpl_map: dict[str, str],
) -> dict:
    tmpl_keys = set(tmpl_map.keys())
    json_keys  = set(json_params.keys())
    missing = sorted(tmpl_keys - json_keys)
    extra   = sorted(json_keys - tmpl_keys)
    unit_mismatch = []
    for param in sorted(tmpl_keys & json_keys):
        if _norm_unit(tmpl_map[param]) != _norm_unit(json_params[param]):
            unit_mismatch.append((param, tmpl_map[param], json_params[param]))
    return {
        'tmpl_total':     len(tmpl_keys),
        'report_total':   len(json_keys),
        'missing':        missing,
        'extra':          extra,
        'unit_mismatch':  unit_mismatch,
        'ok':             not missing and not unit_mismatch,
        '_tmpl_map':      tmpl_map,
        '_reported_units': json_params,
        '_all_params':    sorted(tmpl_keys | json_keys),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 数值比对
# ─────────────────────────────────────────────────────────────────────────────

def _compare_one_value(
    param_key: str,
    aws_value: Optional[float],
    mr: Optional[ModbusResult],
    tol_pct: float,
    tol_abs: float,
) -> ValueCompareResult:
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
    cr.aws_value    = aws_value
    cr.modbus_value = mr.value
    import math
    if math.isnan(aws_value) and math.isnan(mr.value):
        cr.diff_abs = 0.0; cr.diff_pct = 0.0; cr.status = "PASS"; return cr
    diff = abs(aws_value - mr.value)
    ref  = max(abs(aws_value), abs(mr.value))
    cr.diff_abs = diff
    cr.diff_pct = (diff / ref * 100) if ref > 1e-12 else 0.0
    tol = max(tol_abs, ref * tol_pct / 100)
    cr.status = "PASS" if diff <= tol else "FAIL"
    return cr


async def _read_modbus_for_device(
    dev_name: str,
    param_keys: list[str],
) -> tuple[dict[str, ModbusResult] | None, str]:
    """临时覆盖 _pcfg 全局，读取指定设备 Modbus，返回 (result_map, error_msg)。"""
    module_name = _DEV_MODULE_MAP.get(dev_name)
    modbus_info = _pcfg.MODBUS_DEVICE_MAP.get(dev_name)
    # 大小写兜底：网关上报名与 config.py 键名可能大小写不同
    if module_name is None:
        for k, v in _DEV_MODULE_MAP.items():
            if k.lower() == dev_name.lower():
                module_name = v
                break
    if modbus_info is None:
        for k, v in _pcfg.MODBUS_DEVICE_MAP.items():
            if k.lower() == dev_name.lower():
                modbus_info = v
                break
    if not module_name:
        return None, "无设备模块映射（未在 _DEV_MODULE_MAP 中）"
    if not modbus_info:
        return None, "未在 config.MODBUS_DEVICE_MAP 中配置"
    host, port, unit = modbus_info
    _pcfg.DEVICE_NAME   = dev_name
    _pcfg.DEVICE_MODULE = module_name
    _pcfg.MODBUS_HOST   = host
    _pcfg.MODBUS_PORT   = port
    _pcfg.MODBUS_UNIT   = unit
    _mreader._PARAM_MAP = None   # 清除模块缓存，强制重新加载设备映射
    try:
        async with _mreader.get_reader() as reader:
            results = await reader.read_params(param_keys)
        return {r.param_key: r for r in results}, ""
    except Exception as e:
        return None, str(e)


async def _compare_all_devices(
    device_values: dict[str, dict[str, Optional[float]]],
    tol_pct: float,
    tol_abs: float,
) -> dict[str, list[ValueCompareResult] | dict]:
    results: dict[str, list[ValueCompareResult] | dict] = {}
    for dev_name in sorted(device_values.keys()):
        val_map = device_values[dev_name]
        comparable_keys = [k for k, v in val_map.items() if v is not None]
        if not comparable_keys:
            results[dev_name] = {'skip': '无可比对数值（所有参数缺失 value 字段）'}
            continue
        logging.info(f"  [Modbus] {dev_name}：读取 {len(comparable_keys)} 个参数…")
        modbus_map, err = await _read_modbus_for_device(dev_name, comparable_keys)
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


# ─────────────────────────────────────────────────────────────────────────────
# HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

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
    report_path = os.path.join(REPORT_DIR, f'aws_iot_{ts}.html')
    Path(report_path).write_text(content, encoding='utf-8')
    return report_path


# ─────────────────────────────────────────────────────────────────────────────
# 验证主逻辑
# ─────────────────────────────────────────────────────────────────────────────

def run_verifier(cfg: dict, timeout: int = 90, skip_web: bool = False,
                 config_path: str = None) -> bool:
    _driver = _page = None
    try:
        if not skip_web:
            logging.info("=" * 60)
            logging.info("[Web] 启动浏览器，配置 AWS IoT…")
            _driver, _page = _web_setup(cfg, config_path=config_path)
            logging.info("[Web] 配置完成，等待 10 秒让网关开始推送…")
            time.sleep(10)

    except Exception as e:
        import traceback
        logging.error(f"[Web] 配置失败，终止验证：{type(e).__name__}: {e}")
        logging.error(traceback.format_exc())
        if _driver is not None:
            try:
                _driver.quit()
            except Exception:
                pass
        return False

    try:
        return _run_verifier_core(cfg, timeout)
    finally:
        if _driver is not None:
            logging.info("[Web] 禁用 AWS IoT 并关闭浏览器…")
            _web_teardown(_driver, _page)


def _run_verifier_core(cfg: dict, timeout: int) -> bool:
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

    device_params: dict[str, dict[str, str]]            = {}
    device_values: dict[str, dict[str, Optional[float]]] = {}
    raw_messages: list[dict] = []
    msg_count = [0]
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
    if not missing_d and not extra_d:
        print(f"  [OK] 上报设备与勾选设备完全一致：{sorted(reported)}")
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
    tol_pct = _pcfg.MQTT_TOLERANCE_PERCENT
    tol_abs = _pcfg.MQTT_TOLERANCE_ABSOLUTE
    print(f"\n[4] 数值比对（AWS IoT vs 实时 Modbus，容差 ±{tol_pct}% / ±{tol_abs}）")

    value_results = asyncio.run(_compare_all_devices(device_values, tol_pct, tol_abs))

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
    )
    print(f"\nHTML 报告已保存：{report_path}")
    return overall_ok


def main():
    parser = argparse.ArgumentParser(description='AWS IoT 消息验证工具')
    parser.add_argument('--config',
                        default=os.path.join(os.path.dirname(__file__), 'config.yaml'))
    parser.add_argument('--timeout', type=int, default=90)
    parser.add_argument('--no-web', action='store_true',
                        help='跳过 Web UI 配置/禁用步骤（直接订阅，不操作网关页面）')
    args = parser.parse_args()
    cfg = load_config(args.config)
    ok  = run_verifier(cfg, timeout=args.timeout, skip_web=args.no_web,
                       config_path=args.config)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
