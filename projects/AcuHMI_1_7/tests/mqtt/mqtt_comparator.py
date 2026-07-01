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
  # ── 快照模式（原有，读取本地 JSON 文件） ────────────────────────────────────
  python MQTT/mqtt_comparator.py                          # 自动选最新 JSON，第一个在线模块
  python MQTT/mqtt_comparator.py --device acurev4100      # 指定设备
  python MQTT/mqtt_comparator.py --file <json路径>        # 指定文件
  python MQTT/mqtt_comparator.py --module 0               # 指定模块下标（默认第一个在线模块）
  python MQTT/mqtt_comparator.py --no-meta                # 跳过单位检查
  python MQTT/mqtt_comparator.py --keys FREQ_Hz VLN_a_V   # 只比对指定参数

  # ── 实时采集模式（新增，内嵌 Broker 等待设备连入） ──────────────────────────
  python MQTT/mqtt_comparator.py --live                   # 启用实时采集，参数读自 config.py
  python MQTT/mqtt_comparator.py --live --timeout 90      # 采集等待 90s
  python MQTT/mqtt_comparator.py --live --port 1884       # 自定义端口
  python MQTT/mqtt_comparator.py --live --device acurev4100 --timeout 120

  # ── 多设备模式（--all-modules，一次采集生成全部设备的单一 HTML 报告） ────────
  python MQTT/mqtt_comparator.py --live --all-modules              # 等待所有设备推送，生成合并报告
  python MQTT/mqtt_comparator.py --live --all-modules --timeout 180

  # ── mTLS 模式（--ssl，启用双向 TLS） ─────────────────────────────────────────
  python MQTT/mqtt_comparator.py --live --all-modules --ssl              # 启用 mTLS（IP 自动检测）
  python MQTT/mqtt_comparator.py --live --all-modules --ssl --ssl-host www.accu.com  # 额外加入域名
  python MQTT/mqtt_comparator.py --live --all-modules --ssl --cert-dir C:/my_certs   # 指定证书目录
  python MQTT/mqtt_comparator.py --live --ssl --ca-cert ca.crt --server-cert s.crt --server-key s.key --client-cert c.crt --client-key c.key
"""

from __future__ import annotations

import asyncio
import html
import json
import logging
import re
import sys
import threading
import time
import warnings
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import ssl as _ssl

import paho.mqtt.client as mqtt

sys.path.insert(0, str(Path(__file__).parent.parent.parent))          # projects/AcuHMI_1_7/
sys.path.insert(0, str(Path(__file__).parent.parent.parent.parent.parent / "tools" / "Protocols"))  # tools/Protocols/

import settings as config
from modbus_reader import ModbusReader, ModbusResult, get_reader
from template_reader import TemplateParam, find_template_file, get_mqtt_params, natural_sort_key

log = logging.getLogger(__name__)


def _sync_modbus_params() -> None:
    """将 settings 里的 Modbus 连接参数同步到 tools/Protocols/config，
    使 modbus_reader（import config）能读到正确的 HOST/PORT/UNIT。"""
    import importlib
    _tc = sys.modules.get("config") or importlib.import_module("config")
    if _tc is not config:
        _tc.MODBUS_HOST = config.MODBUS_HOST
        _tc.MODBUS_PORT = config.MODBUS_PORT
        _tc.MODBUS_UNIT = config.MODBUS_UNIT
        _tc.DEVICE_NAME = config.DEVICE_NAME


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
# 实时采集（内嵌 Broker + paho 订阅）
# ─────────────────────────────────────────────────────────────────────────────

warnings.filterwarnings("ignore")
logging.getLogger("amqtt").setLevel(logging.ERROR)
logging.getLogger("transitions").setLevel(logging.ERROR)

_LIVE_VALUE_RE = re.compile(r'^([+-]?\d+\.?\d*(?:[eE][+-]?\d+)?)')


def _parse_live_value(v) -> Optional[float]:
    if isinstance(v, (int, float)):
        return float(v)
    m = _LIVE_VALUE_RE.match(str(v).strip())
    return float(m.group(1)) if m else None


def _parse_live_message(
    topic: str, raw: bytes
) -> list[tuple[str, str, bool, dict[str, float], dict[str, str]]]:
    """解析一条实时 MQTT 消息（AcuRev 格式）。
    返回：[(module_name, module_model, online, value_map, unit_map), ...]
    """
    try:
        payload = json.loads(raw.decode("utf-8"))
    except Exception:
        return []
    modules = payload.get("modules")
    if not modules:
        return []
    results = []
    for mod in modules:
        mod_name  = str(mod.get("name") or topic.split("/")[-1])
        mod_model = str(mod.get("model", ""))
        online    = bool(mod.get("online", False))
        value_map: dict[str, float] = {}
        unit_map:  dict[str, str]   = {}
        for entry in mod.get("reading", []):
            param = str(entry.get("param", "")).strip()
            value = entry.get("value")
            unit  = str(entry.get("unit", "")).strip()
            if not param or value is None:
                continue
            parsed = _parse_live_value(value)
            if parsed is not None:
                value_map[param] = parsed
                unit_map[param]  = unit
            elif str(value).strip().lower() in ("-nan", "nan"):
                # 设备在测量值不可用时输出 -nan，参数本身正常发布，仅计入范围检查
                unit_map[param] = unit
        if value_map:
            results.append((mod_name, mod_model, online, value_map, unit_map))
    return results


_SSL_CERT_KEYS = ("cafile", "certfile", "keyfile", "client_cert", "client_key")


def _ensure_ssl_files(
    ssl_config: dict,
    extra_hosts: Optional[list[str]] = None,
) -> None:
    """
    检查 SSL 证书文件是否存在，缺失时自动生成（无需手动操作）。

    生成策略：
      - 全部缺失（全新环境）   → 生成完整套件（CA + server + client）
      - 仅 server.crt 缺失    → 复用已有共享 CA，只重生成 server.crt
      - 全部存在              → 直接跳过

    IP 地址通过检测本机网卡自动获取；extra_hosts 可追加额外 IP/域名。
    """
    missing = [k for k in _SSL_CERT_KEYS
               if not ssl_config.get(k) or not Path(ssl_config[k]).exists()]
    if not missing:
        log.debug("SSL 证书文件已就绪，跳过生成")
        return

    # ── 导入生成函数 ──────────────────────────────────────────────────────────
    _mqtt_dir = Path(__file__).parent
    sys.path.insert(0, str(_mqtt_dir))
    from gen_certs import (                          # noqa: PLC0415
        generate_certs      as _gen_all,
        generate_server_cert as _gen_server,
        _detect_local_ips,
    )

    # ── 自动检测本机 IP ───────────────────────────────────────────────────────
    auto_ips = _detect_local_ips()
    extra    = [h for h in (extra_hosts or []) if h not in auto_ips]
    all_san  = auto_ips + extra
    gen_out  = Path(ssl_config["cafile"]).parent
    gen_out.mkdir(parents=True, exist_ok=True)

    # ── 判断生成模式 ──────────────────────────────────────────────────────────
    server_only = (
        ("certfile" in missing or "keyfile" in missing)
        and (gen_out / "ca.crt").exists()
        and (gen_out / "ca.key").exists()
        and (gen_out / "client.crt").exists()
        and (gen_out / "client.key").exists()
    )

    if server_only:
        log.info("[SSL] 检测到共享 CA，仅重新生成 server.crt（本机 IP：%s）", auto_ips)
        print(f"[SSL] 检测到共享 CA，重新生成 server.crt（本机 IP：{auto_ips}）…")
        _gen_server(gen_out, extra_hosts=all_san, days=3650)
    else:
        log.info("[SSL] 证书不存在，生成完整套件（本机 IP：%s）", auto_ips)
        print(f"[SSL] 正在生成证书套件（本机 IP：{auto_ips}）…")
        _gen_all(gen_out, extra_hosts=all_san, days=3650)

    # ── 最终校验 ──────────────────────────────────────────────────────────────
    still_missing = [k for k in _SSL_CERT_KEYS
                     if not ssl_config.get(k) or not Path(ssl_config[k]).exists()]
    if still_missing:
        raise RuntimeError(
            f"SSL 证书自动生成失败，文件仍缺失：{still_missing}\n"
            f"  请手动运行：python Protocols/MQTT/gen_certs.py --host <本机IP>"
        )
    log.info("[SSL] 证书已就绪")


async def _start_broker(host: str, port: int, ssl_config: Optional[dict] = None):
    from amqtt.broker import Broker
    listeners: dict = {}

    if ssl_config:
        _ensure_ssl_files(ssl_config)
        # ── SSL 主监听端口（amqtt 要求必须有名为 "default" 的 listener） ──────
        listeners["default"] = {
            "type":     "tcp",
            "bind":     f"{host}:{port}",
            "ssl":      True,
            "cafile":   ssl_config["cafile"],
            "certfile": ssl_config["certfile"],
            "keyfile":  ssl_config["keyfile"],
        }
        # ── 同时开明文端口（plain_port=0 表示禁用） ──────────────────────────
        # 关键：必须显式写 "ssl": False
        # amqtt._build_listeners_config 会把 default 的所有字段继承给子 listener，
        # 若不显式覆盖，plain 端口会继承 ssl:True，导致明文连接失败。
        plain_port = ssl_config.get("plain_port", 0)
        if plain_port:
            listeners["plain"] = {
                "type": "tcp",
                "bind": f"{host}:{plain_port}",
                "ssl":  False,   # 阻止继承 default 的 ssl:True
            }
    else:
        listeners["default"] = {"type": "tcp", "bind": f"{host}:{port}"}

    broker = Broker({
        "listeners": listeners,
        "plugins":   {"amqtt.plugins.authentication.AnonymousAuthPlugin": {}},
    })

    if ssl_config:
        # amqtt 默认将 cafile 传给 ssl.create_default_context，导致 Python SSL 自动把
        # CA 证书追加到发送链中（server.crt + ca.crt）。HMI1-7 嵌入式 SSL 库不支持链式
        # 证书，握手报 bad_certificate。此处改用 PROTOCOL_TLS_SERVER 并仅加载 server.crt，
        # 确保 broker 只发送单张证书。
        _orig_create_ssl = Broker._create_ssl_context

        def _single_cert_ssl(self, listener: dict) -> _ssl.SSLContext:
            ctx = _ssl.SSLContext(_ssl.PROTOCOL_TLS_SERVER)
            ctx.load_cert_chain(listener["certfile"], listener["keyfile"])
            ctx.verify_mode = _ssl.CERT_NONE
            # HMI1-7 嵌入式 SSL 库兼容：开放全部 cipher（含旧 RSA/SHA-1），允许 TLS 1.0+
            ctx.set_ciphers("ALL:@SECLEVEL=0")
            ctx.minimum_version = _ssl.TLSVersion.TLSv1
            return ctx

        Broker._create_ssl_context = _single_cert_ssl

    await broker.start()

    if ssl_config:
        Broker._create_ssl_context = _orig_create_ssl  # 恢复，避免影响其他实例

    # 启动后打印端口摘要，方便调试
    if ssl_config:
        plain_port = ssl_config.get("plain_port", 0)
        if plain_port:
            log.info("Broker 已启动：SSL %s:%d  +  明文 %s:%d（双端口模式）",
                     host, port, host, plain_port)
        else:
            log.info("Broker 已启动：SSL-only %s:%d", host, port)
    else:
        log.info("Broker 已启动：明文 %s:%d", host, port)

    return broker


async def _stop_broker(broker) -> None:
    try:
        await broker.shutdown()
    except Exception:
        pass


def start_embedded_broker(
    port: int = 1883,
    ssl_config: Optional[dict] = None,
    duration: int = 300,
) -> threading.Event:
    """在后台线程中启动内嵌 amqtt Broker，供外部测试复用。

    Returns
    -------
    stop_event : threading.Event
        调用 ``stop_event.set()`` 可提前关闭 Broker（否则 duration 秒后自动停止）。

    Usage::

        stop = start_embedded_broker(port=1883)
        try:
            client.connect("127.0.0.1", 1883)
            ...
        finally:
            stop.set()
    """
    broker_ready = threading.Event()
    stop_event   = threading.Event()

    async def _coro():
        b = await _start_broker("0.0.0.0", port, ssl_config=ssl_config)
        broker_ready.set()
        deadline = time.time() + duration
        while not stop_event.is_set() and time.time() < deadline:
            await asyncio.sleep(0.5)
        await _stop_broker(b)

    def _run_coro() -> None:
        # Windows ProactorEventLoop 与 SSL 存在兼容性问题，强制使用 SelectorEventLoop。
        if sys.platform == "win32":
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        asyncio.run(_coro())

    threading.Thread(target=_run_coro, daemon=True).start()
    if not broker_ready.wait(timeout=15):
        raise RuntimeError(f"内嵌 MQTT Broker 启动超时（port={port}），请检查端口是否被占用")
    log.info("内嵌 Broker 已启动 0.0.0.0:%d，最长运行 %ds", port, duration)
    return stop_event


_live_stop_event = threading.Event()


def stop_live_collection() -> None:
    """从外部主动停止正在运行的实时采集，collect_live_mqtt() 将提前返回已收到的数据。"""
    _live_stop_event.set()


async def _collect_modules_data(
    host: str,
    port: int,
    topic: str,
    timeout: int,
    ssl_config: Optional[dict] = None,
    skip_broker: bool = False,
) -> dict[str, dict]:
    """
    启动内嵌 Broker，订阅并收集所有模块数据。
    skip_broker=True 时跳过启动 broker（外部已启动），直接订阅。
    返回 modules_data: {module_name: {model, online, value_map, unit_map, push_count, last_push_time}}
    """
    if not skip_broker:
        broker_holder:      list = []
        broker_loop_holder: list = []
        broker_ready = threading.Event()

        async def _broker_coro():
            b = await _start_broker(host, port, ssl_config=ssl_config)
            broker_holder.append(b)
            broker_loop_holder.append(asyncio.get_running_loop())
            broker_ready.set()
            await asyncio.sleep(timeout + 30)
            await _stop_broker(b)

        def _run_broker_coro() -> None:
            if sys.platform == "win32":
                asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
            asyncio.run(_broker_coro())

        threading.Thread(target=_run_broker_coro, daemon=True).start()
        if not broker_ready.wait(timeout=10):
            raise RuntimeError("内嵌 MQTT Broker 启动超时，请检查端口是否被占用")
    _plain = ssl_config.get("plain_port", 0) if ssl_config else 0
    _ports_str = (f"{port}(SSL)" + (f" + {_plain}(明文)" if _plain else "")) if ssl_config else str(port)
    log.info("内嵌 Broker 已启动 %s [%s]，等待设备连入 %ds …", host, _ports_str, timeout)

    loop  = asyncio.get_running_loop()
    queue: asyncio.Queue = asyncio.Queue()

    def _on_msg(client, userdata, msg):
        pairs = _parse_live_message(msg.topic, msg.payload)
        if pairs:
            loop.call_soon_threadsafe(queue.put_nowait, pairs)

    sub = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id="mqtt-live-collector")
    sub.on_message = _on_msg
    if ssl_config:
        sub.tls_set(
            ca_certs=ssl_config["cafile"],
            certfile=ssl_config["client_cert"],
            keyfile=ssl_config["client_key"],
            tls_version=_ssl.PROTOCOL_TLS_CLIENT,
        )
        log.info("paho 订阅器：SSL 连接 127.0.0.1:%d（mTLS，含客户端证书）", port)
    sub.connect("127.0.0.1", port, keepalive=60)
    sub.subscribe(topic)
    sub.loop_start()

    modules_data: dict[str, dict] = {}
    deadline = loop.time() + timeout
    while loop.time() < deadline and not _live_stop_event.is_set():
        remaining = deadline - loop.time()
        try:
            pairs = await asyncio.wait_for(queue.get(), timeout=min(remaining, 1.0))
        except asyncio.TimeoutError:
            continue
        now = time.time()
        for mod_name, mod_model, online, value_map, unit_map in pairs:
            if mod_name not in modules_data:
                modules_data[mod_name] = {
                    "model": mod_model, "online": online,
                    "value_map": {}, "unit_map": {},
                    "push_count": 0, "last_push_time": None,
                }
            md = modules_data[mod_name]
            md["value_map"].update(value_map)
            md["unit_map"].update(unit_map)
            md["push_count"] += 1
            interval_str = f"{now - md['last_push_time']:.1f}s" if md["last_push_time"] else "—"
            md["last_push_time"] = now
            print(f"[{datetime.now().strftime('%H:%M:%S')}] [RX] [{mod_name}]  "
                  f"第 {md['push_count']} 次推送  间隔 {interval_str}  累计参数 {len(md['value_map'])}")
            log.info("[RX] [%s] 第 %d 次推送  间隔 %s  累计 %d 个参数",
                     mod_name, md["push_count"], interval_str, len(md["value_map"]))

    sub.loop_stop()
    sub.disconnect()
    # skip_broker=True 时 broker_holder 未定义，只有自己启动的 broker 才需要停止
    if not skip_broker:
        if broker_holder and broker_loop_holder:
            future = asyncio.run_coroutine_threadsafe(
                _stop_broker(broker_holder[0]), broker_loop_holder[0]
            )
            try:
                future.result(timeout=5)
            except Exception:
                pass

    return modules_data


async def collect_live_mqtt(
    host: str = "0.0.0.0",
    port: int = 1883,
    topic: str = "#",
    timeout: int = 60,
    module_index: Optional[int] = None,
    ssl_config: Optional[dict] = None,
    skip_broker: bool = False,
) -> tuple[dict[str, float], dict[str, str], str, str, str]:
    """
    启动内嵌 Broker，等待设备连入并采集实时 MQTT 数据（选取单一模块）。
    返回与 load_mqtt_json() 相同的五元组：
        (value_map, unit_map, timestamp_str, module_name, module_model)
    """
    _live_stop_event.clear()
    modules_data = await _collect_modules_data(host, port, topic, timeout, ssl_config=ssl_config, skip_broker=skip_broker)

    if not modules_data:
        raise ValueError(
            f"实时采集超时（{timeout}s），未收到任何 MQTT 数据，"
            f"请确认设备已连接到本机 {host}:{port}"
        )

    mod_names = list(modules_data.keys())
    if module_index is not None:
        if module_index >= len(mod_names):
            raise IndexError(
                f"module_index={module_index} 超出范围（共 {len(mod_names)} 个模块）"
            )
        chosen = mod_names[module_index]
    else:
        online_mods = [n for n in mod_names if modules_data[n]["online"]]
        chosen = online_mods[0] if online_mods else mod_names[0]
        if not online_mods:
            log.warning("所有模块均 offline，使用第一个模块 %s", chosen)

    md = modules_data[chosen]
    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log.info("已选模块：name=%s  model=%s  参数数=%d",
             chosen, md["model"], len(md["value_map"]))
    return md["value_map"], md["unit_map"], timestamp_str, chosen, md["model"]


# ─────────────────────────────────────────────────────────────────────────────
# 多设备模式：全模块采集 + 各自匹配模板 + 合并 HTML 报告
# ─────────────────────────────────────────────────────────────────────────────

def _find_template_for_module(
    module_name: str,
) -> Optional[tuple[str, list]]:
    """根据模块名自动查找并加载 MQTT 参数模板，返回 (path, params) 或 None。"""
    try:
        tmpl_path   = find_template_file(config.TEMPLATE_DIR, module_name)
        tmpl_params = get_mqtt_params(tmpl_path)
        return tmpl_path, tmpl_params
    except FileNotFoundError:
        log.warning("未找到模块 '%s' 的模板文件", module_name)
        return None
    except Exception as exc:
        log.warning("加载模块 '%s' 模板时出错：%s", module_name, exc)
        return None


async def run_all_modules_no_modbus(
    live_host: str = "0.0.0.0",
    live_port: int = 1883,
    live_timeout: int = 60,
    live_topic: str = "#",
    ssl_config: Optional[dict] = None,
) -> tuple[list[dict], str]:
    """
    采集全部在线模块，对每个模块按名称自动匹配模板，做范围+单位检查（无 Modbus 比对）。

    返回 (module_reports, timestamp_str)
    module_reports 每项：
        {module_name, module_model, tmpl_path, scope: MQTTScopeReport,
         unit_results: list[MQTTUnitResult]}
    """
    _live_stop_event.clear()
    modules_data = await _collect_modules_data(live_host, live_port, live_topic, live_timeout, ssl_config=ssl_config)

    if not modules_data:
        raise ValueError(
            f"实时采集超时（{live_timeout}s），未收到任何 MQTT 数据，"
            f"请确认设备已连接到本机 {live_host}:{live_port}"
        )

    timestamp_str  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    module_reports = []

    for mod_name, md in modules_data.items():
        value_map = md["value_map"]
        unit_map  = md["unit_map"]
        json_keys = set(unit_map.keys())   # 含 -nan 等非数值参数，范围检查以此为准

        tmpl_result = _find_template_for_module(mod_name)
        if tmpl_result:
            tmpl_path, tmpl_params = tmpl_result
            tmpl_map  = {p.param_key: p for p in tmpl_params}
            tmpl_keys = set(tmpl_map)
        else:
            tmpl_path, tmpl_map, tmpl_keys = None, {}, set()

        if tmpl_keys:
            matched_keys_set  = tmpl_keys & json_keys
            missing_from_json = sorted(tmpl_keys - json_keys, key=natural_sort_key)
            extra_in_json     = sorted(json_keys - tmpl_keys, key=natural_sort_key)
        else:
            matched_keys_set  = json_keys
            missing_from_json = []
            extra_in_json     = []

        scope = MQTTScopeReport(
            template_count    = len(tmpl_keys),
            json_count        = len(json_keys),
            matched_keys      = sorted(matched_keys_set, key=natural_sort_key),
            missing_from_json = missing_from_json,
            extra_in_json     = extra_in_json,
        )

        unit_results: list[MQTTUnitResult] = []
        if tmpl_map:
            for pkey in sorted(matched_keys_set, key=natural_sort_key):
                unit_results.append(
                    check_unit(pkey, tmpl_map[pkey].unit, unit_map.get(pkey, ""))
                )

        log.info("模块 %s：范围 %d/%d  单位不匹配 %d 项",
                 mod_name, len(matched_keys_set), len(json_keys),
                 sum(1 for r in unit_results if not r.ok))
        module_reports.append({
            "module_name":  mod_name,
            "module_model": md["model"],
            "tmpl_path":    tmpl_path,
            "scope":        scope,
            "unit_results": unit_results,
        })

    return module_reports, timestamp_str


def generate_multi_device_html_report(
    module_reports: list[dict],
    timestamp_str: str = "",
    live_host: str = "",
    live_port: int = 1883,
    live_timeout: int = 0,
    output_path: Optional[str] = None,
) -> str:
    """生成包含全部设备的参数范围 & 单位比对 HTML 报告，返回文件路径。"""
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"mqtt_all_devices_{ts}.html")

    now_str     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_disp = f"实时采集（{live_host}:{live_port}，{live_timeout}s）"
    total       = len(module_reports)
    devices_ok  = sum(
        1 for r in module_reports
        if r["scope"].scope_ok and not any(not u.ok for u in r["unit_results"])
    )

    css = """
  body { font-family: "Microsoft YaHei", Arial, sans-serif; font-size: 13px;
          margin: 20px; background: #f5f5f5; color: #333; }
  h1   { font-size: 20px; margin-bottom: 4px; }
  h3   { font-size: 13px; margin: 8px 0 4px; }
  .meta { color: #666; font-size: 12px; margin-bottom: 16px; }
  .cards { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 6px; padding: 14px 20px;
           box-shadow: 0 1px 4px rgba(0,0,0,.1); min-width: 110px; text-align: center; }
  .card .num { font-size: 28px; font-weight: bold; }
  .card .lbl { font-size: 11px; color: #888; margin-top: 2px; }
  .card.pass  .num { color: #28a745; }
  .card.fail  .num { color: #dc3545; }
  .card.total .num { color: #007bff; }
  table { border-collapse: collapse; width: 100%; background: #fff;
           box-shadow: 0 1px 4px rgba(0,0,0,.1); border-radius: 6px;
           overflow: hidden; table-layout: fixed; margin-bottom: 12px; }
  thead tr { background: #343a40; color: #fff; }
  th   { padding: 9px 8px; text-align: center; font-size: 12px; white-space: nowrap; }
  td   { padding: 5px 8px; border-bottom: 1px solid #eee; vertical-align: middle; }
  td.num  { text-align: right; color: #999; font-size: 11px; }
  td.key  { font-family: monospace; font-size: 12px; word-break: break-all; }
  td.val  { text-align: right; font-family: monospace; font-size: 12px; }
  td.stat { text-align: center; font-weight: bold; font-size: 12px; }
  tr:hover td { filter: brightness(0.96); }
  thead th { position: sticky; top: 0; z-index: 1; }
  details.section { background: #fff; border-radius: 8px; margin-bottom: 20px;
                    box-shadow: 0 1px 4px rgba(0,0,0,.1); overflow: hidden; }
  details.section > summary { list-style: none; cursor: pointer; padding: 12px 16px;
    background: #f0f4f8; border-left: 4px solid #0056b3;
    font-size: 15px; font-weight: bold; color: #0056b3;
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap; user-select: none; }
  details.section > summary::-webkit-details-marker { display: none; }
  details.section > summary::before { content: "▶"; font-size: 10px; margin-right: 4px; }
  details[open].section > summary::before { content: "▼"; }
  .sum-info { font-size: 12px; font-weight: normal; color: #666; margin-left: 4px; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 10px;
           font-size: 11px; font-weight: bold; margin-left: 4px; }
  .ok-badge   { background: #d4edda; color: #155724; }
  .err-badge  { background: #f8d7da; color: #721c24; }
  .warn-badge { background: #fff3cd; color: #856404; }
  .section-body { padding: 16px; }"""

    def _list_rows(keys: list[str], bg: str) -> str:
        return "".join(
            f'<tr style="background:{bg}"><td class="num">{i+1}</td>'
            f'<td class="key">{html.escape(k)}</td></tr>'
            for i, k in enumerate(keys)
        )

    device_sections: list[str] = []
    for idx, rpt in enumerate(module_reports):
        scope        = rpt["scope"]
        unit_results = rpt["unit_results"]
        mod_name     = rpt["module_name"]
        mod_model    = rpt["module_model"]
        tmpl_path    = rpt["tmpl_path"]
        unit_fail    = sum(1 for r in unit_results if not r.ok)

        scope_color  = "#d4edda" if scope.scope_ok else "#f8d7da"
        scope_label  = "一致" if scope.scope_ok else "不一致"
        tmpl_display = Path(tmpl_path).name if tmpl_path else "（未找到模板）"

        badges = ""
        if not tmpl_path:
            badges += '<span class="badge warn-badge">无模板</span>'
        if scope.missing_from_json:
            badges += f'<span class="badge err-badge">缺失 {len(scope.missing_from_json)}</span>'
        if scope.extra_in_json:
            badges += f'<span class="badge warn-badge">多余 {len(scope.extra_in_json)}</span>'
        if scope.scope_ok and tmpl_path:
            badges += '<span class="badge ok-badge">范围一致</span>'
        if unit_fail:
            badges += f'<span class="badge err-badge">单位不匹配 {unit_fail}</span>'
        elif unit_results:
            badges += '<span class="badge ok-badge">单位一致</span>'

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

        if unit_results:
            unit_rows = "".join(
                f'<tr style="background:{"#f8d7da" if not r.ok else "#d4edda"}">'
                f'<td class="num">{j+1}</td>'
                f'<td class="key">{html.escape(r.param_key)}</td>'
                f'<td class="val">{html.escape(r.tmpl_unit)}</td>'
                f'<td class="val">{html.escape(r.json_unit)}</td>'
                f'<td class="stat">{"不匹配" if not r.ok else "一致"}</td>'
                f'</tr>'
                for j, r in enumerate(unit_results)
            )
            unit_badge = (f'<span class="badge err-badge">不匹配 {unit_fail}</span>'
                          if unit_fail else '<span class="badge ok-badge">全部一致</span>')
            unit_block = f"""
<details {"open" if unit_fail else ""} style="margin-top:12px">
<summary style="cursor:pointer;padding:6px 0;font-size:13px;font-weight:bold;color:#555">
  单位检查（共 {len(unit_results)} 项）{unit_badge}
</summary>
<table style="margin-top:8px">
<colgroup>
  <col style="width:40px"><col style="width:280px">
  <col style="width:100px"><col style="width:100px"><col style="width:80px">
</colgroup>
<thead><tr><th>#</th><th>param_key</th><th>模板 unit</th><th>JSON unit</th><th>结果</th></tr></thead>
<tbody>{unit_rows}</tbody>
</table>
</details>"""
        else:
            unit_block = '<p style="color:#888;font-size:12px;margin-top:8px">未找到模板，跳过单位检查。</p>'

        has_issue = scope.missing_from_json or scope.extra_in_json or unit_fail or not tmpl_path
        device_sections.append(f"""
<details {"open" if has_issue else ""} class="section">
<summary>
  {idx+1}. {html.escape(mod_name)}&nbsp;<span style="font-weight:normal;color:#666;font-size:13px">({html.escape(mod_model)})</span>
  <span class="sum-info">模板 {scope.template_count} / JSON {scope.json_count} / 匹配 {len(scope.matched_keys)}</span>
  {badges}
</summary>
<div class="section-body">
<p style="font-size:11px;color:#888;margin:0 0 10px">模板：{html.escape(tmpl_display)}</p>
<div class="cards">
  <div class="card total"><div class="num">{scope.template_count}</div><div class="lbl">模板参数</div></div>
  <div class="card total"><div class="num">{scope.json_count}</div><div class="lbl">JSON 发布</div></div>
  <div class="card pass"> <div class="num">{len(scope.matched_keys)}</div><div class="lbl">匹配</div></div>
  <div class="card fail"> <div class="num">{len(scope.missing_from_json)}</div><div class="lbl">模板有/JSON缺</div></div>
  <div class="card fail"> <div class="num">{len(scope.extra_in_json)}</div><div class="lbl">JSON多/模板无</div></div>
  <div class="card" style="background:{scope_color}"><div class="num" style="font-size:16px">{scope_label}</div><div class="lbl">范围结论</div></div>
</div>
{missing_html}{extra_html}{unit_block}
</div>
</details>""")

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>MQTT 多设备参数/单位比对报告</title>
<style>{css}
</style>
</head>
<body>
<h1>MQTT 多设备参数范围 &amp; 单位比对报告</h1>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  数据来源：{html.escape(source_disp)} &nbsp;|&nbsp;
  采集时间：{html.escape(timestamp_str)} &nbsp;|&nbsp;
  设备数：{total}
</div>
<div class="cards" style="margin-bottom:24px">
  <div class="card total"><div class="num">{total}</div><div class="lbl">设备总数</div></div>
  <div class="card pass"> <div class="num">{devices_ok}</div><div class="lbl">全部正常</div></div>
  <div class="card fail"> <div class="num">{total - devices_ok}</div><div class="lbl">有问题</div></div>
</div>
{"".join(device_sections)}
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("多设备 HTML 报告已保存：%s", output_path)
    return output_path


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
    json_path: Optional[str] = None,
    module_index: Optional[int] = None,
    param_keys: Optional[list[str]] = None,
    check_meta: bool = True,
    live: bool = False,
    live_host: str = "0.0.0.0",
    live_port: int = 1883,
    live_timeout: int = 60,
    live_topic: str = "#",
    no_modbus: bool = False,
    ssl_config: Optional[dict] = None,
    skip_broker: bool = False,
) -> tuple[MQTTScopeReport, list[MQTTUnitResult], list[MQTTCompareResult], str, str, str, str]:
    """
    三段式比对：范围检查 → 单位检查 → 数值比对。

    live=False（默认）：从本地 JSON 文件加载 MQTT 快照。
    live=True         ：启动内嵌 Broker，等待设备实时连入并采集数据。
    no_modbus=True    ：跳过 Modbus 读取和数值比对，仅输出范围/单位检查结果。

    Returns:
        (scope_report, unit_results, compare_results,
         timestamp_str, module_name, module_model, source_label)
    """
    t0 = time.time()

    # ── 确保 Modbus 连接参数与设备映射一致（从测试代码直接调用时 __main__ 不执行）───
    if config.DEVICE_NAME in config.MODBUS_DEVICE_MAP:
        config.MODBUS_HOST, config.MODBUS_PORT, config.MODBUS_UNIT = \
            config.MODBUS_DEVICE_MAP[config.DEVICE_NAME]
    _sync_modbus_params()

    # ── 加载数据（快照 or 实时采集） ───────────────────────────────────────────
    if live:
        log.info("实时采集模式：Broker %s:%d，超时 %ds", live_host, live_port, live_timeout)
        value_map, unit_map, timestamp_str, module_name, module_model = \
            await collect_live_mqtt(
                host=live_host, port=live_port, topic=live_topic,
                timeout=live_timeout, module_index=module_index,
                ssl_config=ssl_config,
                skip_broker=skip_broker,
            )
        source_label = f"实时采集（{live_host}:{live_port}，{live_timeout}s）"
    else:
        log.info("加载 JSON：%s", json_path)
        value_map, unit_map, timestamp_str, module_name, module_model = load_mqtt_json(
            json_path, module_index
        )
        source_label = str(json_path or "")

    # ── 加载模板（按 MQTT 列过滤） ────────────────────────────────────────────
    try:
        tmpl_path   = find_template_file(config.TEMPLATE_DIR, config.DEVICE_NAME)
        tmpl_params = get_mqtt_params(tmpl_path)         # 仅 MQTT 列非空的参数
        tmpl_map    = {p.param_key: p for p in tmpl_params}
        tmpl_keys   = set(tmpl_map)
        log.info("已加载模板：%s（MQTT 范围 %d 个参数）", tmpl_path, len(tmpl_keys))
    except Exception as exc:
        log.warning("无法加载模板文件，范围/单位检查将跳过：%s", exc)
        tmpl_map, tmpl_keys = {}, set()

    json_keys = set(value_map.keys())

    # ── 范围检查 ───────────────────────────────────────────────────────────────
    if tmpl_keys:
        matched_keys_set   = tmpl_keys & json_keys
        missing_from_json  = sorted(tmpl_keys - json_keys, key=natural_sort_key)
        extra_in_json      = sorted(json_keys - tmpl_keys, key=natural_sort_key)
    else:
        matched_keys_set  = json_keys
        missing_from_json = []
        extra_in_json     = []

    scope_report = MQTTScopeReport(
        template_count   = len(tmpl_keys),
        json_count       = len(json_keys),
        matched_keys     = sorted(matched_keys_set, key=natural_sort_key),
        missing_from_json = missing_from_json,
        extra_in_json    = extra_in_json,
    )
    log.info("范围检查：模板=%d  JSON=%d  匹配=%d  缺失=%d  多余=%d",
             len(tmpl_keys), len(json_keys), len(matched_keys_set),
             len(missing_from_json), len(extra_in_json))

    # ── 单位检查 ───────────────────────────────────────────────────────────────
    unit_results: list[MQTTUnitResult] = []
    if check_meta and tmpl_map:
        for pkey in sorted(matched_keys_set, key=natural_sort_key):
            tmpl_unit = tmpl_map[pkey].unit
            json_unit = unit_map.get(pkey, "")
            unit_results.append(check_unit(pkey, tmpl_unit, json_unit))
        unit_fail = sum(1 for r in unit_results if not r.ok)
        log.info("单位检查：共 %d 项，不匹配 %d 项", len(unit_results), unit_fail)

    # ── 数值比对（可跳过） ────────────────────────────────────────────────────
    compare_results: list[MQTTCompareResult] = []
    if no_modbus:
        log.info("已跳过 Modbus 读取和数值比对（--no-modbus）")
    else:
        compare_scope = matched_keys_set if tmpl_keys else json_keys
        comparable: dict[str, float] = {
            k: v for k, v in value_map.items()
            if k in compare_scope and (param_keys is None or k in param_keys)
        }
        if not comparable:
            raise ValueError("无可比对参数（JSON 与模板无交集，或指定的 --keys 均不存在）")
        log.info("数值比对参数数量：%d", len(comparable))
        log.info("读取实时 Modbus 寄存器…")
        async with get_reader() as modbus:
            modbus_results = await modbus.read_params(list(comparable.keys()))
        modbus_map: dict[str, ModbusResult] = {r.param_key: r for r in modbus_results}
        for pkey in sorted(comparable, key=natural_sort_key):
            mv = comparable[pkey]
            mr = modbus_map.get(pkey)
            compare_results.append(_compare_one(pkey, mv, mr))

    elapsed = time.time() - t0
    log.info("比对完成，耗时 %.1f 秒，共 %d 项", elapsed, len(compare_results))
    return scope_report, unit_results, compare_results, timestamp_str, module_name, module_model, source_label


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
    if v is None:
        return "—"
    s = f"{v:.{digits}f}"
    return s.rstrip('0').rstrip('.') if '.' in s else s


def _build_comparison_sections_html(
    scope: MQTTScopeReport,
    unit_results: list[MQTTUnitResult],
    results: list[MQTTCompareResult],
) -> tuple[str, str, str, dict]:
    """构建范围/单位/数值三段式 HTML 片段，供报告函数复用。
    返回 (scope_html, unit_html, val_html, summary_dict)。
    """
    s = summary(results)
    unit_fail = sum(1 for r in unit_results if not r.ok)

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

    return scope_html, unit_html, val_html, s


def generate_html_report(
    scope: MQTTScopeReport,
    unit_results: list[MQTTUnitResult],
    results: list[MQTTCompareResult],
    timestamp_str: str = "",
    module_name: str = "",
    module_model: str = "",
    json_path: str = "",
    output_path: Optional[str] = None,
    source_label: str = "",
) -> str:
    """生成三段式 HTML 比对报告（范围检查 / 单位检查 / 数值比对），返回文件路径。"""
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"mqtt_{config.DEVICE_NAME}_{ts}.html")

    s = summary(results)
    now_str      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_disp  = source_label or (Path(json_path).name if json_path else "—")
    is_live      = source_label.startswith("实时采集")
    unit_fail    = sum(1 for r in unit_results if not r.ok)

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

    report_title = "MQTT 实时采集 vs 实时 Modbus 比对报告" if is_live else "MQTT 快照 vs 实时 Modbus 比对报告"
    data_label   = "数据来源" if is_live else "数据文件"
    time_label   = "采集时间" if is_live else "快照时间戳"

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>MQTT vs Modbus 比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>{css}
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
<h1>{report_title}</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  {data_label}：{html.escape(source_disp)} &nbsp;|&nbsp;
  {time_label}：{html.escape(timestamp_str)} &nbsp;|&nbsp;
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
# 双模式报告（明文 + mTLS 同一 HTML，Tab 切换）
# ─────────────────────────────────────────────────────────────────────────────

def generate_dual_html_report(
    plain_results: tuple,
    ssl_results: tuple,
    plain_port: int = 1883,
    ssl_port: int = 8883,
    output_path: Optional[str] = None,
) -> str:
    """生成明文 + mTLS 双模式比对报告（Tab 切换），返回文件路径。

    plain_results / ssl_results 均为 run_mqtt_comparison() 返回的 7-元组：
        (scope, unit_results, compare_results,
         timestamp_str, module_name, module_model, source_label)
    """
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_dir = Path(config.REPORT_DIR)
        report_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(report_dir / f"mqtt_{config.DEVICE_NAME}_dual_{ts}.html")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _unpack(t: tuple):
        scope, unit_res, cmp_res, ts_str, mod_name, mod_model, src_label = t
        return scope, unit_res, cmp_res, ts_str, mod_name, mod_model, src_label

    p_scope, p_unit, p_cmp, p_ts, p_mod, p_model, p_src = _unpack(plain_results)
    s_scope, s_unit, s_cmp, s_ts, s_mod, s_model, s_src = _unpack(ssl_results)

    p_scope_html, p_unit_html, p_val_html, p_s = _build_comparison_sections_html(p_scope, p_unit, p_cmp)
    s_scope_html, s_unit_html, s_val_html, s_s = _build_comparison_sections_html(s_scope, s_unit, s_cmp)

    def _overall_badge(s_dict: dict, scope: MQTTScopeReport) -> str:
        ok = not s_dict["fail"] and not s_dict["error"] and scope.scope_ok
        color = "#28a745" if ok else "#dc3545"
        label = "✅ 全部通过" if ok else f"❌ FAIL={s_dict['fail']} ERR={s_dict['error']}"
        return f'<span style="color:{color};font-weight:bold">{label}</span>'

    css_base = """
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
  .section-body {{ padding: 16px; }}
  /* ── 双模式摘要卡片 ── */
  .dual-summary {{ display:flex; gap:16px; margin-bottom:20px; }}
  .dual-card {{ flex:1; background:#fff; border-radius:8px; padding:16px 20px;
               box-shadow:0 1px 4px rgba(0,0,0,.12); cursor:pointer;
               border-top:4px solid #ccc; transition:box-shadow .15s; }}
  .dual-card:hover {{ box-shadow:0 3px 10px rgba(0,0,0,.18); }}
  .dual-card.plain-card {{ border-color:#17a2b8; }}
  .dual-card.ssl-card   {{ border-color:#6f42c1; }}
  .dual-card.active {{ box-shadow:0 0 0 3px #0056b3; }}
  .dual-card .dc-title {{ font-size:15px; font-weight:bold; margin-bottom:6px; }}
  .dual-card .dc-stats {{ font-size:12px; color:#555; line-height:1.7; }}
  /* ── Tab 栏 ── */
  .tab-bar {{ display:flex; gap:0; margin-bottom:0; border-bottom:2px solid #dee2e6; }}
  .tab-btn {{ padding:9px 28px; border:none; background:transparent; font-size:14px;
             font-weight:bold; color:#666; cursor:pointer; border-bottom:3px solid transparent;
             margin-bottom:-2px; transition:color .1s; }}
  .tab-btn.active {{ color:#0056b3; border-bottom-color:#0056b3; background:#fff; }}
  .tab-btn:hover:not(.active) {{ color:#333; background:#f0f4f8; }}
  /* ── Tab 内容 ── */
  .tab-pane {{ display:none; padding-top:20px; }}
  .tab-pane.active {{ display:block; }}"""

    # ── 双模式摘要卡片 ──────────────────────────────────────────────────────────
    p_badge = _overall_badge(p_s, p_scope)
    s_badge = _overall_badge(s_s, s_scope)

    dual_summary_html = f"""
<div class="dual-summary">
  <div class="dual-card plain-card active" id="card-plain" onclick="switchTab('plain')">
    <div class="dc-title">🔓 明文 Plain（:{plain_port}）</div>
    <div class="dc-stats">
      {p_badge}<br>
      数值比对：PASS={p_s['pass']} / FAIL={p_s['fail']} / ERR={p_s['error']}
      &nbsp;({p_s['pass_rate']})<br>
      范围：模板 {p_scope.template_count} / JSON {p_scope.json_count}
      &nbsp;缺失={len(p_scope.missing_from_json)}<br>
      采集时间：{html.escape(p_ts)}
    </div>
  </div>
  <div class="dual-card ssl-card" id="card-ssl" onclick="switchTab('ssl')">
    <div class="dc-title">🔒 mTLS SSL（:{ssl_port}）</div>
    <div class="dc-stats">
      {s_badge}<br>
      数值比对：PASS={s_s['pass']} / FAIL={s_s['fail']} / ERR={s_s['error']}
      &nbsp;({s_s['pass_rate']})<br>
      范围：模板 {s_scope.template_count} / JSON {s_scope.json_count}
      &nbsp;缺失={len(s_scope.missing_from_json)}<br>
      采集时间：{html.escape(s_ts)}
    </div>
  </div>
</div>"""

    tab_bar_html = f"""
<div class="tab-bar">
  <button class="tab-btn active" id="tab-plain" onclick="switchTab('plain')">
    🔓 明文 Plain（:{plain_port}）
  </button>
  <button class="tab-btn" id="tab-ssl" onclick="switchTab('ssl')">
    🔒 mTLS SSL（:{ssl_port}）
  </button>
</div>"""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>MQTT 双模式比对报告 — {html.escape(config.DEVICE_NAME)}</title>
<style>{css_base}
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
function switchTab(mode) {{
  ['plain','ssl'].forEach(function(m) {{
    document.getElementById('tab-' + m).classList.toggle('active', m === mode);
    document.getElementById('pane-' + m).classList.toggle('active', m === mode);
    document.getElementById('card-' + m).classList.toggle('active', m === mode);
  }});
}}
function toggleErrOnly() {{
  var btn = document.getElementById('err-toggle');
  var on = btn.dataset.active === '1';
  var pane = document.querySelector('.tab-pane.active');
  if (on) {{
    pane.querySelectorAll('details.section').forEach(function(d) {{ d.style.display = ''; }});
    pane.querySelectorAll('tbody tr').forEach(function(tr) {{ tr.style.display = ''; }});
  }} else {{
    pane.querySelectorAll('details.section').forEach(function(d) {{ d.open = true; }});
    pane.querySelectorAll('tbody tr').forEach(function(tr) {{
      var _s = (tr.getAttribute('style')||'').toLowerCase();
      tr.style.display = _s.indexOf('d4edda') !== -1 ? 'none' : '';
    }});
    pane.querySelectorAll('details.section').forEach(function(d) {{
      var hasErr = Array.from(d.querySelectorAll('tbody tr')).some(function(tr) {{
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
<h1>MQTT 双模式比对报告（明文 + mTLS）</h1>
<div class="device-name">设备：{html.escape(config.DEVICE_NAME)}</div>
<div class="meta">
  生成时间：{now_str} &nbsp;|&nbsp;
  模块：{html.escape(p_mod)} ({html.escape(p_model)}) &nbsp;|&nbsp;
  Modbus：{config.MODBUS_HOST}:{config.MODBUS_PORT} Unit={config.MODBUS_UNIT} &nbsp;|&nbsp;
  容差：±{config.MQTT_TOLERANCE_PERCENT}% / ±{config.MQTT_TOLERANCE_ABSOLUTE}
</div>
{dual_summary_html}
{tab_bar_html}
<div id="pane-plain" class="tab-pane active">
{p_scope_html}
{p_unit_html}
{p_val_html}
</div>
<div id="pane-ssl" class="tab-pane">
{s_scope_html}
{s_unit_html}
{s_val_html}
</div>
</body>
</html>"""

    Path(output_path).write_text(html_content, encoding="utf-8")
    log.info("双模式 HTML 报告已保存：%s", output_path)
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
    json_path: Optional[str],
    module_index: Optional[int],
    param_keys: Optional[list[str]],
    check_meta: bool,
    live: bool = False,
    live_host: str = "0.0.0.0",
    live_port: int = 1883,
    live_timeout: int = 60,
    live_topic: str = "#",
    no_modbus: bool = False,
    ssl_config: Optional[dict] = None,
) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )
    scope, unit_results, results, timestamp_str, module_name, module_model, source_label = \
        await run_mqtt_comparison(
            json_path, module_index, param_keys, check_meta,
            live=live, live_host=live_host, live_port=live_port,
            live_timeout=live_timeout, live_topic=live_topic,
            no_modbus=no_modbus, ssl_config=ssl_config,
        )
    print_summary(scope, unit_results, results, module_name, module_model)
    report_path = generate_html_report(
        scope, unit_results, results,
        timestamp_str=timestamp_str,
        module_name=module_name,
        module_model=module_model,
        json_path=json_path or "",
        source_label=source_label,
    )
    print(f"\n  HTML 报告：{report_path}\n")


async def _main_all_modules(
    live_host: str,
    live_port: int,
    live_timeout: int,
    live_topic: str,
    ssl_config: Optional[dict] = None,
) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )
    module_reports, timestamp_str = await run_all_modules_no_modbus(
        live_host=live_host, live_port=live_port,
        live_timeout=live_timeout, live_topic=live_topic,
        ssl_config=ssl_config,
    )
    report_path = generate_multi_device_html_report(
        module_reports, timestamp_str=timestamp_str,
        live_host=live_host, live_port=live_port, live_timeout=live_timeout,
    )

    print("\n" + "=" * 70)
    print("  MQTT 多设备参数范围 & 单位比对摘要")
    print(f"  时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  设备数: {len(module_reports)}")
    print("=" * 70)
    for rpt in module_reports:
        scope     = rpt["scope"]
        unit_fail = sum(1 for r in rpt["unit_results"] if not r.ok)
        status    = "OK  " if scope.scope_ok and unit_fail == 0 else "WARN"
        print(f"  [{status}] {rpt['module_name']:20s}  "
              f"模板={scope.template_count:4d}  JSON={scope.json_count:4d}  "
              f"匹配={len(scope.matched_keys):4d}  "
              f"缺失={len(scope.missing_from_json):4d}  多余={len(scope.extra_in_json):4d}  "
              f"单位不匹配={unit_fail}")
    print("=" * 70)
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
    _sync_modbus_params()

    # ── --live（实时采集模式） ─────────────────────────────────────────────────
    _live = "--live" in sys.argv
    _live_host    = getattr(config, "MQTT_BROKER_HOST",     "0.0.0.0")
    _live_port    = getattr(config, "MQTT_BROKER_PORT",     1883)
    _live_timeout = getattr(config, "MQTT_COLLECT_TIMEOUT", 60)
    _live_topic   = getattr(config, "MQTT_TOPIC",           "#")

    if "--host" in sys.argv:
        idx = sys.argv.index("--host")
        _live_host = sys.argv[idx + 1]
    if "--port" in sys.argv:
        idx = sys.argv.index("--port")
        _live_port = int(sys.argv[idx + 1])
    if "--timeout" in sys.argv:
        idx = sys.argv.index("--timeout")
        _live_timeout = int(sys.argv[idx + 1])

    # ── --file（快照模式，与 --live 互斥） ────────────────────────────────────
    _json_path: Optional[str] = None
    if not _live:
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

    # ── --no-modbus ───────────────────────────────────────────────────────────
    _no_modbus = "--no-modbus" in sys.argv

    # ── --keys ────────────────────────────────────────────────────────────────
    _keys: Optional[list[str]] = None
    if "--keys" in sys.argv:
        idx = sys.argv.index("--keys")
        _keys = sys.argv[idx + 1:]

    # ── --ssl（mTLS 模式） ────────────────────────────────────────────────────
    _ssl_enabled = "--ssl" in sys.argv
    _ssl_config: Optional[dict] = None
    if _ssl_enabled:
        # 默认从 config.py 读取路径
        _ssl_config = {
            "cafile":      getattr(config, "MQTT_SSL_CA_CERT",     ""),
            "certfile":    getattr(config, "MQTT_SSL_SERVER_CERT", ""),
            "keyfile":     getattr(config, "MQTT_SSL_SERVER_KEY",  ""),
            "client_cert": getattr(config, "MQTT_SSL_CLIENT_CERT", ""),
            "client_key":  getattr(config, "MQTT_SSL_CLIENT_KEY",  ""),
        }
        # --cert-dir 覆盖整个证书目录（使用标准文件名）
        if "--cert-dir" in sys.argv:
            idx = sys.argv.index("--cert-dir")
            _cert_dir = Path(sys.argv[idx + 1])
            _ssl_config.update({
                "cafile":      str(_cert_dir / "ca.crt"),
                "certfile":    str(_cert_dir / "server.crt"),
                "keyfile":     str(_cert_dir / "server.key"),
                "client_cert": str(_cert_dir / "client.crt"),
                "client_key":  str(_cert_dir / "client.key"),
            })
        # 单个证书路径覆盖（粒度最细，优先级最高）
        for _cli_arg, _cfg_key in (
            ("--ca-cert",     "cafile"),
            ("--server-cert", "certfile"),
            ("--server-key",  "keyfile"),
            ("--client-cert", "client_cert"),
            ("--client-key",  "client_key"),
        ):
            if _cli_arg in sys.argv:
                _idx = sys.argv.index(_cli_arg)
                _ssl_config[_cfg_key] = sys.argv[_idx + 1]
        # --ssl-host：额外加入证书 SAN 的域名或 IP（可多个）
        _ssl_extra_hosts: list[str] = []
        if "--ssl-host" in sys.argv:
            idx = sys.argv.index("--ssl-host")
            # 收集 --ssl-host 后面直到下一个 -- 参数为止的所有值
            for _h in sys.argv[idx + 1:]:
                if _h.startswith("-"):
                    break
                _ssl_extra_hosts.append(_h)

        # 证书按需自动生成（检测本机 IP，支持共享 CA 模式）
        _ensure_ssl_files(_ssl_config, extra_hosts=_ssl_extra_hosts)
        # --ssl 默认端口改为 8883
        if "--port" not in sys.argv:
            _live_port = getattr(config, "MQTT_SSL_PORT", 8883)
        # --plain-port：SSL 模式下同时监听的明文端口（默认 1883，0 表示禁用）
        _plain_port = getattr(config, "MQTT_BROKER_PORT", 1883)
        if "--plain-port" in sys.argv:
            idx = sys.argv.index("--plain-port")
            _plain_port = int(sys.argv[idx + 1])
        _ssl_config["plain_port"] = _plain_port
        _ports_info = f"{_live_port}(SSL)" + (f" + {_plain_port}(明文)" if _plain_port else "")
        log.info("mTLS 已启用，Broker 监听端口：%s", _ports_info)

    # ── --all-modules（多设备合并报告，需与 --live 合用） ─────────────────────
    _all_modules = "--all-modules" in sys.argv

    if _all_modules:
        if not _live:
            print("[ERROR] --all-modules 必须与 --live 一起使用")
            sys.exit(1)
        asyncio.run(_main_all_modules(
            live_host=_live_host, live_port=_live_port,
            live_timeout=_live_timeout, live_topic=_live_topic,
            ssl_config=_ssl_config,
        ))
    else:
        asyncio.run(_main(
            _json_path, _module_index, _keys, _check_meta,
            live=_live, live_host=_live_host, live_port=_live_port,
            live_timeout=_live_timeout, live_topic=_live_topic,
            no_modbus=_no_modbus, ssl_config=_ssl_config,
        ))
