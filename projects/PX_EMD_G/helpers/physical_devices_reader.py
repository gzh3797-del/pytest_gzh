# -*- coding: utf-8 -*-
"""
physical_devices_reader.py — 从 AcuHMI-1-7 网关 REST API 动态发现下挂 Modbus TCP 电表。

复用一个**已登录且处于网关 origin** 的 Playwright Page（其 sessionStorage 含 authToken），
调用：
    GET /api/device/list/modbus?token=<token>     —— 设备清单（deviceType==3 为 Modbus TCP）
    GET /api/device/config/<serialNumber>?token=  —— 单台连接配置（modbusTCPConfig.*）
转成 DiscoveredDevice 列表，替代旧的 config.yaml device_modbus 静态配置。

设计见 docs/superpowers/specs/2026-06-18-hmi17-dynamic-device-discovery-design.md。
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Optional

from projects.PX_EMD_G.helpers.template_matcher import resolve_template_from_model
from projects.PX_EMD_G.settings import HMI_URL

log = logging.getLogger(__name__)

# list 接口里 Modbus 设备的 deviceType 值：2=Modbus RTU，3=Modbus TCP。
_DEVICE_TYPE_MODBUS_RTU = 2
_DEVICE_TYPE_MODBUS_TCP = 3
_MODBUS_DEVICE_TYPES = (_DEVICE_TYPE_MODBUS_RTU, _DEVICE_TYPE_MODBUS_TCP)


@dataclass(frozen=True)
class DiscoveredDevice:
    """网关下挂的一台 Modbus（TCP 或 RTU）设备的运行时发现结果。"""
    name: str               # deviceName（BACnet 下拉 / HMI 选择显示名）
    model: str              # deviceModel，如 "PXB-M24-XMA-GEN" / "PXE2"
    template: Optional[str]  # 解析出的模板名；未知型号为 None
    online: bool            # isOnline
    serial_id: str          # serialNumber（内部 MODTCP…/MODRTU… id，用于 config 接口）
    transport: str          # "tcp" | "rtu"
    ip: str                 # TCP 设备的 ipAddr；RTU 设备为 ""
    port: int               # TCP 设备的 port；RTU 设备为 0
    unit: int               # slaveAddress（TCP/RTU 均有）


def build_device(list_entry: dict, cfg_entry: dict) -> Optional[DiscoveredDevice]:
    """把一条 list 记录 + 对应 config 记录合成 DiscoveredDevice。

    TCP 设备从 modbusTCPConfig 取 ip/port/slaveAddress（缺任一字段则返回 None）；
    RTU 设备从 modbusRTUConfig 取 slaveAddress，ip/port 置空——RTU 走网关串口、无法
    网络直连 Modbus 比对，但按型号读模板的参数列表/COV 用例仍可用。两种配置都没有
    则返回 None（跳过该台）。
    """
    model = str(list_entry.get("deviceModel", ""))
    common: dict[str, Any] = dict(
        name=str(list_entry.get("deviceName", "")),
        model=model,
        template=resolve_template_from_model(model),
        online=bool(list_entry.get("isOnline", False)),
        serial_id=str(list_entry.get("serialNumber", "")),
    )
    tcp = cfg_entry.get("modbusTCPConfig")
    if isinstance(tcp, dict) and tcp:
        ip = tcp.get("ipAddr")
        port = tcp.get("port")
        unit = tcp.get("slaveAddress")
        if ip in (None, "") or port is None or unit is None:
            return None
        return DiscoveredDevice(
            **common, transport="tcp", ip=str(ip), port=int(port), unit=int(unit)
        )
    rtu = cfg_entry.get("modbusRTUConfig")
    if isinstance(rtu, dict) and rtu:
        unit = rtu.get("slaveAddress")
        return DiscoveredDevice(
            **common, transport="rtu", ip="", port=0,
            unit=int(unit) if unit is not None else 0,
        )
    return None


def connection_map(devices: list[DiscoveredDevice]) -> dict[str, tuple[str, int, int]]:
    """{deviceName: (ip, port, unit)}，仅含可直连的 Modbus TCP 设备。

    RTU 设备走网关侧串口、无 IP，无法做网络直连 Modbus 比对，故排除——protocol
    数值比对取不到其连接信息时按既有逻辑「跳过数值比对」（不报错，BACnet 读取仍跑）。
    """
    return {d.name: (d.ip, d.port, d.unit) for d in devices if d.transport == "tcp"}


def pick_device_for_template(
    devices: list[DiscoveredDevice], template: str, online_only: bool = True
) -> Optional[DiscoveredDevice]:
    """按模板名取一台匹配设备（默认仅在线）；无匹配返回 None。

    多台同型号时优先可直连的 TCP 设备、其次在线设备：protocol 数值比对需要 TCP 连接，
    而按型号读模板的参数列表/COV 用例任意一台等价。
    """
    candidates = [
        d for d in devices
        if d.template == template and (d.online or not online_only)
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda d: (d.transport != "tcp", not d.online))
    return candidates[0]


def _read_auth_token(page) -> str:
    """从页面 sessionStorage['common'].authToken 取鉴权 token。"""
    token = page.evaluate(
        "() => { const c = sessionStorage.getItem('common');"
        " return c ? (JSON.parse(c).authToken || null) : null; }"
    )
    if not token:
        raise RuntimeError(
            "未能从 sessionStorage['common'].authToken 取到网关鉴权 token——"
            "页面可能未登录或会话已失效。请确认发现 fixture 复用的是已登录页面。"
        )
    return str(token)


def _api_get_json(page, path: str, token: str) -> Any:
    """调网关 API（GET），返回解析后的 JSON；非 2xx 抛错并带 URL/状态码。"""
    sep = "&" if "?" in path else "?"
    url = f"{HMI_URL}{path}{sep}token={token}"
    resp = page.request.get(url)
    if not resp.ok:
        raise RuntimeError(f"网关 API 请求失败 {resp.status}：{url}")
    return resp.json()


def discover_modbus_devices(page) -> list[DiscoveredDevice]:
    """从网关 API 发现全部下挂 Modbus 设备（TCP + RTU，含离线，online 字段标识）。

    page 须为已登录、处于网关 origin 的 Playwright Page。
    """
    token = _read_auth_token(page)
    listing = _api_get_json(page, "/api/device/list/modbus", token)
    if not isinstance(listing, list) or not listing:
        raise RuntimeError(
            f"网关 /api/device/list/modbus 返回空或非列表：{type(listing).__name__}。"
            "请确认网关下挂了 Modbus 设备。"
        )
    devices: list[DiscoveredDevice] = []
    for entry in listing:
        if entry.get("deviceType") not in _MODBUS_DEVICE_TYPES:
            continue
        serial = entry.get("serialNumber")
        if not serial:
            continue
        cfg_entry = _api_get_json(page, f"/api/device/config/{serial}", token)
        dev = build_device(entry, cfg_entry)
        if dev is None:
            log.warning(
                "跳过设备 %r：config 缺 modbusTCPConfig / modbusRTUConfig 连接字段",
                entry.get("deviceName"),
            )
            continue
        devices.append(dev)
    return devices


# 向后兼容旧名（现含 TCP + RTU）。
discover_modbus_tcp_devices = discover_modbus_devices
